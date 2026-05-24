"""
surface_temperature_models.py

Excelで作成した「各面独立モデル」と「均一温度（結合）モデル」を
Pythonで計算するためのスクリプトです。

特徴:
- 表面発熱が外表面に一様分布する中空筒/導波管モデル
- 端面からの放熱は無視
- 空気物性は膜温度 Tf=(Ts+Tinf)/2 で更新
- 自然対流 h(Ts) と放射 q_rad=eps*sigma*(Ts^4-Tsur^4) を温度依存で評価
- ExcelのGoal Seek相当は、二分法で非線形方程式を解く
- scipy / pandas 不要。標準ライブラリのみで実行可能

使い方:
    python surface_temperature_models.py

入力値を変えたい場合:
    ファイル上部の USER_A, USER_B, USER_L, USER_P などを変更してください。

correlation_mode:
    "low_ra_floor" : 推奨。水平面の低Rayleigh数範囲で Nu=1 とする。
    "continuous"   : Nu=max(1, 相関式) として連続的に使う。
                     アップロードされた Uniform_Body シートの結果に近い。
"""

from __future__ import annotations

from dataclasses import dataclass
import csv
import math
from pathlib import Path
from typing import Dict, List, Literal, Tuple

CorrelationMode = Literal["low_ra_floor", "continuous"]
FaceName = Literal["top", "bottom", "side"]


# =============================================================================
# User input section
# =============================================================================
# ここだけ変更すれば、下側のクラス定義や関数を編集しなくても計算できます。
USER_A = 0.0254       # width [m]
USER_B = 0.0127       # height [m]
USER_L = 1.0          # length [m]
USER_P = 33.9         # total heat generation [W]
USER_TAMB_C = 25.0    # ambient air temperature [degC]
USER_TSUR_C = 25.0    # radiation surrounding temperature [degC]
USER_EPS = 0.9        # emissivity [-]


@dataclass
class Inputs:
    # Geometry / heat input
    a: float = 0.0254          # width [m]
    b: float = 0.0127          # height [m]
    L: float = 1.0             # length [m]
    P: float = 33.9            # total heat generation [W]
    Tamb_C: float = 25.0       # ambient air temperature [degC]
    Tsur_C: float = 25.0       # radiation surrounding temperature [degC]
    eps: float = 0.9           # emissivity [-]

    # Air / radiation constants
    p: float = 101325.0        # pressure [Pa]
    R_air: float = 287.058     # gas constant for dry air [J/(kg K)]
    mu0: float = 1.716e-5      # Sutherland reference viscosity [Pa s]
    T0: float = 273.15         # Sutherland reference temperature [K]
    S: float = 111.0           # Sutherland constant [K]
    k0: float = 0.0241         # air thermal conductivity at T0 [W/(m K)]
    k_exp: float = 0.9         # exponent for k_air = k0*(T/T0)^k_exp
    cp: float = 1007.0         # specific heat [J/(kg K)]
    sigma: float = 5.670374419e-8  # Stefan-Boltzmann constant [W/(m2 K4)]
    g: float = 9.80665         # gravity [m/s2]


@dataclass
class Geometry:
    A_top: float
    A_bottom: float
    A_side: float
    A_total: float
    Lc_horizontal: float
    Lc_vertical: float


@dataclass
class AirProps:
    rho: float
    mu: float
    k_air: float
    cp: float
    Pr: float
    nu: float
    alpha: float
    beta: float


@dataclass
class FaceResult:
    face: str
    Ts_C: float
    A: float
    Lc: float
    Tfilm_K: float
    rho: float
    mu: float
    k_air: float
    Pr: float
    nu: float
    alpha: float
    beta: float
    Ra: float
    Nu: float
    h_conv: float
    h_rad_eq: float
    q_conv: float
    q_rad: float
    q_out: float
    Q_face: float
    radiation_fraction: float


def calc_geometry(inp: Inputs) -> Geometry:
    """端面を除く外表面積と代表長さを計算する。"""
    A_top = inp.a * inp.L
    A_bottom = A_top
    A_side = 2.0 * inp.b * inp.L
    A_total = A_top + A_bottom + A_side

    # 水平長方形板の代表長さ: Lc = A / perimeter
    # ここでは上面/下面の平面寸法を a x L として perimeter=2(a+L)
    Lc_horizontal = A_top / (2.0 * (inp.a + inp.L))

    # 垂直平板の代表長さ: 高さ b
    Lc_vertical = inp.b

    return Geometry(A_top, A_bottom, A_side, A_total, Lc_horizontal, Lc_vertical)


def air_properties(Tfilm_K: float, inp: Inputs) -> AirProps:
    """膜温度で空気物性を計算する。Excelの簡易式に合わせている。"""
    rho = inp.p / (inp.R_air * Tfilm_K)
    mu = inp.mu0 * (Tfilm_K / inp.T0) ** 1.5 * (inp.T0 + inp.S) / (Tfilm_K + inp.S)
    k_air = inp.k0 * (Tfilm_K / inp.T0) ** inp.k_exp
    cp = inp.cp
    Pr = mu * cp / k_air
    nu = mu / rho
    alpha = k_air / (rho * cp)
    beta = 1.0 / Tfilm_K
    return AirProps(rho, mu, k_air, cp, Pr, nu, alpha, beta)


def nusselt_number(face: FaceName, Ra: float, Pr: float, mode: CorrelationMode) -> float:
    """自然対流のNusselt数を計算する。

    mode="low_ra_floor":
        水平面の一般的な適用範囲より低いRaでは Nu=1 とする。
        Excelを修正した推奨モデル。

    mode="continuous":
        Raの範囲外でも Nu=max(1, 相関式) として使う。
        アップロードされたUniform_Bodyシートの結果に近い。
    """
    Ra_abs = abs(Ra)

    if face == "top":
        if mode == "low_ra_floor":
            if Ra_abs < 1.0e4:
                return 1.0
            if Ra_abs <= 1.0e7:
                return max(1.0, 0.54 * Ra_abs ** 0.25)
            return max(1.0, 0.15 * Ra_abs ** (1.0 / 3.0))

        # continuous
        if Ra_abs <= 1.0e7:
            return max(1.0, 0.54 * Ra_abs ** 0.25)
        return max(1.0, 0.15 * Ra_abs ** (1.0 / 3.0))

    if face == "bottom":
        if mode == "low_ra_floor":
            if Ra_abs < 1.0e5:
                return 1.0
            return max(1.0, 0.27 * Ra_abs ** 0.25)

        # continuous
        return max(1.0, 0.27 * Ra_abs ** 0.25)

    if face == "side":
        # 垂直平板 Churchill-Chu 型の相関式
        return (
            0.825
            + 0.387 * Ra_abs ** (1.0 / 6.0)
            / (1.0 + (0.492 / Pr) ** (9.0 / 16.0)) ** (8.0 / 27.0)
        ) ** 2

    raise ValueError(f"unknown face: {face}")


def face_heat_balance(face: FaceName, Ts_C: float, A: float, Lc: float, inp: Inputs, mode: CorrelationMode) -> FaceResult:
    """指定した表面温度 Ts_C で、対流・放射・放熱量を計算する。"""
    Ts_K = Ts_C + 273.15
    Tamb_K = inp.Tamb_C + 273.15
    Tsur_K = inp.Tsur_C + 273.15
    dT = Ts_K - Tamb_K
    Tfilm_K = 0.5 * (Ts_K + Tamb_K)

    props = air_properties(Tfilm_K, inp)

    # 加熱問題が基本。もし dT<0 の場合でもRaの符号のみ保持し、Nu計算はabs(Ra)で行う。
    if props.nu <= 0 or props.alpha <= 0:
        raise ValueError("invalid air properties")
    Ra = inp.g * props.beta * dT * Lc ** 3 / (props.nu * props.alpha)
    Nu = nusselt_number(face, Ra, props.Pr, mode)
    h_conv = Nu * props.k_air / Lc

    q_conv = h_conv * dT
    q_rad = inp.eps * inp.sigma * (Ts_K ** 4 - Tsur_K ** 4)
    q_out = q_conv + q_rad
    Q_face = q_out * A

    h_rad_eq = inp.eps * inp.sigma * (Ts_K + Tsur_K) * (Ts_K ** 2 + Tsur_K ** 2)
    radiation_fraction = q_rad / q_out if abs(q_out) > 1e-30 else math.nan

    return FaceResult(
        face=face,
        Ts_C=Ts_C,
        A=A,
        Lc=Lc,
        Tfilm_K=Tfilm_K,
        rho=props.rho,
        mu=props.mu,
        k_air=props.k_air,
        Pr=props.Pr,
        nu=props.nu,
        alpha=props.alpha,
        beta=props.beta,
        Ra=Ra,
        Nu=Nu,
        h_conv=h_conv,
        h_rad_eq=h_rad_eq,
        q_conv=q_conv,
        q_rad=q_rad,
        q_out=q_out,
        Q_face=Q_face,
        radiation_fraction=radiation_fraction,
    )


def solve_bisection(func, lo: float, hi: float, *, tol: float = 1e-10, max_iter: int = 200) -> float:
    """Goal Seekの代わりに二分法で func(x)=0 を解く。"""
    f_lo = func(lo)
    f_hi = func(hi)

    # 上限が足りない場合は自動拡張
    expand_count = 0
    while f_lo * f_hi > 0 and expand_count < 100:
        width = hi - lo
        hi = hi + max(width, 10.0) * 2.0
        f_hi = func(hi)
        expand_count += 1

    if f_lo * f_hi > 0:
        raise RuntimeError(f"root is not bracketed: f({lo})={f_lo}, f({hi})={f_hi}")

    for _ in range(max_iter):
        mid = 0.5 * (lo + hi)
        f_mid = func(mid)
        if abs(f_mid) < tol or abs(hi - lo) < tol:
            return mid
        if f_lo * f_mid > 0:
            lo = mid
            f_lo = f_mid
        else:
            hi = mid
            f_hi = f_mid

    return 0.5 * (lo + hi)


def solve_independent_faces(inp: Inputs, mode: CorrelationMode = "low_ra_floor") -> Dict[str, FaceResult]:
    """各面独立モデル。

    各面で同じ表面発熱密度 q''=P/A_total が発生し、その面からのみ放熱すると仮定する。
    Excel Face_Calc シートの考え方。
    """
    geom = calc_geometry(inp)
    qpp = inp.P / geom.A_total

    face_specs: List[Tuple[FaceName, float, float]] = [
        ("top", geom.A_top, geom.Lc_horizontal),
        ("bottom", geom.A_bottom, geom.Lc_horizontal),
        ("side", geom.A_side, geom.Lc_vertical),
    ]

    results: Dict[str, FaceResult] = {}
    for face, A, Lc in face_specs:
        def residual(Ts_C: float) -> float:
            r = face_heat_balance(face, Ts_C, A, Lc, inp, mode)
            return qpp - r.q_out

        Ts_C = solve_bisection(residual, inp.Tamb_C, inp.Tamb_C + 300.0)
        results[face] = face_heat_balance(face, Ts_C, A, Lc, inp, mode)

    return results


def solve_uniform_body(inp: Inputs, mode: CorrelationMode = "low_ra_floor") -> Tuple[float, List[FaceResult], float]:
    """均一温度モデル。

    4面全体が1つの温度 Ts を持つとして、
    Σ Q_face(Ts) = P を満たす Ts を解く。
    Excel Uniform_Body シートの考え方。
    """
    geom = calc_geometry(inp)
    face_specs: List[Tuple[FaceName, float, float]] = [
        ("top", geom.A_top, geom.Lc_horizontal),
        ("bottom", geom.A_bottom, geom.Lc_horizontal),
        ("side", geom.A_side, geom.Lc_vertical),
    ]

    def total_Q(Ts_C: float) -> float:
        return sum(
            face_heat_balance(face, Ts_C, A, Lc, inp, mode).Q_face
            for face, A, Lc in face_specs
        )

    def residual(Ts_C: float) -> float:
        return inp.P - total_Q(Ts_C)

    Ts_C = solve_bisection(residual, inp.Tamb_C, inp.Tamb_C + 300.0)
    results = [face_heat_balance(face, Ts_C, A, Lc, inp, mode) for face, A, Lc in face_specs]
    Q_total = sum(r.Q_face for r in results)
    return Ts_C, results, Q_total


def print_independent(results: Dict[str, FaceResult], inp: Inputs) -> None:
    geom = calc_geometry(inp)
    qpp = inp.P / geom.A_total
    print("\n[各面独立モデル / Face_Calc相当]")
    print(f"q'' = {qpp:.6f} W/m2")
    print("face      Ts[C]      Ra[-]      Nu[-]    h_conv   h_rad_eq  q_conv   q_rad    Q_face")
    for face in ["top", "bottom", "side"]:
        r = results[face]
        print(
            f"{face:6s} "
            f"{r.Ts_C:10.6f} {r.Ra:10.3f} {r.Nu:9.4f} "
            f"{r.h_conv:9.4f} {r.h_rad_eq:9.4f} "
            f"{r.q_conv:8.3f} {r.q_rad:8.3f} {r.Q_face:9.4f}"
        )


def print_uniform(Ts_C: float, results: List[FaceResult], Q_total: float, inp: Inputs) -> None:
    print("\n[均一温度モデル / Uniform_Body相当]")
    print(f"Ts = {Ts_C:.6f} degC")
    print(f"Q_total = {Q_total:.9f} W, residual = {inp.P - Q_total:.3e} W")
    print("face      Ra[-]      Nu[-]    h_conv   h_rad_eq  q_conv   q_rad    Q_face")
    for r in results:
        print(
            f"{r.face:6s} "
            f"{r.Ra:10.3f} {r.Nu:9.4f} "
            f"{r.h_conv:9.4f} {r.h_rad_eq:9.4f} "
            f"{r.q_conv:8.3f} {r.q_rad:8.3f} {r.Q_face:9.4f}"
        )


def write_csv(path: Path, independent: Dict[str, FaceResult], uniform_results: List[FaceResult], uniform_Ts_C: float, uniform_Q_total: float) -> None:
    """結果確認用CSVを出力する。"""
    headers = [
        "model", "face", "Ts_C", "A_m2", "Lc_m", "Tfilm_K", "Ra", "Nu",
        "h_conv_W_m2K", "h_rad_eq_W_m2K", "q_conv_W_m2", "q_rad_W_m2",
        "q_out_W_m2", "Q_face_W", "radiation_fraction",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for face in ["top", "bottom", "side"]:
            r = independent[face]
            writer.writerow([
                "independent", r.face, r.Ts_C, r.A, r.Lc, r.Tfilm_K, r.Ra, r.Nu,
                r.h_conv, r.h_rad_eq, r.q_conv, r.q_rad, r.q_out, r.Q_face, r.radiation_fraction,
            ])
        for r in uniform_results:
            writer.writerow([
                "uniform", r.face, uniform_Ts_C, r.A, r.Lc, r.Tfilm_K, r.Ra, r.Nu,
                r.h_conv, r.h_rad_eq, r.q_conv, r.q_rad, r.q_out, r.Q_face, r.radiation_fraction,
            ])
        writer.writerow(["uniform_total", "all", uniform_Ts_C, "", "", "", "", "", "", "", "", "", "", uniform_Q_total, ""])


def run_case(inp: Inputs, mode: CorrelationMode, csv_path: Path | None = None) -> None:
    print("=" * 80)
    print(f"correlation_mode = {mode}")
    print(f"a={inp.a} m, b={inp.b} m, L={inp.L} m, P={inp.P} W, Tamb={inp.Tamb_C} degC, eps={inp.eps}")

    independent = solve_independent_faces(inp, mode)
    print_independent(independent, inp)

    uniform_Ts_C, uniform_results, uniform_Q_total = solve_uniform_body(inp, mode)
    print_uniform(uniform_Ts_C, uniform_results, uniform_Q_total, inp)

    if csv_path is not None:
        write_csv(csv_path, independent, uniform_results, uniform_Ts_C, uniform_Q_total)
        print(f"\nCSV written: {csv_path}")


def main() -> None:
    # WR-90相当の例。
    # ユーザーがアップロードしたExcelの入力値に合わせています。
    # 上部の USER_* 値から入力条件を作成します。
    # NameError: name 'Inputs' is not defined が出る場合は、
    # このファイル全体を実行してください。入力部だけを単独実行するとエラーになります。
    inp = Inputs(
        a=USER_A,
        b=USER_B,
        L=USER_L,
        P=USER_P,
        Tamb_C=USER_TAMB_C,
        Tsur_C=USER_TSUR_C,
        eps=USER_EPS,
    )

    # 推奨: 低Rayleigh数の水平面を Nu=1 に制限する修正版
    run_case(inp, mode="low_ra_floor", csv_path=Path("surface_temperature_results_low_ra_floor.csv"))

    # 比較用: Uniform_Bodyシートの現状に近い、連続相関式モード
    run_case(inp, mode="continuous", csv_path=Path("surface_temperature_results_continuous.csv"))


if __name__ == "__main__":
    main()
