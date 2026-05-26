"""
surface_temperature_models_v3.py

Excelで作成した表面温度計算をPython化したスクリプトです。

含まれるモデル:
1) 各面独立モデル / Face_Calc 相当
   - 各面が独立して発熱し、その面だけから放熱する。

2) 均一温度モデル / Uniform_Body 相当
   - 筒壁全体が1つの表面温度を持つ。
   - 全面の合計放熱量 = 総発熱量 で解く。

3) 熱回路網モデル / Thermal_Network
   - 上面、下面、側面合計を別ノードとする。
   - 上面<->側面、下面<->側面を壁内熱伝導で接続する。
   - 各ノードは、対流+放射で外気へ放熱する。
   - 各面で表面発熱が一様に発生する。

特徴:
- 表面発熱が外表面に一様分布する中空筒/導波管モデル
- 端面からの放熱は無視
- 空気物性は膜温度 Tf=(Ts+Tinf)/2 で更新
- 自然対流 h(Ts) と放射 q_rad=eps*sigma*(Ts^4-Tsur^4) を温度依存で評価
- ExcelのGoal Seek相当は、二分法または減衰Newton法で解く
- scipy / pandas 不要。標準ライブラリのみで実行可能

使い方:
    python surface_temperature_models_v3.py

入力値を変えたい場合:
    ファイル上部の USER_* を変更してください。

correlation_mode:
    "low_ra_floor" : 推奨。水平面の低Rayleigh数範囲で Nu=1 とする。
    "continuous"   : Nu=max(1, 相関式) として連続的に使う。
                     Excelの元の連続相関式に近い。

熱回路網モデルの追加入力:
    USER_WALL_THICKNESS : 筒壁厚さ [m]
    USER_K_WALL         : 壁材料の熱伝導率 [W/(m K)]

熱回路網の考え方:
    face nodes = top, bottom, side

    P_top    = Qout_top    + G_top_side    (T_top - T_side)
    P_bottom = Qout_bottom + G_bottom_side (T_bottom - T_side)
    P_side   = Qout_side   + G_top_side    (T_side - T_top)
                         + G_bottom_side (T_side - T_bottom)

    Qout_i = A_i * [ h_i(T_i)(T_i-Tamb) + eps*sigma*(T_i^4-Tsur^4) ]

    top<->side の熱伝導コンダクタンスは、薄肉筒の周方向伝導近似として、

        G_top_side = 2 * k_wall * t_wall * L / ((a+b)/2)
                   = 4 * k_wall * t_wall * L / (a+b)

    としています。2は左右2つの上側コーナーを意味します。
    bottom<->side も同じです。
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

# Thermal network model parameters
USER_WALL_THICKNESS = 0.0010  # wall thickness [m], example: 1.0 mm
USER_K_WALL = 167.0           # wall thermal conductivity [W/(m K)], example: aluminum


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

    # Wall properties for thermal-network model
    wall_thickness: float = 0.0010  # wall thickness [m]
    k_wall: float = 167.0           # wall thermal conductivity [W/(m K)]

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


@dataclass
class NetworkNodeResult:
    face: str
    Ts_C: float
    A: float
    Q_gen: float
    Q_out: float
    Q_cond_out: float
    residual: float
    Ra: float
    Nu: float
    h_conv: float
    h_rad_eq: float
    q_conv: float
    q_rad: float


@dataclass
class NetworkResult:
    Ts_top_C: float
    Ts_bottom_C: float
    Ts_side_C: float
    G_top_side: float
    G_bottom_side: float
    Q_top_to_side: float
    Q_bottom_to_side: float
    nodes: List[NetworkNodeResult]
    residuals: Tuple[float, float, float]
    Q_total_out: float
    Q_total_gen: float


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
    """自然対流のNusselt数を計算する。"""
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


def solve_linear_system_3x3(A: List[List[float]], b: List[float]) -> List[float]:
    """3x3線形方程式 A x = b を部分ピボット付きGauss消去で解く。"""
    n = 3
    M = [row[:] + [rhs] for row, rhs in zip(A, b)]

    for col in range(n):
        pivot = max(range(col, n), key=lambda r: abs(M[r][col]))
        if abs(M[pivot][col]) < 1e-18:
            raise RuntimeError("singular matrix in Newton solver")
        if pivot != col:
            M[col], M[pivot] = M[pivot], M[col]

        piv = M[col][col]
        for j in range(col, n + 1):
            M[col][j] /= piv

        for r in range(n):
            if r == col:
                continue
            factor = M[r][col]
            for j in range(col, n + 1):
                M[r][j] -= factor * M[col][j]

    return [M[i][n] for i in range(n)]


def solve_nonlinear_3var(func, x0: List[float], *, tol: float = 1e-9, max_iter: int = 80) -> List[float]:
    """有限差分Jacobian + 減衰Newton法で3変数非線形方程式を解く。"""
    x = x0[:]

    def norm_inf(v: List[float]) -> float:
        return max(abs(vi) for vi in v)

    for _ in range(max_iter):
        f = func(x)
        if norm_inf(f) < tol:
            return x

        J = [[0.0] * 3 for _ in range(3)]
        for j in range(3):
            step = 1e-5 * max(1.0, abs(x[j]))
            xp = x[:]
            xp[j] += step
            fp = func(xp)
            for i in range(3):
                J[i][j] = (fp[i] - f[i]) / step

        # J dx = -f
        dx = solve_linear_system_3x3(J, [-fi for fi in f])

        # Damping: residualが悪化しない範囲で進める
        base_norm = norm_inf(f)
        lam = 1.0
        accepted = False
        for _ls in range(30):
            xn = [x[i] + lam * dx[i] for i in range(3)]
            # 極端な負温度に飛ばないようにする
            if min(xn) < -250.0:
                lam *= 0.5
                continue
            fn = func(xn)
            if norm_inf(fn) <= base_norm or lam < 1e-6:
                x = xn
                accepted = True
                break
            lam *= 0.5

        if not accepted:
            raise RuntimeError("Newton damping failed")

    f = func(x)
    raise RuntimeError(f"Newton solver did not converge: x={x}, residual={f}")


def solve_independent_faces(inp: Inputs, mode: CorrelationMode = "low_ra_floor") -> Dict[str, FaceResult]:
    """各面独立モデル。"""
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
    """均一温度モデル。"""
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


def thermal_network_conductance(inp: Inputs) -> Tuple[float, float]:
    """上面<->側面、下面<->側面の熱コンダクタンスを計算する。

    薄肉筒の周方向熱伝導として、上面中心から側面中心までの距離を
    (a+b)/2 と置く。上面と側面は左右2つのコーナーでつながるため、
    コンダクタンスは2本並列。
    """
    if inp.wall_thickness <= 0:
        raise ValueError("wall_thickness must be positive")
    if inp.k_wall <= 0:
        raise ValueError("k_wall must be positive")

    path_length = 0.5 * (inp.a + inp.b)
    G_one_corner = inp.k_wall * inp.wall_thickness * inp.L / path_length
    G_top_side = 2.0 * G_one_corner
    G_bottom_side = 2.0 * G_one_corner
    return G_top_side, G_bottom_side


def solve_thermal_network(inp: Inputs, mode: CorrelationMode = "low_ra_floor") -> NetworkResult:
    """熱回路網モデル。

    unknowns: [T_top_C, T_bottom_C, T_side_C]
    equations:
        Q_gen_i - Q_out_i - Q_cond_out_i = 0
    """
    geom = calc_geometry(inp)
    qpp = inp.P / geom.A_total
    Qgen_top = qpp * geom.A_top
    Qgen_bottom = qpp * geom.A_bottom
    Qgen_side = qpp * geom.A_side

    G_ts, G_bs = thermal_network_conductance(inp)

    def eval_faces(x: List[float]) -> Tuple[FaceResult, FaceResult, FaceResult]:
        Tt, Tb, Ts = x
        rt = face_heat_balance("top", Tt, geom.A_top, geom.Lc_horizontal, inp, mode)
        rb = face_heat_balance("bottom", Tb, geom.A_bottom, geom.Lc_horizontal, inp, mode)
        rs = face_heat_balance("side", Ts, geom.A_side, geom.Lc_vertical, inp, mode)
        return rt, rb, rs

    def residuals(x: List[float]) -> List[float]:
        rt, rb, rs = eval_faces(x)
        Tt, Tb, Ts = x
        Q_t_to_s = G_ts * (Tt - Ts)
        Q_b_to_s = G_bs * (Tb - Ts)
        # positive residual means heat generation is larger than outgoing heat
        f_top = Qgen_top - rt.Q_face - Q_t_to_s
        f_bottom = Qgen_bottom - rb.Q_face - Q_b_to_s
        f_side = Qgen_side - rs.Q_face + Q_t_to_s + Q_b_to_s
        return [f_top, f_bottom, f_side]

    # 初期値は均一温度解にする。収束しやすい。
    uniform_Ts_C, _, _ = solve_uniform_body(inp, mode)
    x = solve_nonlinear_3var(residuals, [uniform_Ts_C, uniform_Ts_C, uniform_Ts_C])

    rt, rb, rs = eval_faces(x)
    Tt, Tb, Ts = x
    Q_t_to_s = G_ts * (Tt - Ts)
    Q_b_to_s = G_bs * (Tb - Ts)
    f = residuals(x)

    nodes = [
        NetworkNodeResult(
            face="top", Ts_C=Tt, A=geom.A_top, Q_gen=Qgen_top, Q_out=rt.Q_face,
            Q_cond_out=Q_t_to_s, residual=f[0], Ra=rt.Ra, Nu=rt.Nu,
            h_conv=rt.h_conv, h_rad_eq=rt.h_rad_eq, q_conv=rt.q_conv, q_rad=rt.q_rad,
        ),
        NetworkNodeResult(
            face="bottom", Ts_C=Tb, A=geom.A_bottom, Q_gen=Qgen_bottom, Q_out=rb.Q_face,
            Q_cond_out=Q_b_to_s, residual=f[1], Ra=rb.Ra, Nu=rb.Nu,
            h_conv=rb.h_conv, h_rad_eq=rb.h_rad_eq, q_conv=rb.q_conv, q_rad=rb.q_rad,
        ),
        NetworkNodeResult(
            face="side", Ts_C=Ts, A=geom.A_side, Q_gen=Qgen_side, Q_out=rs.Q_face,
            Q_cond_out=-(Q_t_to_s + Q_b_to_s), residual=f[2], Ra=rs.Ra, Nu=rs.Nu,
            h_conv=rs.h_conv, h_rad_eq=rs.h_rad_eq, q_conv=rs.q_conv, q_rad=rs.q_rad,
        ),
    ]

    Q_total_out = rt.Q_face + rb.Q_face + rs.Q_face
    return NetworkResult(
        Ts_top_C=Tt,
        Ts_bottom_C=Tb,
        Ts_side_C=Ts,
        G_top_side=G_ts,
        G_bottom_side=G_bs,
        Q_top_to_side=Q_t_to_s,
        Q_bottom_to_side=Q_b_to_s,
        nodes=nodes,
        residuals=(f[0], f[1], f[2]),
        Q_total_out=Q_total_out,
        Q_total_gen=inp.P,
    )


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


def print_network(net: NetworkResult, inp: Inputs) -> None:
    print("\n[熱回路網モデル / Thermal_Network]")
    print(f"wall_thickness = {inp.wall_thickness:.6g} m, k_wall = {inp.k_wall:.6g} W/(m K)")
    print(f"G_top_side = {net.G_top_side:.6f} W/K, G_bottom_side = {net.G_bottom_side:.6f} W/K")
    print(f"Q_top_to_side = {net.Q_top_to_side:.6f} W, Q_bottom_to_side = {net.Q_bottom_to_side:.6f} W")
    print(f"Q_total_out = {net.Q_total_out:.9f} W, residual_total = {net.Q_total_gen - net.Q_total_out:.3e} W")
    print("face      Ts[C]      Q_gen    Q_out   Q_cond_out residual     Ra[-]      Nu[-]    h_conv   h_rad_eq")
    for n in net.nodes:
        print(
            f"{n.face:6s} "
            f"{n.Ts_C:10.6f} {n.Q_gen:9.4f} {n.Q_out:8.4f} {n.Q_cond_out:11.4f} {n.residual:9.2e} "
            f"{n.Ra:10.3f} {n.Nu:9.4f} {n.h_conv:9.4f} {n.h_rad_eq:9.4f}"
        )


def write_csv(
    path: Path,
    independent: Dict[str, FaceResult],
    uniform_results: List[FaceResult],
    uniform_Ts_C: float,
    uniform_Q_total: float,
    network: NetworkResult,
) -> None:
    """結果確認用CSVを出力する。"""
    headers = [
        "model", "face", "Ts_C", "A_m2", "Lc_m", "Tfilm_K", "Ra", "Nu",
        "h_conv_W_m2K", "h_rad_eq_W_m2K", "q_conv_W_m2", "q_rad_W_m2",
        "q_out_W_m2", "Q_face_or_Qout_W", "Q_gen_W", "Q_cond_out_W", "residual_W", "radiation_fraction",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for face in ["top", "bottom", "side"]:
            r = independent[face]
            writer.writerow([
                "independent", r.face, r.Ts_C, r.A, r.Lc, r.Tfilm_K, r.Ra, r.Nu,
                r.h_conv, r.h_rad_eq, r.q_conv, r.q_rad, r.q_out, r.Q_face,
                "", "", "", r.radiation_fraction,
            ])
        for r in uniform_results:
            writer.writerow([
                "uniform", r.face, uniform_Ts_C, r.A, r.Lc, r.Tfilm_K, r.Ra, r.Nu,
                r.h_conv, r.h_rad_eq, r.q_conv, r.q_rad, r.q_out, r.Q_face,
                "", "", "", r.radiation_fraction,
            ])
        writer.writerow(["uniform_total", "all", uniform_Ts_C, "", "", "", "", "", "", "", "", "", "", uniform_Q_total, "", "", "", ""])

        for n in network.nodes:
            writer.writerow([
                "thermal_network", n.face, n.Ts_C, n.A, "", "", n.Ra, n.Nu,
                n.h_conv, n.h_rad_eq, n.q_conv, n.q_rad, "", n.Q_out,
                n.Q_gen, n.Q_cond_out, n.residual, "",
            ])
        writer.writerow(["thermal_network_total", "all", "", "", "", "", "", "", "", "", "", "", "", network.Q_total_out, network.Q_total_gen, "", network.Q_total_gen - network.Q_total_out, ""])
        writer.writerow(["thermal_network_G", "top_side", "", "", "", "", "", "", "", "", "", "", "", "", "", network.G_top_side, network.Q_top_to_side, ""])
        writer.writerow(["thermal_network_G", "bottom_side", "", "", "", "", "", "", "", "", "", "", "", "", "", network.G_bottom_side, network.Q_bottom_to_side, ""])


def run_case(inp: Inputs, mode: CorrelationMode, csv_path: Path | None = None) -> None:
    print("=" * 80)
    print(f"correlation_mode = {mode}")
    print(
        f"a={inp.a} m, b={inp.b} m, L={inp.L} m, P={inp.P} W, "
        f"Tamb={inp.Tamb_C} degC, eps={inp.eps}"
    )

    independent = solve_independent_faces(inp, mode)
    print_independent(independent, inp)

    uniform_Ts_C, uniform_results, uniform_Q_total = solve_uniform_body(inp, mode)
    print_uniform(uniform_Ts_C, uniform_results, uniform_Q_total, inp)

    network = solve_thermal_network(inp, mode)
    print_network(network, inp)

    if csv_path is not None:
        write_csv(csv_path, independent, uniform_results, uniform_Ts_C, uniform_Q_total, network)
        print(f"\nCSV written: {csv_path}")


def main() -> None:
    inp = Inputs(
        a=USER_A,
        b=USER_B,
        L=USER_L,
        P=USER_P,
        Tamb_C=USER_TAMB_C,
        Tsur_C=USER_TSUR_C,
        eps=USER_EPS,
        wall_thickness=USER_WALL_THICKNESS,
        k_wall=USER_K_WALL,
    )

    # 推奨: 低Rayleigh数の水平面を Nu=1 に制限する修正版
    run_case(inp, mode="low_ra_floor", csv_path=Path("surface_temperature_results_low_ra_floor_v3.csv"))

    # 比較用: 連続相関式モード
    run_case(inp, mode="continuous", csv_path=Path("surface_temperature_results_continuous_v3.csv"))


if __name__ == "__main__":
    main()
