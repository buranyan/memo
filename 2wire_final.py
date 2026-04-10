"""
2芯シールド線断面の静電容量行列を通常FDMで求めるプログラム
====================================================

【目的】
本コードは、2芯シールド線の2次元断面に対して静電ポテンシャル分布を
有限差分法（FDM: Finite Difference Method）で求め、そこから

    1. 線間静電容量
    2. 1芯-シールド間静電容量
    3. シールドを含む容量行列

を単位長さ当たり [F/m] で計算することを目的とする。

【解析対象】
断面形状は以下とする。

    ・芯線半径                  : r1
    ・2芯の中心間距離           : d
    ・シールド内半径            : r2
    ・芯線-シールド間の媒質     : 一様誘電体（比誘電率 eps_r）

2本の芯線は x 軸上に対称配置され、
中心座標は (+d/2, 0), (-d/2, 0) とする。
外周シールドは原点中心・半径 r2 の円とする。

【支配方程式】
誘電体領域では静電場の支配方程式としてラプラス方程式

    ∇²φ = 0

を解く。

ここで φ は静電ポテンシャルである。
芯線およびシールドは理想導体とし、各導体表面ではディリクレ境界条件
（既知電位条件）を与える。

【数値解法】
計算領域を正方格子に分割し、誘電体領域内の節点に対して
5点差分による有限差分近似を用いる。

各励振ケースについて疎行列連立方程式を構築し、SciPy の spsolve により
ポテンシャル分布を求める。

【容量行列の求め方】
シールドを 0 V に固定し、以下の2ケースを計算する。

    Case 1: core1 = 1 V, core2 = 0 V, shield = 0 V
    Case 2: core1 = 0 V, core2 = 1 V, shield = 0 V

各ケースで得られた芯線電荷から、縮約容量行列

    [q1]   [C11  C12] [V1]
    [q2] = [C21  C22] [V2]

を求める。

さらに、この 2×2 行列からシールドを含む 3×3 の Maxwell 容量行列を
構成する。

【電荷評価】
各導体の単位長さ当たり電荷 q' [C/m] は、導体表面に隣接する誘電体格子点との
電位差から法線方向電束を近似して求める。

1本の格子辺について

    dq' ≈ ε (Vc - Vn)

とし、導体全周にわたって総和を取る。

ここで
    ε = ε0 * eps_r
    Vc : 導体電位
    Vn : 隣接誘電体節点の電位
である。

【エネルギー法による整合確認】
静電エネルギー密度に基づく離散エネルギー

    W' ≈ (1/2) ε Σ(Δφ)^2

も同時に計算し、電荷法で得られた容量との一致を確認する。
これにより、離散化と後処理の内部整合性をチェックする。

【収束確認】
格子数 nx を段階的に増やした規則的細密格子列に対して計算を行い、

    ・容量値の変化
    ・対称性誤差
    ・残差
    ・Richardson 外挿による収束値推定

を確認する。

本コードでは、最終的な採用値として
規則的格子列に対する Richardson 外挿値を用いる。

【出力】
本コードは以下を出力する。

    ・縮約容量行列（2×2）
    ・Maxwell 容量行列（3×3）
    ・線間静電容量
    ・1芯-シールド間静電容量
    ・エネルギー法との比較結果
    ・格子収束結果
    ・Richardson 外挿による推定収束値

【注意】
本コードは2次元断面の静電場解析であり、ケーブル端部の3次元効果は含まない。
また、円境界は直交格子で近似しているため、厳密解ではなく数値近似解である。
ただし、規則格子列での収束確認と Richardson 外挿により、
実用上十分な精度で容量を評価できる。
"""

import numpy as np
from scipy.sparse import csr_matrix
from scipy.sparse.linalg import spsolve

EPS0 = 8.8541878128e-12  # vacuum permittivity [F/m]
TINY = 1e-30


def build_geometry(r1, d, r2, nx):
    """
    Node-centered grid on [-r2, r2] x [-r2, r2].

    Parameters
    ----------
    r1 : float
        Core radius [m]
    d : float
        Center-to-center spacing between the two cores [m]
    r2 : float
        Shield inner radius [m]
    nx : int
        Number of grid points in x and y (odd number)

    Returns
    -------
    geom : dict
    """
    if not (r1 > 0 and d > 0 and r2 > 0):
        raise ValueError("r1, d, r2 must be positive.")
    if d <= 2.0 * r1:
        raise ValueError("Require d > 2*r1.")
    if d / 2.0 + r1 >= r2:
        raise ValueError("Require d/2 + r1 < r2.")
    if nx < 51:
        raise ValueError("Use nx >= 51.")
    if nx % 2 == 0:
        raise ValueError("Use an odd nx.")

    x = np.linspace(-r2, r2, nx)
    h = x[1] - x[0]
    X, Y = np.meshgrid(x, x, indexing="xy")

    shield_inner = (X**2 + Y**2) <= r2**2
    core1 = ((X - d / 2.0)**2 + Y**2) <= r1**2
    core2 = ((X + d / 2.0)**2 + Y**2) <= r1**2
    dielectric = shield_inner & ~(core1 | core2)

    return {
        "x": x,
        "X": X,
        "Y": Y,
        "h": h,
        "nx": nx,
        "r1": r1,
        "d": d,
        "r2": r2,
        "shield_inner": shield_inner,
        "core1": core1,
        "core2": core2,
        "dielectric": dielectric,
    }


def assemble_system(geom, V1, V2, Vshield=0.0):
    """
    Assemble sparse linear system for Laplace equation:
        ∇²φ = 0
    in dielectric region with Dirichlet boundary conditions.
    """
    dielectric = geom["dielectric"]
    core1 = geom["core1"]
    core2 = geom["core2"]
    ny, nx = dielectric.shape

    ids = -np.ones((ny, nx), dtype=int)
    pts = np.argwhere(dielectric)
    n_unknown = len(pts)
    ids[dielectric] = np.arange(n_unknown)

    rows, cols, data = [], [], []
    b = np.zeros(n_unknown, dtype=float)

    for row_id, (i, j) in enumerate(pts):
        rows.append(row_id)
        cols.append(row_id)
        data.append(4.0)

        for di, dj in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
            ii, jj = i + di, j + dj

            if 0 <= ii < ny and 0 <= jj < nx:
                if dielectric[ii, jj]:
                    rows.append(row_id)
                    cols.append(ids[ii, jj])
                    data.append(-1.0)
                elif core1[ii, jj]:
                    b[row_id] += V1
                elif core2[ii, jj]:
                    b[row_id] += V2
                else:
                    b[row_id] += Vshield
            else:
                b[row_id] += Vshield

    A = csr_matrix((data, (rows, cols)), shape=(n_unknown, n_unknown))
    return A, b


def solve_potential(geom, V1, V2, Vshield=0.0):
    """
    Solve potential distribution.

    Returns
    -------
    phi : ndarray
        Potential on the whole grid
    rel_res : float
        Relative residual norm
    """
    A, b = assemble_system(geom, V1, V2, Vshield)
    phi_unknown = spsolve(A, b)

    r = A @ phi_unknown - b
    rel_res = np.linalg.norm(r) / max(np.linalg.norm(b), TINY)

    phi = np.full_like(geom["X"], Vshield, dtype=float)
    phi[geom["dielectric"]] = phi_unknown
    phi[geom["core1"]] = V1
    phi[geom["core2"]] = V2
    return phi, rel_res


def conductor_charge_per_length(phi, conductor_mask, dielectric_mask, Vc, eps):
    """
    Charge per unit length [C/m] on one conductor from discrete normal flux.

    For one conductor-dielectric grid edge:
        dq' ≈ eps * (Vc - V_neighbor)
    because
        E_n ≈ (Vc - V_neighbor) / h
        ds = h
    """
    q = 0.0
    ny, nx = phi.shape

    for i, j in np.argwhere(conductor_mask):
        for di, dj in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
            ii, jj = i + di, j + dj
            if 0 <= ii < ny and 0 <= jj < nx and dielectric_mask[ii, jj]:
                q += eps * (Vc - phi[ii, jj])

    return q


def discrete_energy_per_length(phi, geom, eps_r):
    """
    Discrete electrostatic energy per unit length [J/m].

    Consistent with 5-point FDM graph energy:
        W' ≈ 1/2 * eps * Σ_edges (Δφ)^2
    """
    eps = EPS0 * eps_r
    dielectric = geom["dielectric"]

    dphi_x = phi[:, 1:] - phi[:, :-1]
    active_x = dielectric[:, 1:] | dielectric[:, :-1]

    dphi_y = phi[1:, :] - phi[:-1, :]
    active_y = dielectric[1:, :] | dielectric[:-1, :]

    Wp = 0.5 * eps * (
        np.sum((dphi_x[active_x])**2) +
        np.sum((dphi_y[active_y])**2)
    )
    return Wp


def solve_case(geom, eps_r, V1, V2, Vshield=0.0):
    """
    Solve one excitation case and evaluate charges, energy, and diagnostics.
    """
    eps = EPS0 * eps_r
    phi, rel_res = solve_potential(geom, V1, V2, Vshield)

    q1 = conductor_charge_per_length(phi, geom["core1"], geom["dielectric"], V1, eps)
    q2 = conductor_charge_per_length(phi, geom["core2"], geom["dielectric"], V2, eps)
    qs = -(q1 + q2)

    Wp = discrete_energy_per_length(phi, geom, eps_r)
    neutrality_rel = abs(q1 + q2 + qs) / max(abs(q1) + abs(q2) + abs(qs), TINY)

    return {
        "phi": phi,
        "q_per_m": np.array([q1, q2, qs], dtype=float),
        "energy_per_m": Wp,
        "rel_residual": rel_res,
        "neutrality_rel": neutrality_rel,
    }


def compute_capacitance_matrices(r1, d, r2, eps_r, nx=301, symmetrize=True):
    """
    Compute reduced 2x2 capacitance matrix (shield grounded)
    and full 3x3 Maxwell capacitance matrix.
    """
    geom = build_geometry(r1=r1, d=d, r2=r2, nx=nx)

    case10 = solve_case(geom, eps_r=eps_r, V1=1.0, V2=0.0, Vshield=0.0)
    case01 = solve_case(geom, eps_r=eps_r, V1=0.0, V2=1.0, Vshield=0.0)

    C_reduced = np.array([
        [case10["q_per_m"][0], case01["q_per_m"][0]],
        [case10["q_per_m"][1], case01["q_per_m"][1]],
    ], dtype=float)

    if symmetrize:
        C_reduced = 0.5 * (C_reduced + C_reduced.T)

    a, b = C_reduced[0, 0], C_reduced[0, 1]
    c, d_ = C_reduced[1, 0], C_reduced[1, 1]

    C_full = np.array([
        [a,       b,       -(a + b)],
        [c,       d_,      -(c + d_)],
        [-(a+c), -(b+d_),   (a + b + c + d_)],
    ], dtype=float)

    return geom, C_reduced, C_full, case10, case01


def extracted_capacitances(C_reduced):
    """
    Derive practical capacitances from reduced 2x2 matrix.
    """
    C11, C12 = C_reduced[0, 0], C_reduced[0, 1]
    C21, C22 = C_reduced[1, 0], C_reduced[1, 1]

    C_line_line = (C11 - C12 - C21 + C22) / 4.0
    C_core1_shield_other_grounded = C11
    C_core1_shield_other_floating = C11 - (C12 * C21) / max(C22, TINY)

    return {
        "C_line_line_per_m": C_line_line,
        "C_core1_shield_other_grounded_per_m": C_core1_shield_other_grounded,
        "C_core1_shield_other_floating_per_m": C_core1_shield_other_floating,
    }


def richardson_extrapolation_three(h1, c1, h2, c2, h3, c3, p_min=0.05, p_max=8.0, nscan=4000):
    """
    Three-grid Richardson extrapolation for
        C(h) = C_inf + a * h^p
    with arbitrary h1 > h2 > h3.
    """
    d12 = c1 - c2
    d23 = c2 - c3

    if abs(d12) < TINY or abs(d23) < TINY:
        return {"success": False, "status": "differences too small"}
    if d12 * d23 <= 0:
        return {"success": False, "status": "non-monotone"}

    data_ratio = d12 / d23

    def model_ratio(p):
        return (h**p for h in (h1, h2, h3))

    def ratio_value(p):
        hp1, hp2, hp3 = model_ratio(p)
        return (hp1 - hp2) / max(hp2 - hp3, TINY)

    ps = np.linspace(p_min, p_max, nscan)
    vals = np.array([ratio_value(p) - data_ratio for p in ps])

    idx = np.where(np.sign(vals[:-1]) * np.sign(vals[1:]) <= 0)[0]
    if len(idx) > 0:
        k = idx[0]
        lo, hi = ps[k], ps[k + 1]
        flo = vals[k]
        for _ in range(80):
            mid = 0.5 * (lo + hi)
            fmid = ratio_value(mid) - data_ratio
            if flo * fmid <= 0:
                hi = mid
            else:
                lo = mid
                flo = fmid
        p_est = 0.5 * (lo + hi)
    else:
        p_est = ps[np.argmin(np.abs(vals))]

    H = np.array([
        [1.0, h1**p_est],
        [1.0, h2**p_est],
        [1.0, h3**p_est],
    ], dtype=float)
    y = np.array([c1, c2, c3], dtype=float)

    coeff, _, _, _ = np.linalg.lstsq(H, y, rcond=None)
    c_inf, a_est = coeff
    fit = H @ coeff

    fit_rel_max = np.max(np.abs(fit - y)) / max(abs(c_inf), TINY)
    fine_rel_error = abs(c_inf - c3) / max(abs(c_inf), TINY)

    return {
        "success": True,
        "p": p_est,
        "c_inf": c_inf,
        "a": a_est,
        "fit_rel_max": fit_rel_max,
        "fine_rel_error": fine_rel_error,
        "status": "ok",
    }


def convergence_study(r1, d, r2, eps_r, nx_list, symmetrize=True):
    """
    Run convergence study on multiple grid sizes.
    """
    rows = []
    prev_C11 = None
    prev_Cll = None

    for nx in nx_list:
        geom, C_reduced, _, case10, case01 = compute_capacitance_matrices(
            r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=symmetrize
        )

        case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0)
        caps = extracted_capacitances(C_reduced)

        C11_charge = C_reduced[0, 0]
        C22_charge = C_reduced[1, 1]
        C12_charge = C_reduced[0, 1]
        C21_charge = C_reduced[1, 0]
        Cll_charge = caps["C_line_line_per_m"]

        C11_energy = 2.0 * case10["energy_per_m"]
        Cll_energy = 2.0 * case_diff["energy_per_m"]

        rel_C11_charge_vs_energy = abs(C11_charge - C11_energy) / max(abs(C11_charge), TINY)
        rel_Cll_charge_vs_energy = abs(Cll_charge - Cll_energy) / max(abs(Cll_charge), TINY)
        rel_sym_diag = abs(C11_charge - C22_charge) / max(max(abs(C11_charge), abs(C22_charge)), TINY)
        rel_sym_off = abs(C12_charge - C21_charge) / max(max(abs(C12_charge), abs(C21_charge)), TINY)

        if prev_C11 is None:
            rel_C11_vs_prev = np.nan
            rel_Cll_vs_prev = np.nan
        else:
            rel_C11_vs_prev = abs(C11_charge - prev_C11) / max(abs(C11_charge), TINY)
            rel_Cll_vs_prev = abs(Cll_charge - prev_Cll) / max(abs(Cll_charge), TINY)

        prev_C11 = C11_charge
        prev_Cll = Cll_charge

        rows.append({
            "nx": nx,
            "h": geom["h"],
            "C11_charge": C11_charge,
            "C11_energy": C11_energy,
            "Cll_charge": Cll_charge,
            "Cll_energy": Cll_energy,
            "rel_C11_charge_vs_energy": rel_C11_charge_vs_energy,
            "rel_Cll_charge_vs_energy": rel_Cll_charge_vs_energy,
            "rel_C11_vs_prev": rel_C11_vs_prev,
            "rel_Cll_vs_prev": rel_Cll_vs_prev,
            "rel_sym_diag": rel_sym_diag,
            "rel_sym_off": rel_sym_off,
            "neutrality10": case10["neutrality_rel"],
            "neutrality01": case01["neutrality_rel"],
            "neutralityDiff": case_diff["neutrality_rel"],
            "res10": case10["rel_residual"],
            "res01": case01["rel_residual"],
            "resDiff": case_diff["rel_residual"],
            "rich_C11": None,
            "rich_Cll": None,
        })

    for i in range(2, len(rows)):
        a, b, c = rows[i - 2], rows[i - 1], rows[i]
        rows[i]["rich_C11"] = richardson_extrapolation_three(
            a["h"], a["C11_charge"],
            b["h"], b["C11_charge"],
            c["h"], c["C11_charge"]
        )
        rows[i]["rich_Cll"] = richardson_extrapolation_three(
            a["h"], a["Cll_charge"],
            b["h"], b["Cll_charge"],
            c["h"], c["Cll_charge"]
        )

    return rows


def pf_per_m(x):
    return 1e12 * x


def pct(x):
    return 100.0 * x


def print_single_result(C_reduced, C_full, case10, geom, eps_r):
    """
    Print single-grid result and energy check.
    """
    np.set_printoptions(precision=6, suppress=False)

    print("Reduced capacitance matrix with shield grounded [F/m]:")
    print(C_reduced)
    print()

    print("Full 3x3 Maxwell capacitance matrix [F/m]:")
    print(C_full)
    print()

    vals = extracted_capacitances(C_reduced)
    print("Derived capacitances [F/m]:")
    for k, v in vals.items():
        print(f"  {k:45s} = {v:.6e}")

    case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0)

    C11_charge = C_reduced[0, 0]
    C11_energy = 2.0 * case10["energy_per_m"]
    Cll_charge = vals["C_line_line_per_m"]
    Cll_energy = 2.0 * case_diff["energy_per_m"]

    print("\nSingle-grid energy check:")
    print(f"  C11 charge-based      = {C11_charge:.6e} F/m  ({pf_per_m(C11_charge):.6f} pF/m)")
    print(f"  C11 energy-based      = {C11_energy:.6e} F/m  ({pf_per_m(C11_energy):.6f} pF/m)")
    print(f"  relative difference   = {pct(abs(C11_charge - C11_energy)/max(abs(C11_charge), TINY)):.6f} %")
    print()
    print(f"  Cline charge-based    = {Cll_charge:.6e} F/m  ({pf_per_m(Cll_charge):.6f} pF/m)")
    print(f"  Cline energy-based    = {Cll_energy:.6e} F/m  ({pf_per_m(Cll_energy):.6f} pF/m)")
    print(f"  relative difference   = {pct(abs(Cll_charge - Cll_energy)/max(abs(Cll_charge), TINY)):.6f} %")


def print_convergence_summary(rows):
    """
    Print concise convergence and final Richardson estimate.
    """
    print("\n=== Convergence summary ===")
    header = (
        f"{'nx':>6} "
        f"{'C11 [pF/m]':>14} {'Cline [pF/m]':>14} "
        f"{'Δprev C11 [%]':>14} {'Δprev Cline [%]':>16} "
        f"{'sym diag [%]':>12} {'res10':>12}"
    )
    print(header)
    print("-" * len(header))

    for r in rows:
        s1 = "---" if np.isnan(r["rel_C11_vs_prev"]) else f"{pct(r['rel_C11_vs_prev']):.6f}"
        s2 = "---" if np.isnan(r["rel_Cll_vs_prev"]) else f"{pct(r['rel_Cll_vs_prev']):.6f}"

        print(
            f"{r['nx']:6d} "
            f"{pf_per_m(r['C11_charge']):14.6f} {pf_per_m(r['Cll_charge']):14.6f} "
            f"{s1:>14} {s2:>16} "
            f"{pct(r['rel_sym_diag']):12.6f} {r['res10']:12.3e}"
        )

    print("\n=== Final Richardson estimate ===")
    if len(rows) < 3:
        print("Not enough grid sizes.")
        return

    last = rows[-1]
    rc11 = last["rich_C11"]
    rcll = last["rich_Cll"]

    print(f"Using nx = {rows[-3]['nx']}, {rows[-2]['nx']}, {rows[-1]['nx']}")

    if rc11 is not None and rc11["success"]:
        print(f"  C11_inf estimate       = {rc11['c_inf']:.6e} F/m  ({pf_per_m(rc11['c_inf']):.6f} pF/m)")
        print(f"  estimated order p      = {rc11['p']:.6f}")
        print(f"  finest-grid rel. error = {pct(rc11['fine_rel_error']):.6f} %")
    else:
        print("  C11_inf estimate       = not available")

    if rcll is not None and rcll["success"]:
        print(f"  Cline_inf estimate     = {rcll['c_inf']:.6e} F/m  ({pf_per_m(rcll['c_inf']):.6f} pF/m)")
        print(f"  estimated order p      = {rcll['p']:.6f}")
        print(f"  finest-grid rel. error = {pct(rcll['fine_rel_error']):.6f} %")
    else:
        print("  Cline_inf estimate     = not available")


def main():
    # Geometry / material
    r1 = 0.50e-3   # [m]
    d = 2.00e-3    # [m]
    r2 = 2.00e-3   # [m]
    eps_r = 2.20

    # Single-grid result
    nx_single = 301
    geom, C_reduced, C_full, case10, _ = compute_capacitance_matrices(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx_single, symmetrize=True
    )
    print_single_result(C_reduced, C_full, case10, geom, eps_r)

    # Regular refined grid sequence for final estimate
    nx_list = [301, 601, 1201]
    rows = convergence_study(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx_list=nx_list, symmetrize=True
    )
    print_convergence_summary(rows)


if __name__ == "__main__":
    main()