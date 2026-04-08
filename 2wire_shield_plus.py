import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
from scipy.sparse import csr_matrix
from scipy.sparse.linalg import spsolve

EPS0 = 8.8541878128e-12  # vacuum permittivity [F/m]
TINY = 1e-30


def build_geometry(r1, d, r2, nx):
    """
    Geometry:
      - outer circular shield (inner radius r2)
      - two identical circular cores (radius r1)
      - core centers at x = +/- d/2, y = 0
    """
    if not (r1 > 0 and d > 0 and r2 > 0):
        raise ValueError("r1, d, r2 must be positive.")
    if d <= 2 * r1:
        raise ValueError("The two cores overlap: require d > 2*r1.")
    if d / 2 + r1 >= r2:
        raise ValueError("The cores touch or penetrate the shield: require d/2 + r1 < r2.")
    if nx < 51:
        raise ValueError("nx is too small. Use at least 51, preferably 201 or more.")
    if nx % 2 == 0:
        raise ValueError("Use an odd nx so the geometry stays centered on the grid.")

    x = np.linspace(-r2, r2, nx)
    h = x[1] - x[0]
    X, Y = np.meshgrid(x, x, indexing="xy")

    shield_inner = (X**2 + Y**2) <= r2**2
    core1 = ((X - d / 2)**2 + Y**2) <= r1**2
    core2 = ((X + d / 2)**2 + Y**2) <= r1**2
    dielectric = shield_inner & ~(core1 | core2)

    return {
        "x": x,
        "h": h,
        "X": X,
        "Y": Y,
        "shield_inner": shield_inner,
        "core1": core1,
        "core2": core2,
        "dielectric": dielectric,
        "nx": nx,
        "r1": r1,
        "d": d,
        "r2": r2,
    }


def assemble_system(geom, V1, V2, Vshield=0.0):
    """
    Assemble the sparse linear system for Laplace equation:
        ∇²φ = 0
    in the dielectric region, with Dirichlet BCs on the conductors.
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
                    # shield or outside dielectric
                    b[row_id] += Vshield
            else:
                b[row_id] += Vshield

    A = csr_matrix((data, (rows, cols)), shape=(n_unknown, n_unknown))
    return A, b


def solve_potential(geom, V1, V2, Vshield=0.0):
    """
    Solve the potential distribution for one excitation case.
    Returns phi and relative residual.
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
    Estimate charge per unit length [C/m] on one conductor by summing normal
    electric flux across conductor-dielectric grid edges.

    On a uniform square grid, for one edge:
        dq' ≈ eps * (Vc - V_neighbor)
    because E_n ≈ (Vc - V_neighbor)/h and ds = h.
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

    Consistent with the 5-point FDM graph-energy:
        W' ≈ (1/2) * eps * Σ_edges (Δφ)^2
    where each neighboring edge is counted once and only edges touching the
    dielectric are included.
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
    Solve one excitation case and return conductor charges, energy, and diagnostics.
    """
    eps = EPS0 * eps_r
    phi, rel_res = solve_potential(geom, V1, V2, Vshield)

    q1 = conductor_charge_per_length(phi, geom["core1"], geom["dielectric"], V1, eps)
    q2 = conductor_charge_per_length(phi, geom["core2"], geom["dielectric"], V2, eps)
    qs = -(q1 + q2)  # charge neutrality in the closed shielded system

    Wp = discrete_energy_per_length(phi, geom, eps_r)
    neutrality_rel = abs(q1 + q2 + qs) / max(abs(q1) + abs(q2) + abs(qs), TINY)

    return {
        "phi": phi,
        "q_per_m": np.array([q1, q2, qs], dtype=float),
        "energy_per_m": Wp,
        "rel_residual": rel_res,
        "neutrality_rel": neutrality_rel,
    }


def compute_capacitance_matrices(r1, d, r2, eps_r, nx=401, symmetrize=True):
    """
    Compute:
      1) reduced 2x2 capacitance matrix with shield grounded
      2) full 3x3 Maxwell capacitance matrix including the shield
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
    Derive practical capacitances from the reduced 2x2 matrix.
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


def pretty_print_results(C_reduced, C_full):
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


def plot_potential_contours(geom, phi, title="Potential distribution", nfill=40, nline=16):
    """
    Plot filled contours and contour lines of potential distribution.
    """
    X = geom["X"] * 1e3  # convert to mm for display
    Y = geom["Y"] * 1e3
    r1 = geom["r1"] * 1e3
    r2 = geom["r2"] * 1e3
    d = geom["d"] * 1e3

    phi_plot = np.ma.array(phi, mask=~geom["shield_inner"])

    fig, ax = plt.subplots(figsize=(7, 7))
    cf = ax.contourf(X, Y, phi_plot, levels=nfill)
    ax.contour(X, Y, phi_plot, levels=nline, linewidths=0.8)

    shield = Circle((0.0, 0.0), r2, fill=False, linewidth=2.0)
    core_left = Circle((-d / 2, 0.0), r1, fill=False, linewidth=2.0)
    core_right = Circle(( d / 2, 0.0), r1, fill=False, linewidth=2.0)

    ax.add_patch(shield)
    ax.add_patch(core_left)
    ax.add_patch(core_right)

    ax.set_aspect("equal")
    ax.set_xlabel("x [mm]")
    ax.set_ylabel("y [mm]")
    ax.set_title(title)
    fig.colorbar(cf, ax=ax, label="Potential [V]")
    plt.tight_layout()
    plt.show()


def pct(x):
    return 100.0 * x


def pf_per_m(x):
    return 1e12 * x


def richardson_extrapolation_three(h1, c1, h2, c2, h3, c3,
                                   p_min=0.05, p_max=8.0, nscan=4000):
    """
    Estimate p and C_inf from three arbitrary grid spacings h1 > h2 > h3
    assuming:
        C(h) = C_inf + a * h^p

    Returns
    -------
    dict with keys:
        success
        p
        c_inf
        a
        fit_rel_max
        fine_rel_error
        data_ratio
        model_ratio
        status
    """
    d12 = c1 - c2
    d23 = c2 - c3

    if abs(d12) < TINY or abs(d23) < TINY:
        return {
            "success": False, "p": np.nan, "c_inf": np.nan, "a": np.nan,
            "fit_rel_max": np.nan, "fine_rel_error": np.nan,
            "data_ratio": np.nan, "model_ratio": np.nan,
            "status": "differences too small"
        }

    # For asymptotic monotone convergence, d12 and d23 should have same sign
    if d12 * d23 <= 0:
        return {
            "success": False, "p": np.nan, "c_inf": np.nan, "a": np.nan,
            "fit_rel_max": np.nan, "fine_rel_error": np.nan,
            "data_ratio": d12 / d23, "model_ratio": np.nan,
            "status": "non-monotone or not yet asymptotic"
        }

    data_ratio = d12 / d23

    def model_ratio(p):
        num = h1**p - h2**p
        den = h2**p - h3**p
        return num / max(den, TINY)

    ps = np.linspace(p_min, p_max, nscan)
    vals = np.array([model_ratio(p) - data_ratio for p in ps])

    idx_candidates = np.where(np.sign(vals[:-1]) * np.sign(vals[1:]) <= 0)[0]

    if len(idx_candidates) > 0:
        # Bisection on the first sign-change interval
        k = idx_candidates[0]
        lo, hi = ps[k], ps[k + 1]
        flo, fhi = vals[k], vals[k + 1]

        for _ in range(80):
            mid = 0.5 * (lo + hi)
            fmid = model_ratio(mid) - data_ratio
            if flo * fmid <= 0:
                hi = mid
                fhi = fmid
            else:
                lo = mid
                flo = fmid

        p_est = 0.5 * (lo + hi)
    else:
        # Fallback: choose p minimizing mismatch
        k = np.argmin(np.abs(vals))
        p_est = ps[k]

    # Least-squares fit for C_inf and a with the estimated p
    H = np.array([
        [1.0, h1**p_est],
        [1.0, h2**p_est],
        [1.0, h3**p_est],
    ], dtype=float)
    y = np.array([c1, c2, c3], dtype=float)

    coeff, _, _, _ = np.linalg.lstsq(H, y, rcond=None)
    c_inf, a_est = coeff

    y_fit = H @ coeff
    fit_rel_max = np.max(np.abs(y_fit - y)) / max(abs(c_inf), TINY)
    fine_rel_error = abs(c_inf - c3) / max(abs(c_inf), TINY)

    return {
        "success": True,
        "p": p_est,
        "c_inf": c_inf,
        "a": a_est,
        "fit_rel_max": fit_rel_max,
        "fine_rel_error": fine_rel_error,
        "data_ratio": data_ratio,
        "model_ratio": model_ratio(p_est),
        "status": "ok"
    }


def convergence_study(r1, d, r2, eps_r, nx_list, symmetrize=True):
    """
    Run multiple grid sizes and build convergence / consistency diagnostics.

    Returns
    -------
    rows : list of dict
    """
    rows = []
    prev_C11 = None
    prev_Cll = None

    for nx in nx_list:
        geom, C_reduced, C_full, case10, case01 = compute_capacitance_matrices(
            r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=symmetrize
        )

        case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0)
        caps = extracted_capacitances(C_reduced)

        C11_charge = C_reduced[0, 0]
        C22_charge = C_reduced[1, 1]
        C12_charge = C_reduced[0, 1]
        C21_charge = C_reduced[1, 0]
        Cll_charge = caps["C_line_line_per_m"]

        # Energy-method estimates
        C11_energy = 2.0 * case10["energy_per_m"]      # because W' = 1/2 * C11 for [1,0,0]
        Cll_energy = 2.0 * case_diff["energy_per_m"]   # because W' = 1/2 * Cline for [+0.5,-0.5,0]

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

    # Richardson extrapolation on every last-three-grid window
    for i in range(2, len(rows)):
        r1_, r2_, r3_ = rows[i - 2], rows[i - 1], rows[i]

        rows[i]["rich_C11"] = richardson_extrapolation_three(
            r1_["h"], r1_["C11_charge"],
            r2_["h"], r2_["C11_charge"],
            r3_["h"], r3_["C11_charge"]
        )

        rows[i]["rich_Cll"] = richardson_extrapolation_three(
            r1_["h"], r1_["Cll_charge"],
            r2_["h"], r2_["Cll_charge"],
            r3_["h"], r3_["Cll_charge"]
        )

    return rows


def print_convergence_tables(rows):
    """
    Print:
      1) convergence + energy check
      2) symmetry / neutrality / residual
      3) Richardson extrapolation summary
    """
    print("\n=== Convergence and energy-method check ===")
    header1 = (
        f"{'nx':>6} "
        f"{'C11[q] pF/m':>14} {'C11[E] pF/m':>14} {'|Δ|/C11 [%]':>12} "
        f"{'Cll[q] pF/m':>14} {'Cll[E] pF/m':>14} {'|Δ|/Cll [%]':>12} "
        f"{'vs prev C11 [%]':>14} {'vs prev Cll [%]':>14}"
    )
    print(header1)
    print("-" * len(header1))

    for r in rows:
        s_prev_C11 = "   ---        " if np.isnan(r["rel_C11_vs_prev"]) else f"{pct(r['rel_C11_vs_prev']):14.6f}"
        s_prev_Cll = "   ---        " if np.isnan(r["rel_Cll_vs_prev"]) else f"{pct(r['rel_Cll_vs_prev']):14.6f}"

        print(
            f"{r['nx']:6d} "
            f"{pf_per_m(r['C11_charge']):14.6f} {pf_per_m(r['C11_energy']):14.6f} {pct(r['rel_C11_charge_vs_energy']):12.6f} "
            f"{pf_per_m(r['Cll_charge']):14.6f} {pf_per_m(r['Cll_energy']):14.6f} {pct(r['rel_Cll_charge_vs_energy']):12.6f} "
            f"{s_prev_C11} {s_prev_Cll}"
        )

    print("\n=== Symmetry / charge-neutrality / solver-residual check ===")
    header2 = (
        f"{'nx':>6} "
        f"{'sym diag [%]':>12} {'sym off [%]':>12} "
        f"{'neut10 [%]':>12} {'neut01 [%]':>12} {'neutDiff [%]':>14} "
        f"{'res10':>12} {'res01':>12} {'resDiff':>12}"
    )
    print(header2)
    print("-" * len(header2))

    for r in rows:
        print(
            f"{r['nx']:6d} "
            f"{pct(r['rel_sym_diag']):12.6f} {pct(r['rel_sym_off']):12.6f} "
            f"{pct(r['neutrality10']):12.6f} {pct(r['neutrality01']):12.6f} {pct(r['neutralityDiff']):14.6f} "
            f"{r['res10']:12.3e} {r['res01']:12.3e} {r['resDiff']:12.3e}"
        )

    print("\n=== Richardson extrapolation (from last 3 grids in each window) ===")
    header3 = (
        f"{'nx':>6} "
        f"{'p(C11)':>10} {'C11_inf pF/m':>16} {'fine err [%]':>14} {'fit err [%]':>12} "
        f"{'p(Cll)':>10} {'Cll_inf pF/m':>16} {'fine err [%]':>14} {'fit err [%]':>12}"
    )
    print(header3)
    print("-" * len(header3))

    for r in rows:
        rc11 = r["rich_C11"]
        rcll = r["rich_Cll"]

        if rc11 is None or rcll is None or (not rc11["success"]) or (not rcll["success"]):
            print(
                f"{r['nx']:6d} "
                f"{'---':>10} {'---':>16} {'---':>14} {'---':>12} "
                f"{'---':>10} {'---':>16} {'---':>14} {'---':>12}"
            )
        else:
            print(
                f"{r['nx']:6d} "
                f"{rc11['p']:10.4f} {pf_per_m(rc11['c_inf']):16.6f} {pct(rc11['fine_rel_error']):14.6f} {pct(rc11['fit_rel_max']):12.6f} "
                f"{rcll['p']:10.4f} {pf_per_m(rcll['c_inf']):16.6f} {pct(rcll['fine_rel_error']):14.6f} {pct(rcll['fit_rel_max']):12.6f}"
            )


def print_final_richardson_estimate(rows):
    """
    Print the best Richardson estimate from the last available 3-grid window.
    """
    if len(rows) < 3:
        print("\nNot enough grid points for Richardson extrapolation.")
        return

    last = rows[-1]
    rc11 = last["rich_C11"]
    rcll = last["rich_Cll"]

    print("\n=== Final Richardson estimate from the finest 3-grid window ===")
    print(f"Using nx = {rows[-3]['nx']}, {rows[-2]['nx']}, {rows[-1]['nx']}")

    if rc11 is not None and rc11["success"]:
        print(f"  C11_inf estimate       = {rc11['c_inf']:.6e} F/m  ({pf_per_m(rc11['c_inf']):.6f} pF/m)")
        print(f"  estimated order p      = {rc11['p']:.6f}")
        print(f"  finest-grid rel. error = {pct(rc11['fine_rel_error']):.6f} %")
        print(f"  3-point fit mismatch   = {pct(rc11['fit_rel_max']):.6f} %")
    else:
        print(f"  C11_inf estimate       = not available ({None if rc11 is None else rc11['status']})")

    if rcll is not None and rcll["success"]:
        print(f"  Cline_inf estimate     = {rcll['c_inf']:.6e} F/m  ({pf_per_m(rcll['c_inf']):.6f} pF/m)")
        print(f"  estimated order p      = {rcll['p']:.6f}")
        print(f"  finest-grid rel. error = {pct(rcll['fine_rel_error']):.6f} %")
        print(f"  3-point fit mismatch   = {pct(rcll['fit_rel_max']):.6f} %")
    else:
        print(f"  Cline_inf estimate     = not available ({None if rcll is None else rcll['status']})")


if __name__ == "__main__":
    # Example parameters
    r1 = 0.50e-3   # core radius [m]
    d  = 2.00e-3   # center-to-center spacing [m]
    r2 = 2.00e-3   # shield inner radius [m]
    eps_r = 2.20   # relative permittivity

    # Single-grid calculation
    nx = 301
    geom, C_reduced, C_full, case10, case01 = compute_capacitance_matrices(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=True
    )
    pretty_print_results(C_reduced, C_full)

    # Single-grid energy check
    case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0)
    caps = extracted_capacitances(C_reduced)

    C11_charge = C_reduced[0, 0]
    C11_energy = 2.0 * case10["energy_per_m"]
    Cll_charge = caps["C_line_line_per_m"]
    Cll_energy = 2.0 * case_diff["energy_per_m"]

    print("\nSingle-grid energy check:")
    print(f"  C11 charge-based      = {C11_charge:.6e} F/m  ({pf_per_m(C11_charge):.6f} pF/m)")
    print(f"  C11 energy-based      = {C11_energy:.6e} F/m  ({pf_per_m(C11_energy):.6f} pF/m)")
    print(f"  relative difference   = {pct(abs(C11_charge - C11_energy)/max(abs(C11_charge), TINY)):.6f} %")
    print()
    print(f"  Cline charge-based    = {Cll_charge:.6e} F/m  ({pf_per_m(Cll_charge):.6f} pF/m)")
    print(f"  Cline energy-based    = {Cll_energy:.6e} F/m  ({pf_per_m(Cll_energy):.6f} pF/m)")
    print(f"  relative difference   = {pct(abs(Cll_charge - Cll_energy)/max(abs(Cll_charge), TINY)):.6f} %")

    # Convergence study
    # ここは 3点以上必要。できれば finest 3点が十分細かい値にしてください。
    # 例: [151, 201, 301, 401, 501]
    # Richardson をより安定させたいなら、比が揃った [151, 301, 601] なども有効です。
    nx_list = [301, 601, 1201]
    rows = convergence_study(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx_list=nx_list, symmetrize=True
    )
    print_convergence_tables(rows)
    print_final_richardson_estimate(rows)

    # Plot examples
    plot_potential_contours(
        geom,
        case10["phi"],
        title="Potential contours: core1=1 V, core2=0 V, shield=0 V"
    )

    plot_potential_contours(
        geom,
        case_diff["phi"],
        title="Potential contours: core1=+0.5 V, core2=-0.5 V, shield=0 V"
    )