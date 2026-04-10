import numpy as np
from scipy.sparse import csr_matrix
from scipy.sparse.linalg import spsolve

EPS0 = 8.8541878128e-12
TINY = 1e-30


# ============================================================
# Geometry / region classification
# ============================================================

def classify_region(x, y, r1, d, r2):
    """
    Returns one of:
      'dielectric', 'core1', 'core2', 'shield'
    """
    if x * x + y * y >= r2 * r2:
        return "shield"
    if (x - d / 2) ** 2 + y * y <= r1 * r1:
        return "core1"
    if (x + d / 2) ** 2 + y * y <= r1 * r1:
        return "core2"
    return "dielectric"


def build_geometry(r1, d, r2, nx):
    """
    Cell-centered grid.
    nx = number of cells across the full diameter 2*r2.
    Use odd nx.
    """
    if not (r1 > 0 and d > 0 and r2 > 0):
        raise ValueError("r1, d, r2 must be positive.")
    if d <= 2 * r1:
        raise ValueError("Need d > 2*r1.")
    if d / 2 + r1 >= r2:
        raise ValueError("Need d/2 + r1 < r2.")
    if nx < 51:
        raise ValueError("Use nx >= 51.")
    if nx % 2 == 0:
        raise ValueError("Use odd nx.")

    h = 2.0 * r2 / nx
    x = np.linspace(-r2 + 0.5 * h, r2 - 0.5 * h, nx)
    X, Y = np.meshgrid(x, x, indexing="xy")

    region = np.empty((nx, nx), dtype=object)
    dielectric = np.zeros((nx, nx), dtype=bool)

    for i in range(nx):
        for j in range(nx):
            tag = classify_region(x[j], x[i], r1, d, r2)
            region[i, j] = tag
            dielectric[i, j] = (tag == "dielectric")

    return {
        "x": x,
        "X": X,
        "Y": Y,
        "h": h,
        "nx": nx,
        "r1": r1,
        "d": d,
        "r2": r2,
        "region": region,
        "dielectric": dielectric,
    }


# ============================================================
# Exact circle hit along cardinal direction
# ============================================================

def x_hit_from_point(y, x0, sign, cx, cy, r):
    """
    On horizontal line y = const, find nearest circle intersection in +x or -x.
    sign = +1 or -1
    """
    yy = y - cy
    if abs(yy) > r:
        return None

    dx = np.sqrt(max(r * r - yy * yy, 0.0))
    roots = [cx - dx, cx + dx]
    cand = [xr for xr in roots if sign * (xr - x0) > 1e-14]
    if not cand:
        return None
    return min(cand, key=lambda z: sign * (z - x0))


def y_hit_from_point(x, y0, sign, cx, cy, r):
    """
    On vertical line x = const, find nearest circle intersection in +y or -y.
    sign = +1 or -1
    """
    xx = x - cx
    if abs(xx) > r:
        return None

    dy = np.sqrt(max(r * r - xx * xx, 0.0))
    roots = [cy - dy, cy + dy]
    cand = [yr for yr in roots if sign * (yr - y0) > 1e-14]
    if not cand:
        return None
    return min(cand, key=lambda z: sign * (z - y0))


def boundary_distance_along_direction(x0, y0, direction, transverse, geom):
    """
    Distance from cell center (x0,y0) to nearest conductor boundary along
    one face-normal line.

    direction: 'xp', 'xm', 'yp', 'ym'
    transverse:
      - if direction is xpm/xm: transverse = y sample coordinate
      - if direction is ypm/ym: transverse = x sample coordinate

    Returns
    -------
    dist, tag
      tag in {'core1','core2','shield'}
    """
    r1 = geom["r1"]
    d = geom["d"]
    r2 = geom["r2"]

    best_dist = 1e100
    best_tag = "shield"

    circles = [
        ("core1", +d / 2, 0.0, r1),
        ("core2", -d / 2, 0.0, r1),
        ("shield", 0.0, 0.0, r2),
    ]

    if direction == "xp":
        y = transverse
        for tag, cx, cy, r in circles:
            xb = x_hit_from_point(y, x0, +1, cx, cy, r)
            if xb is not None:
                dist = xb - x0
                if dist < best_dist:
                    best_dist = dist
                    best_tag = tag

    elif direction == "xm":
        y = transverse
        for tag, cx, cy, r in circles:
            xb = x_hit_from_point(y, x0, -1, cx, cy, r)
            if xb is not None:
                dist = x0 - xb
                if dist < best_dist:
                    best_dist = dist
                    best_tag = tag

    elif direction == "yp":
        x = transverse
        for tag, cx, cy, r in circles:
            yb = y_hit_from_point(x, y0, +1, cx, cy, r)
            if yb is not None:
                dist = yb - y0
                if dist < best_dist:
                    best_dist = dist
                    best_tag = tag

    elif direction == "ym":
        x = transverse
        for tag, cx, cy, r in circles:
            yb = y_hit_from_point(x, y0, -1, cx, cy, r)
            if yb is not None:
                dist = y0 - yb
                if dist < best_dist:
                    best_dist = dist
                    best_tag = tag

    else:
        raise ValueError("Unknown direction")

    return max(best_dist, 1e-14), best_tag


# ============================================================
# 1D Gauss-Legendre quadrature points on a face
# ============================================================

def face_quadrature(center, h, direction, nquad):
    """
    Returns sample coordinates along one face:
      - for xp/xm: array of y-samples on the face
      - for yp/ym: array of x-samples on the face
    """
    # Gauss-Legendre points on [-1,1]
    xi, wi = np.polynomial.legendre.leggauss(nquad)
    s = 0.5 * h * xi
    w = 0.5 * h * wi

    if direction in ("xp", "xm"):
        return center + s, w
    elif direction in ("yp", "ym"):
        return center + s, w
    else:
        raise ValueError("Unknown direction")


# ============================================================
# Conductance network assembly
# ============================================================

def internal_face_conductance(i, j, direction, geom, eps, nquad):
    """
    Shared conductance between two neighboring dielectric cells across one face.
    Only called for +x and +y faces so each shared face is handled once.

    direction: 'xp' or 'yp'
    """
    x = geom["x"]
    h = geom["h"]
    dielectric = geom["dielectric"]
    nx = geom["nx"]
    r1 = geom["r1"]
    d = geom["d"]
    r2 = geom["r2"]

    x0 = x[j]
    y0 = x[i]

    if direction == "xp":
        ni, nj = i, j + 1
        if not (0 <= nj < nx and dielectric[ni, nj]):
            return 0.0

        x_face = x0 + 0.5 * h
        transverse, weights = face_quadrature(y0, h, "xp", nquad)
        G = 0.0

        for yq, wq in zip(transverse, weights):
            left_tag = classify_region(x_face - 1e-14, yq, r1, d, r2)
            right_tag = classify_region(x_face + 1e-14, yq, r1, d, r2)

            if left_tag == "dielectric" and right_tag == "dielectric":
                G += eps * wq / h

        return G

    elif direction == "yp":
        ni, nj = i + 1, j
        if not (0 <= ni < nx and dielectric[ni, nj]):
            return 0.0

        y_face = y0 + 0.5 * h
        transverse, weights = face_quadrature(x0, h, "yp", nquad)
        G = 0.0

        for xq, wq in zip(transverse, weights):
            down_tag = classify_region(xq, y_face - 1e-14, r1, d, r2)
            up_tag = classify_region(xq, y_face + 1e-14, r1, d, r2)

            if down_tag == "dielectric" and up_tag == "dielectric":
                G += eps * wq / h

        return G

    else:
        raise ValueError("direction must be 'xp' or 'yp'")


def boundary_face_conductance(i, j, direction, geom, eps, nquad):
    """
    Boundary conductance from one dielectric cell to conductors/shield
    across one face.

    This is only evaluated when the neighbor cell center in that direction
    is not dielectric. That avoids double-counting shared dielectric faces.

    Returns dict:
      {'core1': G1, 'core2': G2, 'shield': Gs}
    """
    x = geom["x"]
    h = geom["h"]
    dielectric = geom["dielectric"]
    nx = geom["nx"]
    r1 = geom["r1"]
    d = geom["d"]
    r2 = geom["r2"]

    x0 = x[j]
    y0 = x[i]

    out = {"core1": 0.0, "core2": 0.0, "shield": 0.0}

    if direction == "xp":
        ni, nj = i, j + 1
        if 0 <= nj < nx and dielectric[ni, nj]:
            return out

        x_face = x0 + 0.5 * h
        transverse, weights = face_quadrature(y0, h, "xp", nquad)

        for yq, wq in zip(transverse, weights):
            tag = classify_region(x_face + 1e-14, yq, r1, d, r2)
            if tag != "dielectric":
                dist, btag = boundary_distance_along_direction(x0, y0, "xp", yq, geom)
                out[btag] += eps * wq / dist

    elif direction == "xm":
        ni, nj = i, j - 1
        if 0 <= nj < nx and dielectric[ni, nj]:
            return out

        x_face = x0 - 0.5 * h
        transverse, weights = face_quadrature(y0, h, "xm", nquad)

        for yq, wq in zip(transverse, weights):
            tag = classify_region(x_face - 1e-14, yq, r1, d, r2)
            if tag != "dielectric":
                dist, btag = boundary_distance_along_direction(x0, y0, "xm", yq, geom)
                out[btag] += eps * wq / dist

    elif direction == "yp":
        ni, nj = i + 1, j
        if 0 <= ni < nx and dielectric[ni, nj]:
            return out

        y_face = y0 + 0.5 * h
        transverse, weights = face_quadrature(x0, h, "yp", nquad)

        for xq, wq in zip(transverse, weights):
            tag = classify_region(xq, y_face + 1e-14, r1, d, r2)
            if tag != "dielectric":
                dist, btag = boundary_distance_along_direction(x0, y0, "yp", xq, geom)
                out[btag] += eps * wq / dist

    elif direction == "ym":
        ni, nj = i - 1, j
        if 0 <= ni < nx and dielectric[ni, nj]:
            return out

        y_face = y0 - 0.5 * h
        transverse, weights = face_quadrature(x0, h, "ym", nquad)

        for xq, wq in zip(transverse, weights):
            tag = classify_region(xq, y_face - 1e-14, r1, d, r2)
            if tag != "dielectric":
                dist, btag = boundary_distance_along_direction(x0, y0, "ym", xq, geom)
                out[btag] += eps * wq / dist

    else:
        raise ValueError("Unknown direction")

    return out


def build_conductance_network(geom, eps_r, nquad=8):
    """
    Builds a conductance network:
      - internal edges between neighboring dielectric cells
      - boundary conductances from dielectric cells to core1/core2/shield
    """
    eps = EPS0 * eps_r
    dielectric = geom["dielectric"]
    nx = geom["nx"]

    ids = -np.ones((nx, nx), dtype=int)
    pts = np.argwhere(dielectric)
    ids[dielectric] = np.arange(len(pts))
    n_unknown = len(pts)

    # boundary conductances per unknown
    G_core1 = np.zeros(n_unknown, dtype=float)
    G_core2 = np.zeros(n_unknown, dtype=float)
    G_shield = np.zeros(n_unknown, dtype=float)

    # internal edges: list of (p, q, Gpq)
    edges = []

    for p, (i, j) in enumerate(pts):
        # shared faces to east and north only
        if j + 1 < nx and dielectric[i, j + 1]:
            Gpq = internal_face_conductance(i, j, "xp", geom, eps, nquad)
            q = ids[i, j + 1]
            if Gpq > 0.0:
                edges.append((p, q, Gpq))

        if i + 1 < nx and dielectric[i + 1, j]:
            Gpq = internal_face_conductance(i, j, "yp", geom, eps, nquad)
            q = ids[i + 1, j]
            if Gpq > 0.0:
                edges.append((p, q, Gpq))

        # boundary conductances on all four faces,
        # but only on faces without dielectric neighbor center
        g = boundary_face_conductance(i, j, "xp", geom, eps, nquad)
        G_core1[p] += g["core1"]
        G_core2[p] += g["core2"]
        G_shield[p] += g["shield"]

        g = boundary_face_conductance(i, j, "xm", geom, eps, nquad)
        G_core1[p] += g["core1"]
        G_core2[p] += g["core2"]
        G_shield[p] += g["shield"]

        g = boundary_face_conductance(i, j, "yp", geom, eps, nquad)
        G_core1[p] += g["core1"]
        G_core2[p] += g["core2"]
        G_shield[p] += g["shield"]

        g = boundary_face_conductance(i, j, "ym", geom, eps, nquad)
        G_core1[p] += g["core1"]
        G_core2[p] += g["core2"]
        G_shield[p] += g["shield"]

    return {
        "ids": ids,
        "pts": pts,
        "n_unknown": n_unknown,
        "edges": edges,
        "G_core1": G_core1,
        "G_core2": G_core2,
        "G_shield": G_shield,
    }


# ============================================================
# Solve one excitation case
# ============================================================

def solve_network_case(net, V1, V2, Vshield):
    """
    Solve A phi = b for one conductor excitation.
    """
    n = net["n_unknown"]
    rows = []
    cols = []
    data = []
    b = np.zeros(n, dtype=float)

    diag = np.zeros(n, dtype=float)

    # internal edges
    for p, q, Gpq in net["edges"]:
        diag[p] += Gpq
        diag[q] += Gpq

        rows.append(p)
        cols.append(q)
        data.append(-Gpq)

        rows.append(q)
        cols.append(p)
        data.append(-Gpq)

    # boundary conductances
    G1 = net["G_core1"]
    G2 = net["G_core2"]
    Gs = net["G_shield"]

    diag += G1 + G2 + Gs
    b += G1 * V1 + G2 * V2 + Gs * Vshield

    # diagonal
    idx = np.arange(n)
    rows.extend(idx.tolist())
    cols.extend(idx.tolist())
    data.extend(diag.tolist())

    A = csr_matrix((data, (rows, cols)), shape=(n, n))
    phi = spsolve(A, b)

    r = A @ phi - b
    rel_res = np.linalg.norm(r) / max(np.linalg.norm(b), TINY)

    return phi, rel_res


def conductor_charges_and_energy(net, phi, V1, V2, Vshield):
    """
    Charges and energy from the same conductance network.
    This guarantees internal consistency of charge- and energy-based extraction.
    """
    G1 = net["G_core1"]
    G2 = net["G_core2"]
    Gs = net["G_shield"]

    q1 = np.sum(G1 * (V1 - phi))
    q2 = np.sum(G2 * (V2 - phi))
    qs = np.sum(Gs * (Vshield - phi))

    W = 0.0
    for p, q, Gpq in net["edges"]:
        W += 0.5 * Gpq * (phi[p] - phi[q]) ** 2

    W += 0.5 * np.sum(G1 * (V1 - phi) ** 2)
    W += 0.5 * np.sum(G2 * (V2 - phi) ** 2)
    W += 0.5 * np.sum(Gs * (Vshield - phi) ** 2)

    neutrality_rel = abs(q1 + q2 + qs) / max(abs(q1) + abs(q2) + abs(qs), TINY)

    return np.array([q1, q2, qs], dtype=float), W, neutrality_rel


def solve_case(geom, eps_r, V1, V2, Vshield=0.0, nquad=8):
    net = build_conductance_network(geom, eps_r, nquad=nquad)
    phi_unknown, rel_res = solve_network_case(net, V1, V2, Vshield)
    q_per_m, Wp, neutrality_rel = conductor_charges_and_energy(net, phi_unknown, V1, V2, Vshield)

    # reconstruct array on cell centers for optional plotting
    phi = np.full((geom["nx"], geom["nx"]), np.nan, dtype=float)
    phi[geom["dielectric"]] = phi_unknown

    # fill conductor cells with conductor potentials for convenience
    phi[geom["region"] == "core1"] = V1
    phi[geom["region"] == "core2"] = V2
    phi[geom["region"] == "shield"] = Vshield

    return {
        "phi": phi,
        "q_per_m": q_per_m,
        "energy_per_m": Wp,
        "rel_residual": rel_res,
        "neutrality_rel": neutrality_rel,
    }


# ============================================================
# Capacitance matrices
# ============================================================

def compute_capacitance_matrices(r1, d, r2, eps_r, nx=301, symmetrize=True, nquad=8):
    geom = build_geometry(r1=r1, d=d, r2=r2, nx=nx)

    case10 = solve_case(geom, eps_r=eps_r, V1=1.0, V2=0.0, Vshield=0.0, nquad=nquad)
    case01 = solve_case(geom, eps_r=eps_r, V1=0.0, V2=1.0, Vshield=0.0, nquad=nquad)

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


# ============================================================
# Diagnostics
# ============================================================

def pct(x):
    return 100.0 * x


def pf_per_m(x):
    return 1e12 * x


def richardson_extrapolation_three(h1, c1, h2, c2, h3, c3, p_min=0.05, p_max=8.0, nscan=4000):
    d12 = c1 - c2
    d23 = c2 - c3

    if abs(d12) < TINY or abs(d23) < TINY:
        return {"success": False, "status": "differences too small"}
    if d12 * d23 <= 0:
        return {"success": False, "status": "non-monotone"}

    data_ratio = d12 / d23

    def model_ratio(p):
        return (h1 ** p - h2 ** p) / max(h2 ** p - h3 ** p, TINY)

    ps = np.linspace(p_min, p_max, nscan)
    vals = np.array([model_ratio(p) - data_ratio for p in ps])

    idx = np.where(np.sign(vals[:-1]) * np.sign(vals[1:]) <= 0)[0]
    if len(idx) > 0:
        k = idx[0]
        lo, hi = ps[k], ps[k + 1]
        flo = vals[k]
        for _ in range(80):
            mid = 0.5 * (lo + hi)
            fmid = model_ratio(mid) - data_ratio
            if flo * fmid <= 0:
                hi = mid
            else:
                lo = mid
                flo = fmid
        p_est = 0.5 * (lo + hi)
    else:
        p_est = ps[np.argmin(np.abs(vals))]

    H = np.array([
        [1.0, h1 ** p_est],
        [1.0, h2 ** p_est],
        [1.0, h3 ** p_est],
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


def convergence_study(r1, d, r2, eps_r, nx_list, symmetrize=True, nquad=8):
    rows = []
    prev_C11 = None
    prev_Cll = None

    for nx in nx_list:
        geom, C_reduced, C_full, case10, case01 = compute_capacitance_matrices(
            r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=symmetrize, nquad=nquad
        )

        case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0, nquad=nquad)
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


def print_convergence_tables(rows):
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

    print("\n=== Richardson extrapolation ===")
    header3 = (
        f"{'nx':>6} "
        f"{'p(C11)':>10} {'C11_inf pF/m':>16} {'fine err [%]':>14} "
        f"{'p(Cll)':>10} {'Cll_inf pF/m':>16} {'fine err [%]':>14}"
    )
    print(header3)
    print("-" * len(header3))

    for r in rows:
        rc11 = r["rich_C11"]
        rcll = r["rich_Cll"]

        if rc11 is None or rcll is None or (not rc11["success"]) or (not rcll["success"]):
            print(
                f"{r['nx']:6d} "
                f"{'---':>10} {'---':>16} {'---':>14} "
                f"{'---':>10} {'---':>16} {'---':>14}"
            )
        else:
            print(
                f"{r['nx']:6d} "
                f"{rc11['p']:10.4f} {pf_per_m(rc11['c_inf']):16.6f} {pct(rc11['fine_rel_error']):14.6f} "
                f"{rcll['p']:10.4f} {pf_per_m(rcll['c_inf']):16.6f} {pct(rcll['fine_rel_error']):14.6f}"
            )


def print_final_richardson_estimate(rows):
    if len(rows) < 3:
        print("\nNot enough grid sizes for Richardson extrapolation.")
        return

    last = rows[-1]
    rc11 = last["rich_C11"]
    rcll = last["rich_Cll"]

    print("\n=== Final Richardson estimate ===")
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


# ============================================================
# Main
# ============================================================

if __name__ == "__main__":
    r1 = 0.50e-3
    d = 2.00e-3
    r2 = 2.00e-3
    eps_r = 2.20

    # nquad: 6〜12 くらいから開始
    nquad = 8

    # single-grid run
    nx = 301
    geom, C_reduced, C_full, case10, case01 = compute_capacitance_matrices(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=True, nquad=nquad
    )

    print("Reduced capacitance matrix with shield grounded [F/m]:")
    print(C_reduced)
    print()
    print("Full 3x3 Maxwell capacitance matrix [F/m]:")
    print(C_full)
    print()

    caps = extracted_capacitances(C_reduced)
    print("Derived capacitances [F/m]:")
    for k, v in caps.items():
        print(f"  {k:45s} = {v:.6e}")

    # energy check
    case_diff = solve_case(geom, eps_r=eps_r, V1=+0.5, V2=-0.5, Vshield=0.0, nquad=nquad)

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

    # convergence
    nx_list = [301, 601, 1201]
    rows = convergence_study(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx_list=nx_list, symmetrize=True, nquad=nquad
    )
    print_convergence_tables(rows)
    print_final_richardson_estimate(rows)