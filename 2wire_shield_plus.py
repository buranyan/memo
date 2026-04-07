import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Circle
from scipy.sparse import csr_matrix
from scipy.sparse.linalg import spsolve

EPS0 = 8.8541878128e-12  # vacuum permittivity [F/m]


def build_geometry(r1, d, r2, nx):
    """
    Geometry:
      - outer circular shield (inner radius r2)
      - two identical circular cores (radius r1)
      - core centers at x = +/- d/2, y = 0

    Parameters
    ----------
    r1 : float
        Core radius [m]
    d : float
        Center-to-center spacing between the two cores [m]
    r2 : float
        Inner radius of the shield [m]
    nx : int
        Number of grid points in x and y (use an odd number)

    Returns
    -------
    geom : dict
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
    """
    A, b = assemble_system(geom, V1, V2, Vshield)
    phi_unknown = spsolve(A, b)

    phi = np.full_like(geom["X"], Vshield, dtype=float)
    phi[geom["dielectric"]] = phi_unknown
    phi[geom["core1"]] = V1
    phi[geom["core2"]] = V2
    return phi


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


def solve_case(geom, eps_r, V1, V2, Vshield=0.0):
    """
    Solve one excitation case and return conductor charges per unit length.
    """
    eps = EPS0 * eps_r
    phi = solve_potential(geom, V1, V2, Vshield)

    q1 = conductor_charge_per_length(phi, geom["core1"], geom["dielectric"], V1, eps)
    q2 = conductor_charge_per_length(phi, geom["core2"], geom["dielectric"], V2, eps)
    qs = -(q1 + q2)  # charge neutrality

    return {
        "phi": phi,
        "q_per_m": np.array([q1, q2, qs], dtype=float),
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

    return geom, C_reduced, C_full


def extracted_capacitances(C_reduced):
    """
    Derive practical capacitances from the reduced 2x2 matrix.
    """
    C11, C12 = C_reduced[0, 0], C_reduced[0, 1]
    C21, C22 = C_reduced[1, 0], C_reduced[1, 1]

    C_line_line = (C11 - C12 - C21 + C22) / 4.0
    C_core1_shield_other_grounded = C11
    C_core1_shield_other_floating = C11 - (C12 * C21) / C22

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

    Parameters
    ----------
    geom : dict
        Geometry dictionary returned by build_geometry()
    phi : ndarray
        Potential distribution
    title : str
        Plot title
    nfill : int
        Number of filled contour levels
    nline : int
        Number of contour-line levels
    """
    X = geom["X"] * 1e3  # convert to mm for display
    Y = geom["Y"] * 1e3
    r1 = geom["r1"] * 1e3
    r2 = geom["r2"] * 1e3
    d = geom["d"] * 1e3

    # mask outside shield
    phi_plot = np.ma.array(phi, mask=~geom["shield_inner"])

    fig, ax = plt.subplots(figsize=(7, 7))

    cf = ax.contourf(X, Y, phi_plot, levels=nfill)
    ax.contour(X, Y, phi_plot, levels=nline, linewidths=0.8)

    # draw boundaries
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


if __name__ == "__main__":
    # Example parameters
    r1 = 0.50e-3   # core radius [m]
    d  = 2.00e-3   # center-to-center spacing [m]
    r2 = 2.00e-3   # shield inner radius [m]
    eps_r = 2.20   # relative permittivity
    nx = 501       # odd number; larger = better accuracy, slower

    geom, C_reduced, C_full = compute_capacitance_matrices(
        r1=r1, d=d, r2=r2, eps_r=eps_r, nx=nx, symmetrize=True
    )
    pretty_print_results(C_reduced, C_full)

    # Example 1: core1 = 1 V, core2 = 0 V, shield = 0 V
    case10 = solve_case(geom, eps_r=eps_r, V1=1.0, V2=0.0, Vshield=0.0)
    plot_potential_contours(
        geom,
        case10["phi"],
        title="Potential contours: core1=1 V, core2=0 V, shield=0 V"
    )

    # Example 2: differential excitation
    case_diff = solve_case(geom, eps_r=eps_r, V1=0.5, V2=-0.5, Vshield=0.0)
    plot_potential_contours(
        geom,
        case_diff["phi"],
        title="Potential contours: core1=+0.5 V, core2=-0.5 V, shield=0 V"
    )