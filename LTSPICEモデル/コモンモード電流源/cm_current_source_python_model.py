#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CM current-source model for single-phase AC115V EMI filter.

Purpose
-------
DUMMY側のHOT/RTNとCHASSIS間に同相電流源を2個配置した場合に、
LDM1入力点、すなわち試験系と入力EMIフィルタの接点を流れる
HOT電流、RTN電流、コモンモード電流を計算します。

Outputs
-------
The following files are saved in the same folder as this script.

    cm_current_source_python_results.csv
    cm_current_source_python_current.png
    cm_current_source_python_phase.png

Model
-----
This is the symmetric common-mode equivalent model.

For symmetric HOT/RTN excitation:
    I_HOT = I_RTN
    I_CM  = I_HOT + I_RTN

The calculation target is the current through the LDM1 input plane.
The result was confirmed to match the corresponding LTspice model.

Notes
-----
- X capacitors, CR dampers, and the HOT-RTN dummy load do not affect the
  ideal symmetric common-mode equivalent because HOT and RTN move together.
- If HOT/RTN asymmetry is introduced later, use a full 2-line nodal model
  instead of this symmetric CM equivalent.
"""

from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt


# ============================================================
# User parameters
# ============================================================

PARAMS = {
    # Frequency sweep
    "f_start_Hz": 100.0,
    "f_stop_Hz": 100e6,
    "n_points": 1201,

    # Current sources at DUMMY side
    # HOT-to-CHASSIS and RTN-to-CHASSIS, same phase
    "I_source_each_A": 0.5,

    # Test-system source-side cable LINE_PS
    # For CM: two C_CHA capacitors are in parallel to chassis.
    "LINE_PS_C_CHA_each_F": 50e-12,

    # Feedthrough capacitors, HOT/RTN to chassis
    "C_THROUGH_each_F": 10e-6,
    "C_THROUGH_ESR_each_ohm": 20e-3,
    "C_THROUGH_ESL_each_H": 10e-9,

    # Test-system input line LINE_INPUT, each conductor
    # For CM, HOT/RTN conductors are in parallel:
    # Z_CM = (R + sL)/2
    "LINE_INPUT_L_each_H": 1e-6,
    "LINE_INPUT_DCR_each_ohm": 20e-3,
    "LINE_INPUT_C_CHA_each_F": 50e-12,

    # 1st differential-mode coils, used as common-mode path impedance
    # For CM, two identical branches are in parallel:
    # Z_CM = Z_each/2
    "LDM1_L_each_H": 68e-6,
    "LDM1_DCR_each_ohm": 20e-3,
    "LDM1_C_PA_each_F": 5e-12,

    # 2nd differential-mode coils
    "LDM2_L_each_H": 33e-6,
    "LDM2_DCR_each_ohm": 20e-3,
    "LDM2_C_PA_each_F": 5e-12,

    # Common-mode choke
    # In CM, each winding inductance is approximately L*(1+K).
    # The two windings are then in parallel for common-mode current.
    "CMC_L_each_H": 3e-3,
    "CMC_DCR_each_ohm": 100e-3,
    "CMC_K": 0.999,
    "CMC_C_PA_each_F": 5e-12,

    # Y capacitors, HOT/RTN to chassis
    "CY_each_F": 2.2e-9,
    "CY_ESR_each_ohm": 20e-3,
    "CY_ESL_each_H": 10e-9,
}


# ============================================================
# Basic impedance/admittance functions
# ============================================================

def z_cap_esr_esl(s, capacitance_F, esr_ohm, esl_H):
    """Series model: ESR + ESL + ideal capacitor."""
    return esr_ohm + s * esl_H + 1.0 / (s * capacitance_F)


def z_inductor_with_parallel_cap(s, inductance_H, dcr_ohm, c_parallel_F):
    """Inductor model: (DCR + sL) || C_parallel."""
    z_rl = dcr_ohm + s * inductance_H

    if c_parallel_F <= 0:
        return z_rl

    y_total = 1.0 / z_rl + s * c_parallel_F
    return 1.0 / y_total


def to_dbuA(current_A):
    """Convert current amplitude [A] to dB microampere."""
    return 20.0 * np.log10(np.maximum(np.abs(current_A), 1e-300) * 1e6)


def phase_deg(x):
    """Phase angle in degree."""
    return np.angle(x, deg=True)


# ============================================================
# Common-mode model
# ============================================================

def calculate_cm_currents(f_Hz, p=None):
    """
    Calculate I_HOT, I_RTN, and I_CM at the LDM1 input plane.

    Parameters
    ----------
    f_Hz : ndarray
        Frequency array [Hz]
    p : dict
        Parameter dictionary. If None, PARAMS is used.

    Returns
    -------
    results : dict
        Contains complex currents and useful intermediate quantities.
    """
    if p is None:
        p = PARAMS

    f_Hz = np.asarray(f_Hz, dtype=float)
    w = 2.0 * np.pi * f_Hz
    s = 1j * w

    # --------------------------------------------------------
    # Left-side shunt admittance at the feedthrough/source-side node
    #
    # LINE_PS C_CHA:
    #   HOT-to-chassis and RTN-to-chassis are in parallel for CM.
    #
    # Feedthrough caps:
    #   HOT-to-chassis and RTN-to-chassis are also in parallel for CM.
    # --------------------------------------------------------
    y_line_ps_ccha = s * 2.0 * p["LINE_PS_C_CHA_each_F"]

    z_through_each = z_cap_esr_esl(
        s,
        p["C_THROUGH_each_F"],
        p["C_THROUGH_ESR_each_ohm"],
        p["C_THROUGH_ESL_each_H"],
    )
    y_through_pair = 2.0 / z_through_each

    y_left_node1 = y_line_ps_ccha + y_through_pair

    # --------------------------------------------------------
    # LINE_INPUT common-mode series impedance and shunt capacitance
    # --------------------------------------------------------
    z_line_input_cm = (
        p["LINE_INPUT_DCR_each_ohm"] + s * p["LINE_INPUT_L_each_H"]
    ) / 2.0

    y_line_input_ccha = s * 2.0 * p["LINE_INPUT_C_CHA_each_F"]

    # --------------------------------------------------------
    # LDM1 and LDM2 common-mode path impedances
    # Each line has one coil. In CM, the two branches are in parallel.
    # --------------------------------------------------------
    z_ldm1_each = z_inductor_with_parallel_cap(
        s,
        p["LDM1_L_each_H"],
        p["LDM1_DCR_each_ohm"],
        p["LDM1_C_PA_each_F"],
    )
    z_ldm1_cm = z_ldm1_each / 2.0

    z_ldm2_each = z_inductor_with_parallel_cap(
        s,
        p["LDM2_L_each_H"],
        p["LDM2_DCR_each_ohm"],
        p["LDM2_C_PA_each_F"],
    )
    z_ldm2_cm = z_ldm2_each / 2.0

    # --------------------------------------------------------
    # CMC common-mode impedance
    #
    # For common-mode excitation, the effective inductance of one winding is:
    #     L_cm_each_effective = L_each * (1 + K)
    #
    # Then HOT/RTN windings are in parallel in the common-mode equivalent:
    #     Z_cmc_cm = Z_each / 2
    # --------------------------------------------------------
    l_cmc_each_cm = p["CMC_L_each_H"] * (1.0 + p["CMC_K"])

    z_cmc_each = z_inductor_with_parallel_cap(
        s,
        l_cmc_each_cm,
        p["CMC_DCR_each_ohm"],
        p["CMC_C_PA_each_F"],
    )
    z_cmc_cm = z_cmc_each / 2.0

    # --------------------------------------------------------
    # Y capacitors at load side
    # HOT/RTN to chassis are in parallel for common-mode.
    # --------------------------------------------------------
    z_cy_each = z_cap_esr_esl(
        s,
        p["CY_each_F"],
        p["CY_ESR_each_ohm"],
        p["CY_ESL_each_H"],
    )
    y_ycap_pair = 2.0 / z_cy_each

    # --------------------------------------------------------
    # Recursive left-looking admittance calculation
    #
    # Node naming from source side to load side:
    #
    #   Node1: feedthrough/source-side shunt node
    #   Node2: LDM1 input plane, after LINE_INPUT
    #   Node3: after LDM1
    #   Node4: after LDM2, before CMC
    #   Node5: load-side node where current sources and Y caps exist
    #
    # We inject I_total at Node5 and calculate current flowing leftward
    # through the LDM1 branch. The measured HOT/RTN currents are half of
    # that common-mode branch current due to symmetry.
    # --------------------------------------------------------

    # Looking left from Node1
    y_left_1 = y_left_node1

    # Looking left from Node2:
    # shunt C_CHA at Node2 in parallel with series LINE_INPUT to y_left_1
    y_branch_2_to_1 = 1.0 / (z_line_input_cm + 1.0 / y_left_1)
    y_left_2 = y_line_input_ccha + y_branch_2_to_1

    # Looking left from Node3 through LDM1
    y_branch_3_to_2 = 1.0 / (z_ldm1_cm + 1.0 / y_left_2)
    y_left_3 = y_branch_3_to_2

    # Looking left from Node4 through LDM2
    y_branch_4_to_3 = 1.0 / (z_ldm2_cm + 1.0 / y_left_3)
    y_left_4 = y_branch_4_to_3

    # Looking left from Node5 through CMC, with Y caps also connected
    y_branch_5_to_4 = 1.0 / (z_cmc_cm + 1.0 / y_left_4)
    y_left_5 = y_ycap_pair + y_branch_5_to_4

    # Total injected common-mode current
    i_total_source = 2.0 * p["I_source_each_A"]

    # Node5 voltage versus chassis
    v5 = i_total_source / y_left_5

    # Current flowing leftward through CMC
    i_cmc_cm = v5 * y_branch_5_to_4

    # Node4 voltage
    v4 = i_cmc_cm / y_left_4

    # Current flowing leftward through LDM2
    i_ldm2_cm = v4 * y_branch_4_to_3

    # Node3 voltage
    v3 = i_ldm2_cm / y_left_3

    # Current flowing leftward through LDM1
    i_ldm1_cm = v3 * y_branch_3_to_2

    # Symmetric HOT/RTN currents at LDM1 input plane
    i_hot = i_ldm1_cm / 2.0
    i_rtn = i_ldm1_cm / 2.0
    i_cm = i_hot + i_rtn

    return {
        "frequency_Hz": f_Hz,
        "I_HOT_A": i_hot,
        "I_RTN_A": i_rtn,
        "I_CM_A": i_cm,
        "V_load_node_to_chassis_V": v5,
        "I_total_source_A": np.full_like(f_Hz, i_total_source, dtype=complex),
        "I_CMC_CM_A": i_cmc_cm,
        "I_LDM2_CM_A": i_ldm2_cm,
        "I_LDM1_CM_A": i_ldm1_cm,
        "Z_LINE_INPUT_CM_ohm": z_line_input_cm,
        "Z_LDM1_CM_ohm": z_ldm1_cm,
        "Z_LDM2_CM_ohm": z_ldm2_cm,
        "Z_CMC_CM_ohm": z_cmc_cm,
    }


# ============================================================
# Output helpers
# ============================================================

def get_output_dir():
    """Save files beside this script. In notebooks, save in current folder."""
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def save_csv(path, results):
    f = results["frequency_Hz"]
    i_hot = results["I_HOT_A"]
    i_rtn = results["I_RTN_A"]
    i_cm = results["I_CM_A"]

    data = np.column_stack([
        f,
        np.abs(i_hot),
        np.abs(i_rtn),
        np.abs(i_cm),
        to_dbuA(i_hot),
        to_dbuA(i_rtn),
        to_dbuA(i_cm),
        np.real(i_hot),
        np.imag(i_hot),
        np.real(i_rtn),
        np.imag(i_rtn),
        np.real(i_cm),
        np.imag(i_cm),
        phase_deg(i_hot),
        phase_deg(i_rtn),
        phase_deg(i_cm),
    ])

    header = (
        "frequency_Hz,"
        "I_HOT_abs_A,I_RTN_abs_A,I_CM_abs_A,"
        "I_HOT_dBuA,I_RTN_dBuA,I_CM_dBuA,"
        "I_HOT_real_A,I_HOT_imag_A,"
        "I_RTN_real_A,I_RTN_imag_A,"
        "I_CM_real_A,I_CM_imag_A,"
        "I_HOT_phase_deg,I_RTN_phase_deg,I_CM_phase_deg"
    )

    np.savetxt(str(path), data, delimiter=",", header=header, comments="")


def make_current_plot(path, results):
    f = results["frequency_Hz"]
    i_hot = results["I_HOT_A"]
    i_rtn = results["I_RTN_A"]
    i_cm = results["I_CM_A"]

    plt.figure(figsize=(10, 6))
    plt.semilogx(f, to_dbuA(i_hot), label="I_HOT at LDM1 input")
    plt.semilogx(f, to_dbuA(i_rtn), label="I_RTN at LDM1 input")
    plt.semilogx(f, to_dbuA(i_cm), label="I_CM = I_HOT + I_RTN")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA]")
    plt.title("Common-mode current from load-side current sources")
    plt.legend()
    plt.tight_layout()
    plt.savefig(str(path), dpi=180)
    plt.close()


def make_phase_plot(path, results):
    f = results["frequency_Hz"]
    i_hot = results["I_HOT_A"]
    i_rtn = results["I_RTN_A"]
    i_cm = results["I_CM_A"]

    plt.figure(figsize=(10, 5))
    plt.semilogx(f, phase_deg(i_hot), label="I_HOT phase")
    plt.semilogx(f, phase_deg(i_rtn), label="I_RTN phase")
    plt.semilogx(f, phase_deg(i_cm), label="I_CM phase")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Phase [deg]")
    plt.title("Current phase at LDM1 input")
    plt.legend()
    plt.tight_layout()
    plt.savefig(str(path), dpi=180)
    plt.close()


def print_representative_points(results):
    f = results["frequency_Hz"]
    i_hot = results["I_HOT_A"]
    i_rtn = results["I_RTN_A"]
    i_cm = results["I_CM_A"]

    print("==== Representative points ====")
    for ff in [100, 400, 1e3, 10e3, 100e3, 1e6, 10e6, 100e6]:
        idx = int(np.argmin(np.abs(f - ff)))
        print(
            f"{f[idx]:>12.6g} Hz : "
            f"|I_HOT|={abs(i_hot[idx]):.9g} A, "
            f"|I_RTN|={abs(i_rtn[idx]):.9g} A, "
            f"|I_CM|={abs(i_cm[idx]):.9g} A, "
            f"I_CM={to_dbuA(i_cm[idx]):.6f} dBµA"
        )


def main():
    output_dir = get_output_dir()

    f = np.logspace(
        np.log10(PARAMS["f_start_Hz"]),
        np.log10(PARAMS["f_stop_Hz"]),
        int(PARAMS["n_points"]),
    )

    results = calculate_cm_currents(f, PARAMS)

    csv_path = output_dir / "cm_current_source_python_results.csv"
    current_png_path = output_dir / "cm_current_source_python_current.png"
    phase_png_path = output_dir / "cm_current_source_python_phase.png"

    save_csv(csv_path, results)
    make_current_plot(current_png_path, results)
    make_phase_plot(phase_png_path, results)

    print("==== Output files ====")
    print(f"CSV          : {csv_path}")
    print(f"Current plot : {current_png_path}")
    print(f"Phase plot   : {phase_png_path}")
    print()

    print("==== Model summary ====")
    print(f"I_source_each = {PARAMS['I_source_each_A']} A")
    print(f"I_total_CM    = {2*PARAMS['I_source_each_A']} A")
    print("Measurement plane: LDM1 input, test-system/EMI-filter interface")
    print("Definition: I_CM = I_HOT + I_RTN")
    print()

    print_representative_points(results)


if __name__ == "__main__":
    main()
