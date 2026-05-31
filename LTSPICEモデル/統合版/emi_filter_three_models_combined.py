#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Combined EMI filter analysis model.

One execution calculates the following three analyses:

① 試験系から負荷を見た周波数特性
   - Differential/normal-mode voltage transfer Vload/Vsource
   - ROUT = 0 ohm and 50 ohm

② 負荷側から試験系を含めたコモンモード電流計算
   - Two current sources from DUMMY HOT/RTN to CHASSIS
   - Each amplitude = 0.5 A, same phase
   - Measurement plane: LDM1 input side
   - I_CM = I_HOT + I_RTN

③ 負荷側から試験系を含めたノーマルモード電流計算
   - One current source between DUMMY HOT and DUMMY RTN
   - Amplitude = 1 A
   - Measurement plane: LDM1 input side
   - I_NM = (I_HOT - I_RTN) / 2

Outputs are saved in the same folder as this script:
    01_frequency_response.csv
    01_frequency_response.png
    02_cm_current.csv
    02_cm_current.png
    02_cm_current_phase.png
    03_nm_current.csv
    03_nm_current.png
    03_nm_current_phase.png
    combined_three_models_results.csv
"""

from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt


PARAMS = {
    # Frequency sweep
    "f_start_Hz": 100.0,
    "f_stop_Hz": 100e6,
    "n_points": 1201,

    # Source/load
    "Vline_rms_V": 115.0,
    "Pload_W": 460.0,
    "DUMMY_R_ohm": 28.75,

    # Test-system source-side cable LINE_PS
    "LINE_PS_L_each_H": 1e-6,
    "LINE_PS_DCR_each_ohm": 20e-3,
    "LINE_PS_C_CHA_each_F": 50e-12,
    "LINE_PS_C_LINE_F": 50e-12,

    # Feedthrough capacitors, HOT/RTN to chassis
    "C_THROUGH_each_F": 10e-6,
    "C_THROUGH_ESR_each_ohm": 20e-3,
    "C_THROUGH_ESL_each_H": 10e-9,

    # Test-system input line LINE_INPUT
    "LINE_INPUT_L_each_H": 1e-6,
    "LINE_INPUT_DCR_each_ohm": 20e-3,
    "LINE_INPUT_C_CHA_each_F": 50e-12,
    "LINE_INPUT_C_LINE_F": 50e-12,

    # LDM1
    "LDM1_L_each_H": 68e-6,
    "LDM1_DCR_each_ohm": 20e-3,
    "LDM1_C_PA_each_F": 5e-12,

    # CX1 and damper 1
    "CX1_C_F": 0.47e-6,
    "CX1_ESR_ohm": 20e-3,
    "CX1_ESL_H": 10e-9,
    "DUMP1_C_F": 0.47e-6,
    "DUMP1_ESR_ohm": 20e-3,
    "DUMP1_ESL_H": 10e-9,
    "DUMP1_R_ohm": 15.0,

    # LDM2
    "LDM2_L_each_H": 33e-6,
    "LDM2_DCR_each_ohm": 20e-3,
    "LDM2_C_PA_each_F": 5e-12,

    # CX2 and damper 2
    "CX2_C_F": 0.47e-6,
    "CX2_ESR_ohm": 20e-3,
    "CX2_ESL_H": 10e-9,
    "DUMP2_C_F": 0.47e-6,
    "DUMP2_ESR_ohm": 20e-3,
    "DUMP2_ESL_H": 10e-9,
    "DUMP2_R_ohm": 15.0,

    # Common-mode choke
    "CMC_L_each_H": 3e-3,
    "CMC_DCR_each_ohm": 100e-3,
    "CMC_K": 0.999,
    "CMC_C_PA_each_F": 5e-12,
    "CMC_C_LINE_F": 10e-12,

    # Y capacitors
    "CY_each_F": 2.2e-9,
    "CY_ESR_each_ohm": 20e-3,
    "CY_ESL_each_H": 10e-9,

    # Noise current sources
    "CM_I_source_each_A": 0.5,
    "NM_I_source_A": 1.0,
}


# ============================================================
# Basic circuit functions
# ============================================================

def z_cap_esr_esl(s, C, ESR, ESL):
    return ESR + s * ESL + 1.0 / (s * C)


def z_inductor_with_parallel_cap(s, L, R, Cp):
    z_rl = R + s * L
    if Cp <= 0:
        return z_rl
    return 1.0 / (1.0 / z_rl + s * Cp)


def to_dbuA(i_A):
    return 20.0 * np.log10(np.maximum(np.abs(i_A), 1e-300) * 1e6)


def abcd_series(Z):
    one = np.ones_like(Z, dtype=complex)
    zero = np.zeros_like(Z, dtype=complex)
    return np.array([[one, Z], [zero, one]], dtype=complex)


def abcd_shunt(Y):
    one = np.ones_like(Y, dtype=complex)
    zero = np.zeros_like(Y, dtype=complex)
    return np.array([[one, zero], [Y, one]], dtype=complex)


def cascade_abcd(*matrices):
    M = matrices[0]
    for N in matrices[1:]:
        A = M[0, 0] * N[0, 0] + M[0, 1] * N[1, 0]
        B = M[0, 0] * N[0, 1] + M[0, 1] * N[1, 1]
        C = M[1, 0] * N[0, 0] + M[1, 1] * N[1, 0]
        D = M[1, 0] * N[0, 1] + M[1, 1] * N[1, 1]
        M = np.array([[A, B], [C, D]], dtype=complex)
    return M


# ============================================================
# ① Frequency response from test system to load
# ============================================================

def line_section_dm_for_abcd(s, L_each, R_each, C_cha_each, C_line):
    z_series_dm = 2.0 * (R_each + s * L_each)
    y_shunt_dm = s * (C_line + C_cha_each / 2.0)
    return [abcd_series(z_series_dm), abcd_shunt(y_shunt_dm)]


def shunt_pair_to_chassis_dm(s, C_each, ESR_each, ESL_each):
    z_each = z_cap_esr_esl(s, C_each, ESR_each, ESL_each)
    return abcd_shunt(1.0 / (2.0 * z_each))


def shunt_xcap_with_damper(s, Cx, Cx_ESR, Cx_ESL, Cd, Cd_ESR, Cd_ESL, Rd):
    z_x = z_cap_esr_esl(s, Cx, Cx_ESR, Cx_ESL)
    z_d = Rd + z_cap_esr_esl(s, Cd, Cd_ESR, Cd_ESL)
    return abcd_shunt(1.0 / z_x + 1.0 / z_d)


def dm_coil_pair_abcd(s, L_each, R_each, Cp_each):
    z_each = z_inductor_with_parallel_cap(s, L_each, R_each, Cp_each)
    return abcd_series(2.0 * z_each)


def cmc_dm_abcd(s, L_each, R_each, K, Cp_each, C_line):
    z_each = z_inductor_with_parallel_cap(s, L_each * (1.0 - K), R_each, Cp_each)
    return [abcd_series(2.0 * z_each), abcd_shunt(s * C_line)]


def calc_frequency_response(f_Hz, Rs_ohm=0.0, p=PARAMS):
    f_Hz = np.asarray(f_Hz, dtype=float)
    s = 1j * 2.0 * np.pi * f_Hz
    Rload = p["DUMMY_R_ohm"]

    sections = []
    sections += line_section_dm_for_abcd(s, p["LINE_PS_L_each_H"], p["LINE_PS_DCR_each_ohm"], p["LINE_PS_C_CHA_each_F"], p["LINE_PS_C_LINE_F"])
    sections.append(shunt_pair_to_chassis_dm(s, p["C_THROUGH_each_F"], p["C_THROUGH_ESR_each_ohm"], p["C_THROUGH_ESL_each_H"]))
    sections += line_section_dm_for_abcd(s, p["LINE_INPUT_L_each_H"], p["LINE_INPUT_DCR_each_ohm"], p["LINE_INPUT_C_CHA_each_F"], p["LINE_INPUT_C_LINE_F"])

    sections.append(dm_coil_pair_abcd(s, p["LDM1_L_each_H"], p["LDM1_DCR_each_ohm"], p["LDM1_C_PA_each_F"]))
    sections.append(shunt_xcap_with_damper(s, p["CX1_C_F"], p["CX1_ESR_ohm"], p["CX1_ESL_H"], p["DUMP1_C_F"], p["DUMP1_ESR_ohm"], p["DUMP1_ESL_H"], p["DUMP1_R_ohm"]))
    sections.append(dm_coil_pair_abcd(s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"]))
    sections.append(shunt_xcap_with_damper(s, p["CX2_C_F"], p["CX2_ESR_ohm"], p["CX2_ESL_H"], p["DUMP2_C_F"], p["DUMP2_ESR_ohm"], p["DUMP2_ESL_H"], p["DUMP2_R_ohm"]))
    sections += cmc_dm_abcd(s, p["CMC_L_each_H"], p["CMC_DCR_each_ohm"], p["CMC_K"], p["CMC_C_PA_each_F"], p["CMC_C_LINE_F"])
    sections.append(shunt_pair_to_chassis_dm(s, p["CY_each_F"], p["CY_ESR_each_ohm"], p["CY_ESL_each_H"]))

    M = cascade_abcd(*sections)
    A, B, C, D = M[0, 0], M[0, 1], M[1, 0], M[1, 1]
    H = 1.0 / (A + B / Rload + Rs_ohm * (C + D / Rload))
    return H


# ============================================================
# ② CM current source from load side
# ============================================================

def calc_cm_current(f_Hz, p=PARAMS):
    f_Hz = np.asarray(f_Hz, dtype=float)
    s = 1j * 2.0 * np.pi * f_Hz

    # Node1 shunt: source-side cable C_CHA + feedthrough caps.
    y1 = s * 2.0 * p["LINE_PS_C_CHA_each_F"]
    z_through_each = z_cap_esr_esl(s, p["C_THROUGH_each_F"], p["C_THROUGH_ESR_each_ohm"], p["C_THROUGH_ESL_each_H"])
    y1 += 2.0 / z_through_each

    # LINE_INPUT common-mode series and shunt.
    z_line_input_cm = (p["LINE_INPUT_DCR_each_ohm"] + s * p["LINE_INPUT_L_each_H"]) / 2.0
    y2 = s * 2.0 * p["LINE_INPUT_C_CHA_each_F"]

    z_ldm1_cm = z_inductor_with_parallel_cap(s, p["LDM1_L_each_H"], p["LDM1_DCR_each_ohm"], p["LDM1_C_PA_each_F"]) / 2.0
    z_ldm2_cm = z_inductor_with_parallel_cap(s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"]) / 2.0

    # CMC common-mode inductance per winding = L*(1+K), then two windings in parallel.
    z_cmc_cm = z_inductor_with_parallel_cap(
        s,
        p["CMC_L_each_H"] * (1.0 + p["CMC_K"]),
        p["CMC_DCR_each_ohm"],
        p["CMC_C_PA_each_F"],
    ) / 2.0

    z_cy_each = z_cap_esr_esl(s, p["CY_each_F"], p["CY_ESR_each_ohm"], p["CY_ESL_each_H"])
    y5 = 2.0 / z_cy_each

    y_left_1 = y1
    yb2 = 1.0 / (z_line_input_cm + 1.0 / y_left_1)
    y_left_2 = y2 + yb2

    yb3 = 1.0 / (z_ldm1_cm + 1.0 / y_left_2)
    y_left_3 = yb3

    yb4 = 1.0 / (z_ldm2_cm + 1.0 / y_left_3)
    y_left_4 = yb4

    yb5 = 1.0 / (z_cmc_cm + 1.0 / y_left_4)
    y_left_5 = y5 + yb5

    i_total = 2.0 * p["CM_I_source_each_A"]
    v5 = i_total / y_left_5

    i_cmc = v5 * yb5
    v4 = i_cmc / y_left_4

    i_ldm2 = v4 * yb4
    v3 = i_ldm2 / y_left_3

    i_ldm1_cm = v3 * yb3

    i_hot = i_ldm1_cm / 2.0
    i_rtn = i_ldm1_cm / 2.0
    i_cm = i_hot + i_rtn

    return {"I_HOT_A": i_hot, "I_RTN_A": i_rtn, "I_CM_A": i_cm}


# ============================================================
# ③ NM current source from load side
# ============================================================

def calc_nm_current(f_Hz, p=PARAMS):
    f_Hz = np.asarray(f_Hz, dtype=float)
    s = 1j * 2.0 * np.pi * f_Hz

    z_line_ps = 2.0 * (p["LINE_PS_DCR_each_ohm"] + s * p["LINE_PS_L_each_H"])
    y_line_ps = s * (p["LINE_PS_C_LINE_F"] + p["LINE_PS_C_CHA_each_F"] / 2.0)

    z_through_each = z_cap_esr_esl(s, p["C_THROUGH_each_F"], p["C_THROUGH_ESR_each_ohm"], p["C_THROUGH_ESL_each_H"])
    y_shunt_1 = y_line_ps + 1.0 / (2.0 * z_through_each)

    z_line_input = 2.0 * (p["LINE_INPUT_DCR_each_ohm"] + s * p["LINE_INPUT_L_each_H"])
    y_shunt_2 = s * (p["LINE_INPUT_C_LINE_F"] + p["LINE_INPUT_C_CHA_each_F"] / 2.0)

    z_ldm1_dm = 2.0 * z_inductor_with_parallel_cap(s, p["LDM1_L_each_H"], p["LDM1_DCR_each_ohm"], p["LDM1_C_PA_each_F"])

    z_cx1 = z_cap_esr_esl(s, p["CX1_C_F"], p["CX1_ESR_ohm"], p["CX1_ESL_H"])
    z_dump1 = p["DUMP1_R_ohm"] + z_cap_esr_esl(s, p["DUMP1_C_F"], p["DUMP1_ESR_ohm"], p["DUMP1_ESL_H"])
    y_shunt_3 = 1.0 / z_cx1 + 1.0 / z_dump1

    z_ldm2_dm = 2.0 * z_inductor_with_parallel_cap(s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"])

    z_cx2 = z_cap_esr_esl(s, p["CX2_C_F"], p["CX2_ESR_ohm"], p["CX2_ESL_H"])
    z_dump2 = p["DUMP2_R_ohm"] + z_cap_esr_esl(s, p["DUMP2_C_F"], p["DUMP2_ESR_ohm"], p["DUMP2_ESL_H"])
    y_shunt_4 = 1.0 / z_cx2 + 1.0 / z_dump2

    z_cmc_dm = 2.0 * z_inductor_with_parallel_cap(
        s,
        p["CMC_L_each_H"] * (1.0 - p["CMC_K"]),
        p["CMC_DCR_each_ohm"],
        p["CMC_C_PA_each_F"],
    )

    z_cy_each = z_cap_esr_esl(s, p["CY_each_F"], p["CY_ESR_each_ohm"], p["CY_ESL_each_H"])
    y_shunt_5 = s * p["CMC_C_LINE_F"] + 1.0 / (2.0 * z_cy_each) + 1.0 / p["DUMMY_R_ohm"]

    # Source side AC source is a short for this noise-current analysis.
    y_left_1 = y_shunt_1 + 1.0 / z_line_ps

    yb2 = 1.0 / (z_line_input + 1.0 / y_left_1)
    y_left_2 = y_shunt_2 + yb2

    yb3 = 1.0 / (z_ldm1_dm + 1.0 / y_left_2)
    y_left_3 = y_shunt_3 + yb3

    yb4 = 1.0 / (z_ldm2_dm + 1.0 / y_left_3)
    y_left_4 = y_shunt_4 + yb4

    yb5 = 1.0 / (z_cmc_dm + 1.0 / y_left_4)
    y_left_5 = y_shunt_5 + yb5

    v5 = p["NM_I_source_A"] / y_left_5

    i_cmc = v5 * yb5
    v4 = i_cmc / y_left_4

    i_ldm2 = v4 * yb4
    v3 = i_ldm2 / y_left_3

    i_ldm1_nm = v3 * yb3

    i_hot = i_ldm1_nm
    i_rtn = -i_ldm1_nm
    i_nm = (i_hot - i_rtn) / 2.0

    return {"I_HOT_A": i_hot, "I_RTN_A": i_rtn, "I_NM_A": i_nm}


# ============================================================
# Save functions
# ============================================================

def save_csv(path, headers, data_arrays):
    data = np.column_stack(data_arrays)
    np.savetxt(str(path), data, delimiter=",", header=",".join(headers), comments="")


def make_plots(out_dir, f, fr, cm, nm):
    # ① frequency response
    plt.figure(figsize=(10, 6))
    plt.semilogx(f, 20.0 * np.log10(np.abs(fr["H_Rs0"])), label="ROUT = 0 ohm")
    plt.semilogx(f, 20.0 * np.log10(np.abs(fr["H_Rs50"])), label="ROUT = 50 ohm")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("|Vload/Vsource| [dB]")
    plt.title("① Frequency response from test system to load")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "01_frequency_response.png", dpi=180)
    plt.close()

    # ② CM current
    plt.figure(figsize=(10, 6))
    plt.semilogx(f, to_dbuA(cm["I_HOT_A"]), label="I_HOT")
    plt.semilogx(f, to_dbuA(cm["I_RTN_A"]), label="I_RTN")
    plt.semilogx(f, to_dbuA(cm["I_CM_A"]), label="I_CM = I_HOT + I_RTN")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA]")
    plt.title("② Common-mode current from load side")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "02_cm_current.png", dpi=180)
    plt.close()

    plt.figure(figsize=(10, 5))
    plt.semilogx(f, np.angle(cm["I_CM_A"], deg=True), label="I_CM phase")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Phase [deg]")
    plt.title("② Common-mode current phase")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "02_cm_current_phase.png", dpi=180)
    plt.close()

    # ③ NM current
    plt.figure(figsize=(10, 6))
    plt.semilogx(f, to_dbuA(nm["I_HOT_A"]), label="I_HOT")
    plt.semilogx(f, to_dbuA(nm["I_RTN_A"]), label="I_RTN")
    plt.semilogx(f, to_dbuA(nm["I_NM_A"]), label="I_NM = (I_HOT - I_RTN)/2")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA]")
    plt.title("③ Normal-mode current from load side")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "03_nm_current.png", dpi=180)
    plt.close()

    plt.figure(figsize=(10, 5))
    plt.semilogx(f, np.angle(nm["I_NM_A"], deg=True), label="I_NM phase")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Phase [deg]")
    plt.title("③ Normal-mode current phase")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "03_nm_current_phase.png", dpi=180)
    plt.close()


def run_all(out_dir=None, p=PARAMS):
    if out_dir is None:
        out_dir = Path(__file__).resolve().parent
    else:
        out_dir = Path(out_dir)

    f = np.logspace(np.log10(p["f_start_Hz"]), np.log10(p["f_stop_Hz"]), int(p["n_points"]))

    fr = {
        "H_Rs0": calc_frequency_response(f, Rs_ohm=0.0, p=p),
        "H_Rs50": calc_frequency_response(f, Rs_ohm=50.0, p=p),
    }
    cm = calc_cm_current(f, p=p)
    nm = calc_nm_current(f, p=p)

    # Individual CSVs
    save_csv(
        out_dir / "01_frequency_response.csv",
        ["frequency_Hz", "gain_Rs0_dB", "gain_Rs50_dB", "abs_H_Rs0", "abs_H_Rs50"],
        [f, 20*np.log10(np.abs(fr["H_Rs0"])), 20*np.log10(np.abs(fr["H_Rs50"])), np.abs(fr["H_Rs0"]), np.abs(fr["H_Rs50"])],
    )

    save_csv(
        out_dir / "02_cm_current.csv",
        ["frequency_Hz", "I_HOT_abs_A", "I_RTN_abs_A", "I_CM_abs_A", "I_HOT_dBuA", "I_RTN_dBuA", "I_CM_dBuA", "I_CM_phase_deg"],
        [f, np.abs(cm["I_HOT_A"]), np.abs(cm["I_RTN_A"]), np.abs(cm["I_CM_A"]), to_dbuA(cm["I_HOT_A"]), to_dbuA(cm["I_RTN_A"]), to_dbuA(cm["I_CM_A"]), np.angle(cm["I_CM_A"], deg=True)],
    )

    save_csv(
        out_dir / "03_nm_current.csv",
        ["frequency_Hz", "I_HOT_abs_A", "I_RTN_abs_A", "I_NM_abs_A", "I_HOT_dBuA", "I_RTN_dBuA", "I_NM_dBuA", "I_NM_phase_deg"],
        [f, np.abs(nm["I_HOT_A"]), np.abs(nm["I_RTN_A"]), np.abs(nm["I_NM_A"]), to_dbuA(nm["I_HOT_A"]), to_dbuA(nm["I_RTN_A"]), to_dbuA(nm["I_NM_A"]), np.angle(nm["I_NM_A"], deg=True)],
    )

    # Combined CSV
    save_csv(
        out_dir / "combined_three_models_results.csv",
        [
            "frequency_Hz",
            "FR_gain_Rs0_dB", "FR_gain_Rs50_dB",
            "CM_I_HOT_dBuA", "CM_I_RTN_dBuA", "CM_I_CM_dBuA",
            "NM_I_HOT_dBuA", "NM_I_RTN_dBuA", "NM_I_NM_dBuA",
            "CM_I_CM_abs_A", "NM_I_NM_abs_A",
        ],
        [
            f,
            20*np.log10(np.abs(fr["H_Rs0"])),
            20*np.log10(np.abs(fr["H_Rs50"])),
            to_dbuA(cm["I_HOT_A"]),
            to_dbuA(cm["I_RTN_A"]),
            to_dbuA(cm["I_CM_A"]),
            to_dbuA(nm["I_HOT_A"]),
            to_dbuA(nm["I_RTN_A"]),
            to_dbuA(nm["I_NM_A"]),
            np.abs(cm["I_CM_A"]),
            np.abs(nm["I_NM_A"]),
        ],
    )

    make_plots(out_dir, f, fr, cm, nm)

    return f, fr, cm, nm


def print_representative_points(f, fr, cm, nm):
    print("==== Representative points ====")
    print("freq_Hz, FR_Rs0_dB, FR_Rs50_dB, CM_I_CM_A, CM_I_CM_dBuA, NM_I_NM_A, NM_I_NM_dBuA")
    for ff in [100, 400, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8]:
        i = int(np.argmin(np.abs(f - ff)))
        print(
            f"{f[i]:.6g}, "
            f"{20*np.log10(abs(fr['H_Rs0'][i])):.6f}, "
            f"{20*np.log10(abs(fr['H_Rs50'][i])):.6f}, "
            f"{abs(cm['I_CM_A'][i]):.9g}, "
            f"{to_dbuA(cm['I_CM_A'][i]):.6f}, "
            f"{abs(nm['I_NM_A'][i]):.9g}, "
            f"{to_dbuA(nm['I_NM_A'][i]):.6f}"
        )


def main():
    out_dir = Path(__file__).resolve().parent
    f, fr, cm, nm = run_all(out_dir)
    print("Outputs saved in:", out_dir)
    print("  01_frequency_response.csv / .png")
    print("  02_cm_current.csv / .png / _phase.png")
    print("  03_nm_current.csv / .png / _phase.png")
    print("  combined_three_models_results.csv")
    print()
    print_representative_points(f, fr, cm, nm)


if __name__ == "__main__":
    main()
