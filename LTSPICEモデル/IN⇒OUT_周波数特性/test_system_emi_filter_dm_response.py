#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_system_emi_filter_dm_response.py

単相 AC115V 400Hz 入力系 + EMIフィルタのディファレンシャルモード周波数特性を計算します。

出力ファイル:
    test_system_emi_filter_dm_response.csv
    test_system_emi_filter_dm_gain.png

出力先:
    この .py ファイルと同じフォルダ

周波数範囲:
    100 Hz ～ 100 MHz

モデル方針:
    HOT/RTN が対称な2線回路として、ディファレンシャルモード等価回路に畳み込みます。

    HOT/RTNに同じ直列素子がある場合:
        Z_DM = Z_H + Z_R = 2 * Z_each

    HOT/RTN間のXコンデンサ、線間寄生容量:
        そのまま線間シャント素子として扱います。

    HOT/RTNからシャーシへ同じコンデンサがある場合:
        ディファレンシャルモードでは2個のコンデンサが直列になるため、
        Z_DM = Z_H_to_CHA + Z_R_to_CHA = 2 * Z_each
        としてHOT-RTN間シャントに畳み込みます。
        低周波近似では C_DM = C_each / 2 です。

    コモンモードチョーク:
        ディファレンシャルモードでは主に漏れインダクタンスが効きます。
        各巻線の漏れインダクタンス:
            L_leak_each = L * (1 - K)
        HOT+RTN合成:
            L_DM = 2 * L * (1 - K)
"""

from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt


# ============================================================
# ABCD utilities
# ============================================================

def abcd_series(Z):
    """Series impedance ABCD matrix."""
    one = np.ones_like(Z, dtype=complex)
    zero = np.zeros_like(Z, dtype=complex)
    return np.array([[one, Z], [zero, one]], dtype=complex)


def abcd_shunt(Y):
    """Shunt admittance ABCD matrix."""
    one = np.ones_like(Y, dtype=complex)
    zero = np.zeros_like(Y, dtype=complex)
    return np.array([[one, zero], [Y, one]], dtype=complex)


def cascade(*matrices):
    """Cascade ABCD matrices. Shape of each matrix is [2, 2, N]."""
    M = matrices[0]
    for N in matrices[1:]:
        A = M[0, 0] * N[0, 0] + M[0, 1] * N[1, 0]
        B = M[0, 0] * N[0, 1] + M[0, 1] * N[1, 1]
        C = M[1, 0] * N[0, 0] + M[1, 1] * N[1, 0]
        D = M[1, 0] * N[0, 1] + M[1, 1] * N[1, 1]
        M = np.array([[A, B], [C, D]], dtype=complex)
    return M


# ============================================================
# Component models
# ============================================================

def z_series_rl(s, L, R):
    """Series R-L impedance."""
    return R + s * L


def z_inductor_parallel_cap(s, L, R, C_pa):
    """
    Inductor model:
        (R + sL) || C_pa
    """
    Z_rl = R + s * L
    if C_pa <= 0:
        return Z_rl
    Y = 1.0 / Z_rl + s * C_pa
    return 1.0 / Y


def z_cap_esr_esl(s, C, ESR, ESL):
    """
    Capacitor model:
        ESR + ESL + ideal C in series
    """
    return ESR + s * ESL + 1.0 / (s * C)


def y_direct_cap(s, C):
    """Admittance of ideal direct capacitance."""
    return s * C


def y_pair_to_chassis_from_cap_impedance(Z_each):
    """
    Two identical capacitors from HOT/RTN to chassis.
    Differential-mode equivalent is series connection of the two impedances.
    """
    return 1.0 / (2.0 * Z_each)


def y_pair_to_chassis_ideal_cap(s, C_each):
    """
    Two identical ideal capacitors from HOT/RTN to chassis.
    Differential-mode equivalent capacitance is C_each/2.
    """
    return s * (C_each / 2.0)


# ============================================================
# Network building blocks
# ============================================================

def line_section_dm(s, L_each, R_each, C_cha_each, C_line):
    """
    Symmetric 2-wire line section, lumped approximation.

    Per conductor:
        series R + L
        capacitance from each line to chassis: C_cha_each

    Between HOT and RTN:
        C_line

    DM equivalent:
        series Z = 2 * (R + sL)
        shunt Y = s*C_line + s*(C_cha_each/2)
    """
    Z_series_dm = 2.0 * z_series_rl(s, L_each, R_each)
    Y_shunt_dm = y_direct_cap(s, C_line) + y_pair_to_chassis_ideal_cap(s, C_cha_each)
    return [
        abcd_series(Z_series_dm),
        abcd_shunt(Y_shunt_dm),
    ]


def shunt_cap_pair_to_chassis(s, C_each, ESR_each, ESL_each):
    """Two equal capacitors from HOT/RTN to chassis, DM equivalent."""
    Z_each = z_cap_esr_esl(s, C_each, ESR_each, ESL_each)
    Y_dm = y_pair_to_chassis_from_cap_impedance(Z_each)
    return abcd_shunt(Y_dm)


def shunt_xcap_with_series_rc_damper(s, Cx, Cx_ESR, Cx_ESL, Cd, Cd_ESR, Cd_ESL, Rd):
    """
    X capacitor in parallel with series RC damper branch.

    X cap branch:
        ESR + ESL + Cx

    Damper branch:
        Rd + ESR + ESL + Cd
    """
    Z_x = z_cap_esr_esl(s, Cx, Cx_ESR, Cx_ESL)
    Z_d = Rd + z_cap_esr_esl(s, Cd, Cd_ESR, Cd_ESL)
    Y_total = 1.0 / Z_x + 1.0 / Z_d
    return abcd_shunt(Y_total)


def dm_coil_pair(s, L_each, R_each, C_pa_each):
    """
    HOT/RTNに同じDMコイルがある場合の等価直列素子。
    各コイルは (R+sL)||Cpa。
    """
    Z_each = z_inductor_parallel_cap(s, L_each, R_each, C_pa_each)
    return abcd_series(2.0 * Z_each)


def cmc_dm_section(s, L_each, R_each, K, C_pa_each, C_line):
    """
    Common-mode choke in differential-mode.

    Each winding:
        leakage inductance = L_each * (1 - K)
        DCR = R_each
        parasitic winding capacitance C_pa_each in parallel

    HOT-RTN parasitic capacitance:
        C_line as shunt line-to-line capacitance
    """
    L_leak_each = L_each * (1.0 - K)
    Z_each = z_inductor_parallel_cap(s, L_leak_each, R_each, C_pa_each)
    Z_dm = 2.0 * Z_each
    Y_line = y_direct_cap(s, C_line)
    return [
        abcd_series(Z_dm),
        abcd_shunt(Y_line),
    ]


# ============================================================
# Response calculation
# ============================================================

def calc_dm_response(f, Rs):
    """
    Calculate Vload/Vsource for the whole test system + EMI filter.
    """
    w = 2.0 * np.pi * f
    s = 1j * w

    # Source/load
    Rload = 28.75

    sections = []

    # --------------------------------------------------------
    # Test system: AC source -> power line
    # LINE_PS_H/R:
    # L=1uH, DCR=20mΩ, C_CHA=50pF, C_LINE=50pF
    # --------------------------------------------------------
    sections += line_section_dm(
        s,
        L_each=1e-6,
        R_each=20e-3,
        C_cha_each=50e-12,
        C_line=50e-12,
    )

    # --------------------------------------------------------
    # Feedthrough capacitors
    # C_THROUGH_H/R: C=10uF, L=10nH, ESR=20mΩ
    # Each line to chassis. DM equivalent is two branches in series.
    # --------------------------------------------------------
    sections.append(
        shunt_cap_pair_to_chassis(
            s,
            C_each=10e-6,
            ESR_each=20e-3,
            ESL_each=10e-9,
        )
    )

    # --------------------------------------------------------
    # Test system: input line
    # LINE_INPUT_H/R:
    # L=1uH, DCR=20mΩ, C_CHA=50pF, C_LINE=50pF
    # --------------------------------------------------------
    sections += line_section_dm(
        s,
        L_each=1e-6,
        R_each=20e-3,
        C_cha_each=50e-12,
        C_line=50e-12,
    )

    # ========================================================
    # Input EMI filter
    # ========================================================

    # 1st DM coils
    sections.append(
        dm_coil_pair(
            s,
            L_each=68e-6,
            R_each=20e-3,
            C_pa_each=5e-12,
        )
    )

    # CX1 and CR_DUMP1 in parallel
    sections.append(
        shunt_xcap_with_series_rc_damper(
            s,
            Cx=0.47e-6,
            Cx_ESR=20e-3,
            Cx_ESL=10e-9,
            Cd=0.47e-6,
            Cd_ESR=20e-3,
            Cd_ESL=10e-9,
            Rd=15.0,
        )
    )

    # 2nd DM coils
    # Note: parameter file says "DRR=20mΩ" for LDM2_R.
    # It is interpreted as DCR=20mΩ.
    sections.append(
        dm_coil_pair(
            s,
            L_each=33e-6,
            R_each=20e-3,
            C_pa_each=5e-12,
        )
    )

    # CX2 and CR_DUMP2 in parallel
    sections.append(
        shunt_xcap_with_series_rc_damper(
            s,
            Cx=0.47e-6,
            Cx_ESR=20e-3,
            Cx_ESL=10e-9,
            Cd=0.47e-6,
            Cd_ESR=20e-3,
            Cd_ESL=10e-9,
            Rd=15.0,
        )
    )

    # CMC
    sections += cmc_dm_section(
        s,
        L_each=3e-3,
        R_each=100e-3,
        K=0.999,
        C_pa_each=5e-12,
        C_line=10e-12,
    )

    # Y capacitors
    sections.append(
        shunt_cap_pair_to_chassis(
            s,
            C_each=2.2e-9,
            ESR_each=20e-3,
            ESL_each=10e-9,
        )
    )

    M = cascade(*sections)

    A = M[0, 0]
    B = M[0, 1]
    C = M[1, 0]
    D = M[1, 1]

    # With load:
    # I2 = V2/Rload
    #
    # With source resistance:
    # Vs = V1 + Rs*I1
    #
    # ABCD:
    # V1 = A*V2 + B*I2
    # I1 = C*V2 + D*I2
    #
    # Therefore:
    # V2/Vs = 1 / (A + B/Rload + Rs*(C + D/Rload))
    H = 1.0 / (A + B / Rload + Rs * (C + D / Rload))

    return H


def get_output_dir():
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd()


def main():
    output_dir = get_output_dir()

    f = np.logspace(2, 8, 5000)

    H_Rs0 = calc_dm_response(f, Rs=0.0)
    H_Rs50 = calc_dm_response(f, Rs=50.0)

    gain_Rs0_dB = 20.0 * np.log10(np.abs(H_Rs0))
    gain_Rs50_dB = 20.0 * np.log10(np.abs(H_Rs50))

    csv_path = output_dir / "test_system_emi_filter_dm_response.csv"
    png_path = output_dir / "test_system_emi_filter_dm_gain.png"

    data = np.column_stack([
        f,
        gain_Rs0_dB,
        gain_Rs50_dB,
        np.abs(H_Rs0),
        np.abs(H_Rs50),
    ])

    np.savetxt(
        str(csv_path),
        data,
        delimiter=",",
        header="frequency_Hz,gain_Rs0_dB,gain_Rs50_dB,abs_H_Rs0,abs_H_Rs50",
        comments="",
    )

    plt.figure(figsize=(10, 6))
    plt.semilogx(f, gain_Rs0_dB, label="Vload/Vsource, ROUT = 0 ohm")
    plt.semilogx(f, gain_Rs50_dB, label="Vload/Vsource, ROUT = 50 ohm")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Voltage transfer |Vload/Vsource| [dB]")
    plt.title("Differential-mode frequency response: test system + EMI filter")
    plt.legend()
    plt.tight_layout()
    plt.savefig(str(png_path), dpi=180)
    plt.show()

    print("==== Output files ====")
    print(f"CSV : {csv_path}")
    print(f"PNG : {png_path}")
    print()

    print("==== Equivalent notes ====")
    print("Load R = 28.75 ohm")
    print("LINE_PS DM series L = 2 uH, DCR = 40 mohm")
    print("LINE_INPUT DM series L = 2 uH, DCR = 40 mohm")
    print("Feedthrough cap DM equivalent at low frequency = 5 uF")
    print("LDM1 DM equivalent at low frequency = 136 uH")
    print("LDM2 DM equivalent at low frequency = 66 uH")
    print("CMC DM leakage at low frequency = 6 uH")
    print("Y-cap DM equivalent at low frequency = 1.1 nF")
    print()

    print("==== Example points ====")
    for ff in [100, 400, 1e3, 10e3, 100e3, 1e6, 10e6, 100e6]:
        h0 = calc_dm_response(np.array([ff]), Rs=0.0)[0]
        h50 = calc_dm_response(np.array([ff]), Rs=50.0)[0]
        print(
            f"{ff:>10.0f} Hz : "
            f"ROUT=0 {20*np.log10(abs(h0)):>9.3f} dB, "
            f"ROUT=50 {20*np.log10(abs(h50)):>9.3f} dB"
        )


if __name__ == "__main__":
    main()
