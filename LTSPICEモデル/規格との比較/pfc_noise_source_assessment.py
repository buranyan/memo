#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PFC switching noise source assessment for the EMI filter.

This script estimates DM/CM conducted-noise harmonics at the LDM1 input plane
using the previously validated EMI filter models.

Cases:
    DM noise source:
        Current source between DUMMY HOT and DUMMY RTN.
        Default waveform: rectangular current, 1 A p-p, fsw=125 kHz, D=0.5.

    CM noise source:
        Two in-phase current sources:
            DUMMY HOT -> CHASSIS
            DUMMY RTN -> CHASSIS
        Default total common-mode current: 20 mA p-p total.
        In the LTspice implementation this is split equally:
            each source = Icm_total_pp / 2 p-p
            PULSE(-Icm_total_pp/4, +Icm_total_pp/4, ...)

Outputs:
    pfc_noise_source_assessment.csv
    pfc_noise_source_assessment_summary.csv
    pfc_noise_source_assessment.png
    ltspice_pfc_noise_sources_snippet.txt

Important:
    The included CE03 limit curve is an engineering approximation and should
    be replaced with the exact project/contract limit curve when available.
"""

from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt


# ============================================================
# Parameters from the user's Excel initial study
# ============================================================

PFC = {
    "Vin_rms_V": 115.0,
    "Po_W": 414.0,
    "eta": 0.9,
    "Pin_W": 460.0,
    "Rin_ohm": 28.75,
    "fsw_Hz": 125e3,
    "D": 0.5,
    "Vbus_V": 325.0,
    "Lpfc_H": 200e-6,
    "Cpar_F": 80e-12,

    # Rectangular source settings
    "Idm_pp_A": 1.0,

    # Case A: current Excel setting, total common-mode current p-p
    "Icm_total_pp_A": 0.02,

    # Edge time used in LTspice snippets. 0 makes an ideal rectangular spectrum.
    "tr_s": 50e-9,
    "tf_s": 50e-9,
}


FILTER = {
    # Source/load
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
}


# ============================================================
# Basic functions
# ============================================================

def z_cap_esr_esl(s, C, ESR, ESL):
    return ESR + s * ESL + 1.0 / (s * C)


def z_inductor_with_parallel_cap(s, L, R, Cp):
    z_rl = R + s * L
    if Cp <= 0:
        return z_rl
    return 1.0 / (1.0 / z_rl + s * Cp)


def dbua_from_arms(i_rms):
    return 20.0 * np.log10(np.maximum(np.abs(i_rms), 1e-300) * 1e6)


def rectangular_harmonic_rms(Ipp, n, duty=0.5, fsw=125e3, tr=0.0):
    """
    RMS amplitude of the nth harmonic of a rectangular current waveform.

    The waveform is assumed to swing from -Ipp/2 to +Ipp/2.
    For D=0.5, only odd harmonics remain.

    The optional rise/fall-time correction uses a sinc envelope:
        |sin(pi*f*tr)/(pi*f*tr)|
    which is a practical first-order approximation for finite edge times.
    """
    n = np.asarray(n, dtype=float)
    peak = 2.0 * Ipp / (np.pi * n) * np.abs(np.sin(np.pi * n * duty))

    if tr and tr > 0:
        f = n * fsw
        peak = peak * np.abs(np.sinc(f * tr))

    return peak / np.sqrt(2.0)


def ce03_limit_dbua_approx(f_hz, relaxed_10db=False):
    """
    Approximate MIL-STD-461C CE03 narrowband current limit [dBµA].

    This is a digitized engineering approximation for preliminary design.
    Replace CE03_LIMIT_POINTS with the exact project limit if available.

    Points used:
        15 kHz  : 90 dBµA
        625 kHz : 33 dBµA
        2 MHz   : 20 dBµA
        50 MHz  : 20 dBµA

    A +10 dB relaxation option is included for the note appearing on the
    CE03 figure for some class A1 Navy/Air Force procurements.
    """
    points = np.array([
        [15e3, 90.0],
        [625e3, 33.0],
        [2e6, 20.0],
        [50e6, 20.0],
    ])
    f_hz = np.asarray(f_hz, dtype=float)
    limit = np.interp(np.log10(f_hz), np.log10(points[:, 0]), points[:, 1])
    if relaxed_10db:
        limit = limit + 10.0
    return limit


# ============================================================
# EMI filter current-transfer models
# These are the same normalized models as the previously validated tools.
# ============================================================

def calc_cm_current_transfer(f_Hz, p=FILTER):
    """
    Return normalized transfer from total CM current source to measured I_CM.

    Source normalization:
        two in-phase sources, 0.5 A each, total source current = 1 A.

    Return:
        complex I_CM at LDM1 input per 1 A total CM source.
    """
    f_Hz = np.asarray(f_Hz, dtype=float)
    s = 1j * 2.0 * np.pi * f_Hz

    y1 = s * 2.0 * p["LINE_PS_C_CHA_each_F"]
    z_through_each = z_cap_esr_esl(
        s,
        p["C_THROUGH_each_F"],
        p["C_THROUGH_ESR_each_ohm"],
        p["C_THROUGH_ESL_each_H"],
    )
    y1 += 2.0 / z_through_each

    z_line_input_cm = (p["LINE_INPUT_DCR_each_ohm"] + s * p["LINE_INPUT_L_each_H"]) / 2.0
    y2 = s * 2.0 * p["LINE_INPUT_C_CHA_each_F"]

    z_ldm1_cm = z_inductor_with_parallel_cap(
        s, p["LDM1_L_each_H"], p["LDM1_DCR_each_ohm"], p["LDM1_C_PA_each_F"]
    ) / 2.0
    z_ldm2_cm = z_inductor_with_parallel_cap(
        s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"]
    ) / 2.0

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

    i_total = 1.0
    v5 = i_total / y_left_5

    i_cmc = v5 * yb5
    v4 = i_cmc / y_left_4

    i_ldm2 = v4 * yb4
    v3 = i_ldm2 / y_left_3

    i_ldm1_cm = v3 * yb3
    return i_ldm1_cm


def calc_nm_current_transfer(f_Hz, p=FILTER):
    """
    Return normalized transfer from a HOT-RTN DM current source to measured I_NM.

    Source normalization:
        one current source between HOT and RTN, source current = 1 A.

    Return:
        complex I_NM at LDM1 input per 1 A DM source.
    """
    f_Hz = np.asarray(f_Hz, dtype=float)
    s = 1j * 2.0 * np.pi * f_Hz

    z_line_ps = 2.0 * (p["LINE_PS_DCR_each_ohm"] + s * p["LINE_PS_L_each_H"])
    y_line_ps = s * (p["LINE_PS_C_LINE_F"] + p["LINE_PS_C_CHA_each_F"] / 2.0)

    z_through_each = z_cap_esr_esl(
        s,
        p["C_THROUGH_each_F"],
        p["C_THROUGH_ESR_each_ohm"],
        p["C_THROUGH_ESL_each_H"],
    )
    y_shunt_1 = y_line_ps + 1.0 / (2.0 * z_through_each)

    z_line_input = 2.0 * (p["LINE_INPUT_DCR_each_ohm"] + s * p["LINE_INPUT_L_each_H"])
    y_shunt_2 = s * (p["LINE_INPUT_C_LINE_F"] + p["LINE_INPUT_C_CHA_each_F"] / 2.0)

    z_ldm1_dm = 2.0 * z_inductor_with_parallel_cap(
        s, p["LDM1_L_each_H"], p["LDM1_DCR_each_ohm"], p["LDM1_C_PA_each_F"]
    )

    z_cx1 = z_cap_esr_esl(s, p["CX1_C_F"], p["CX1_ESR_ohm"], p["CX1_ESL_H"])
    z_dump1 = p["DUMP1_R_ohm"] + z_cap_esr_esl(
        s, p["DUMP1_C_F"], p["DUMP1_ESR_ohm"], p["DUMP1_ESL_H"]
    )
    y_shunt_3 = 1.0 / z_cx1 + 1.0 / z_dump1

    z_ldm2_dm = 2.0 * z_inductor_with_parallel_cap(
        s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"]
    )

    z_cx2 = z_cap_esr_esl(s, p["CX2_C_F"], p["CX2_ESR_ohm"], p["CX2_ESL_H"])
    z_dump2 = p["DUMP2_R_ohm"] + z_cap_esr_esl(
        s, p["DUMP2_C_F"], p["DUMP2_ESR_ohm"], p["DUMP2_ESL_H"]
    )
    y_shunt_4 = 1.0 / z_cx2 + 1.0 / z_dump2

    z_cmc_dm = 2.0 * z_inductor_with_parallel_cap(
        s,
        p["CMC_L_each_H"] * (1.0 - p["CMC_K"]),
        p["CMC_DCR_each_ohm"],
        p["CMC_C_PA_each_F"],
    )

    z_cy_each = z_cap_esr_esl(s, p["CY_each_F"], p["CY_ESR_each_ohm"], p["CY_ESL_each_H"])
    y_shunt_5 = s * p["CMC_C_LINE_F"] + 1.0 / (2.0 * z_cy_each) + 1.0 / p["DUMMY_R_ohm"]

    y_left_1 = y_shunt_1 + 1.0 / z_line_ps

    yb2 = 1.0 / (z_line_input + 1.0 / y_left_1)
    y_left_2 = y_shunt_2 + yb2

    yb3 = 1.0 / (z_ldm1_dm + 1.0 / y_left_2)
    y_left_3 = y_shunt_3 + yb3

    yb4 = 1.0 / (z_ldm2_dm + 1.0 / y_left_3)
    y_left_4 = y_shunt_4 + yb4

    yb5 = 1.0 / (z_cmc_dm + 1.0 / y_left_4)
    y_left_5 = y_shunt_5 + yb5

    v5 = 1.0 / y_left_5

    i_cmc = v5 * yb5
    v4 = i_cmc / y_left_4

    i_ldm2 = v4 * yb4
    v3 = i_ldm2 / y_left_3

    i_ldm1_nm = v3 * yb3
    return i_ldm1_nm


# ============================================================
# Assessment
# ============================================================

def assess(pfc=PFC, relaxed_10db=False):
    fsw = pfc["fsw_Hz"]
    nmax = int(np.floor(50e6 / fsw))
    n = np.arange(1, nmax + 1)
    f = n * fsw

    tr_eff = max(float(pfc.get("tr_s", 0.0)), float(pfc.get("tf_s", 0.0)))

    dm_src_rms = rectangular_harmonic_rms(
        pfc["Idm_pp_A"], n, duty=pfc["D"], fsw=fsw, tr=tr_eff
    )

    cm_src_rms = rectangular_harmonic_rms(
        pfc["Icm_total_pp_A"], n, duty=pfc["D"], fsw=fsw, tr=tr_eff
    )

    h_nm = calc_nm_current_transfer(f)
    h_cm = calc_cm_current_transfer(f)

    dm_at_probe_rms = dm_src_rms * np.abs(h_nm)
    cm_at_probe_rms = cm_src_rms * np.abs(h_cm)

    dm_dbua = dbua_from_arms(dm_at_probe_rms)
    cm_dbua = dbua_from_arms(cm_at_probe_rms)
    limit = ce03_limit_dbua_approx(f, relaxed_10db=relaxed_10db)

    dm_margin = limit - dm_dbua
    cm_margin = limit - cm_dbua

    return {
        "n": n,
        "frequency_Hz": f,
        "dm_source_rms_A": dm_src_rms,
        "cm_source_total_rms_A": cm_src_rms,
        "dm_transfer_abs": np.abs(h_nm),
        "cm_transfer_abs": np.abs(h_cm),
        "dm_at_LDM1_rms_A": dm_at_probe_rms,
        "cm_at_LDM1_rms_A": cm_at_probe_rms,
        "dm_at_LDM1_dBuA": dm_dbua,
        "cm_at_LDM1_dBuA": cm_dbua,
        "ce03_limit_dBuA": limit,
        "dm_margin_dB": dm_margin,
        "cm_margin_dB": cm_margin,
    }


def write_outputs(out_dir, pfc=PFC):
    out_dir = Path(out_dir)
    result = assess(pfc)

    csv_path = out_dir / "pfc_noise_source_assessment.csv"
    data = np.column_stack([
        result["n"],
        result["frequency_Hz"],
        result["dm_source_rms_A"],
        result["cm_source_total_rms_A"],
        result["dm_transfer_abs"],
        result["cm_transfer_abs"],
        result["dm_at_LDM1_rms_A"],
        result["cm_at_LDM1_rms_A"],
        result["dm_at_LDM1_dBuA"],
        result["cm_at_LDM1_dBuA"],
        result["ce03_limit_dBuA"],
        result["dm_margin_dB"],
        result["cm_margin_dB"],
    ])
    np.savetxt(
        csv_path,
        data,
        delimiter=",",
        header=(
            "harmonic_n,frequency_Hz,"
            "dm_source_rms_A,cm_source_total_rms_A,"
            "dm_transfer_abs,cm_transfer_abs,"
            "dm_at_LDM1_rms_A,cm_at_LDM1_rms_A,"
            "dm_at_LDM1_dBuA,cm_at_LDM1_dBuA,"
            "ce03_limit_dBuA,dm_margin_dB,cm_margin_dB"
        ),
        comments="",
    )

    # Summary
    dm_worst_idx = int(np.argmin(result["dm_margin_dB"]))
    cm_worst_idx = int(np.argmin(result["cm_margin_dB"]))

    summary_path = out_dir / "pfc_noise_source_assessment_summary.csv"
    with summary_path.open("w", encoding="utf-8") as f:
        f.write("item,value\n")
        f.write(f"fsw_Hz,{pfc['fsw_Hz']}\n")
        f.write(f"duty,{pfc['D']}\n")
        f.write(f"Idm_pp_A,{pfc['Idm_pp_A']}\n")
        f.write(f"Icm_total_pp_A,{pfc['Icm_total_pp_A']}\n")
        f.write(f"tr_s,{pfc['tr_s']}\n")
        f.write(f"dm_worst_harmonic,{int(result['n'][dm_worst_idx])}\n")
        f.write(f"dm_worst_frequency_Hz,{result['frequency_Hz'][dm_worst_idx]}\n")
        f.write(f"dm_worst_level_dBuA,{result['dm_at_LDM1_dBuA'][dm_worst_idx]}\n")
        f.write(f"dm_limit_dBuA,{result['ce03_limit_dBuA'][dm_worst_idx]}\n")
        f.write(f"dm_margin_dB,{result['dm_margin_dB'][dm_worst_idx]}\n")
        f.write(f"cm_worst_harmonic,{int(result['n'][cm_worst_idx])}\n")
        f.write(f"cm_worst_frequency_Hz,{result['frequency_Hz'][cm_worst_idx]}\n")
        f.write(f"cm_worst_level_dBuA,{result['cm_at_LDM1_dBuA'][cm_worst_idx]}\n")
        f.write(f"cm_limit_dBuA,{result['ce03_limit_dBuA'][cm_worst_idx]}\n")
        f.write(f"cm_margin_dB,{result['cm_margin_dB'][cm_worst_idx]}\n")

    # Plot
    png_path = out_dir / "pfc_noise_source_assessment.png"
    plt.figure(figsize=(11, 6.5))
    plt.semilogx(result["frequency_Hz"], result["ce03_limit_dBuA"], label="CE03 limit approx.")
    plt.semilogx(result["frequency_Hz"], result["dm_at_LDM1_dBuA"], "o-", markersize=3, label="DM at LDM1 input")
    plt.semilogx(result["frequency_Hz"], result["cm_at_LDM1_dBuA"], "o-", markersize=3, label="CM at LDM1 input")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA RMS]")
    plt.title("PFC switching noise estimate at LDM1 input")
    plt.legend()
    plt.tight_layout()
    plt.savefig(png_path, dpi=180)
    plt.close()

    # LTspice snippet
    snippet_path = out_dir / "ltspice_pfc_noise_sources_snippet.txt"
    icm_sine_pk = 2.0 * np.pi * pfc["fsw_Hz"] * pfc["Cpar_F"] * pfc["Vbus_V"]
    icm_total_pp_for_same_fundamental = 0.5 * np.pi * icm_sine_pk

    snippet_path.write_text(
        f"""* ============================================================
* PFC DM/CM noise source implementation for LTspice
* ============================================================
* Place these sources at the DUMMY/load side of the EMI filter.
*
* DM source:
*   one current source between DUMMY HOT and DUMMY RTN.
*
* CM source:
*   two equal in-phase current sources from HOT/RTN to CHASSIS.
*
* Definitions:
*   I_DM at measurement plane = (I(VPROBE_HOT)-I(VPROBE_RTN))/2
*   I_CM at measurement plane =  I(VPROBE_HOT)+I(VPROBE_RTN)
*
* Recommended transient command example:
*   .tran 0 2m 1m 5n
* Then FFT I(VPROBE_HOT), I(VPROBE_RTN), and calculate DM/CM traces.

.param fsw=125k
.param Tsw={{1/fsw}}
.param D=0.5
.param tr=50n
.param tf=50n

* --- DM source from Excel initial value ---
* Idm_pp = 1 A p-p.  Fundamental RMS for D=0.5 is about 0.450 A RMS.
.param Idm_pp=1
I_DM LOAD_1 LOAD_2 PULSE({{-Idm_pp/2}} {{Idm_pp/2}} 0 {{tr}} {{tf}} {{D*Tsw}} {{Tsw}})

* --- CM source, current Excel setting ---
* Icm_total_pp is total CM p-p current, i.e. sum of HOT and RTN source waveforms.
* Each source is half of the total current.
.param Icm_total_pp=20m
I_CM_H LOAD_1 CHASSIS PULSE({{-Icm_total_pp/4}} {{Icm_total_pp/4}} 0 {{tr}} {{tf}} {{D*Tsw}} {{Tsw}})
I_CM_R LOAD_2 CHASSIS PULSE({{-Icm_total_pp/4}} {{Icm_total_pp/4}} 0 {{tr}} {{tf}} {{D*Tsw}} {{Tsw}})

* --- Optional: match the 2*pi*fsw*Cpar*Vbus sinusoidal-equivalent estimate ---
* From Cpar={pfc['Cpar_F']:.6g} F and Vbus={pfc['Vbus_V']:.6g} V:
* Icm_sine_pk = 2*pi*fsw*Cpar*Vbus = {icm_sine_pk:.6g} A peak
* For a 50% rectangular total CM current source, set:
* Icm_total_pp = (pi/2)*Icm_sine_pk = {icm_total_pp_for_same_fundamental:.6g} A p-p
*
* .param Icm_total_pp={icm_total_pp_for_same_fundamental:.9g}
""",
        encoding="utf-8",
    )

    return result, csv_path, summary_path, png_path, snippet_path


def main():
    out_dir = Path(__file__).resolve().parent
    result, csv_path, summary_path, png_path, snippet_path = write_outputs(out_dir)

    dm_idx = int(np.argmin(result["dm_margin_dB"]))
    cm_idx = int(np.argmin(result["cm_margin_dB"]))

    print("Outputs:")
    print(" ", csv_path)
    print(" ", summary_path)
    print(" ", png_path)
    print(" ", snippet_path)
    print()
    print("Worst margins using the approximate CE03 limit:")
    print(
        f"DM: harmonic {int(result['n'][dm_idx])}, "
        f"f={result['frequency_Hz'][dm_idx]:.6g} Hz, "
        f"level={result['dm_at_LDM1_dBuA'][dm_idx]:.3f} dBµA, "
        f"limit={result['ce03_limit_dBuA'][dm_idx]:.3f} dBµA, "
        f"margin={result['dm_margin_dB'][dm_idx]:.3f} dB"
    )
    print(
        f"CM: harmonic {int(result['n'][cm_idx])}, "
        f"f={result['frequency_Hz'][cm_idx]:.6g} Hz, "
        f"level={result['cm_at_LDM1_dBuA'][cm_idx]:.3f} dBµA, "
        f"limit={result['ce03_limit_dBuA'][cm_idx]:.3f} dBµA, "
        f"margin={result['cm_margin_dB'][cm_idx]:.3f} dB"
    )


if __name__ == "__main__":
    main()
