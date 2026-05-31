#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Normal-mode current-source model for the AC115V EMI filter.

Noise source:
    One current source between DUMMY HOT and DUMMY RTN, amplitude = 1 A.

Measurement plane:
    LDM1 HOT/RTN input side, i.e. interface between test system and EMI filter.

Definition:
    I_NM = (I_HOT - I_RTN) / 2

Outputs are saved in the same folder as this script:
    nm_current_source_python_results.csv
    nm_current_source_python_current.png
    nm_current_source_python_phase.png

If 461C_FILTER(4).raw is in the same folder, the script also creates:
    nm_ltspice_vs_python_current_comparison.csv
    nm_ltspice_vs_python_current_summary.csv
    nm_ltspice_vs_python_current_overlay.png
    nm_ltspice_vs_python_current_error.png
"""

from pathlib import Path
import re
import csv
import numpy as np
import matplotlib.pyplot as plt

PARAMS = {
    "f_start_Hz": 100.0,
    "f_stop_Hz": 100e6,
    "n_points": 1201,
    "I_DM_SOURCE_A": 1.0,
    "LINE_PS_L_each_H": 1e-6,
    "LINE_PS_DCR_each_ohm": 20e-3,
    "LINE_PS_C_CHA_each_F": 50e-12,
    "LINE_PS_C_LINE_F": 50e-12,
    "C_THROUGH_each_F": 10e-6,
    "C_THROUGH_ESR_each_ohm": 20e-3,
    "C_THROUGH_ESL_each_H": 10e-9,
    "LINE_INPUT_L_each_H": 1e-6,
    "LINE_INPUT_DCR_each_ohm": 20e-3,
    "LINE_INPUT_C_CHA_each_F": 50e-12,
    "LINE_INPUT_C_LINE_F": 50e-12,
    "LDM1_L_each_H": 68e-6,
    "LDM1_DCR_each_ohm": 20e-3,
    "LDM1_C_PA_each_F": 5e-12,
    "CX1_C_F": 0.47e-6,
    "CX1_ESR_ohm": 20e-3,
    "CX1_ESL_H": 10e-9,
    "DUMP1_C_F": 0.47e-6,
    "DUMP1_ESR_ohm": 20e-3,
    "DUMP1_ESL_H": 10e-9,
    "DUMP1_R_ohm": 15.0,
    "LDM2_L_each_H": 33e-6,
    "LDM2_DCR_each_ohm": 20e-3,
    "LDM2_C_PA_each_F": 5e-12,
    "CX2_C_F": 0.47e-6,
    "CX2_ESR_ohm": 20e-3,
    "CX2_ESL_H": 10e-9,
    "DUMP2_C_F": 0.47e-6,
    "DUMP2_ESR_ohm": 20e-3,
    "DUMP2_ESL_H": 10e-9,
    "DUMP2_R_ohm": 15.0,
    "CMC_L_each_H": 3e-3,
    "CMC_DCR_each_ohm": 100e-3,
    "CMC_K": 0.999,
    "CMC_C_PA_each_F": 5e-12,
    "CMC_C_LINE_F": 10e-12,
    "CY_each_F": 2.2e-9,
    "CY_ESR_each_ohm": 20e-3,
    "CY_ESL_each_H": 10e-9,
    "DUMMY_R_ohm": 28.75,
}

def z_cap_esr_esl(s, C, ESR, ESL):
    return ESR + s * ESL + 1.0 / (s * C)

def z_inductor_with_parallel_cap(s, L, R, Cp):
    z_rl = R + s * L
    if Cp <= 0:
        return z_rl
    return 1.0 / (1.0 / z_rl + s * Cp)

def to_dbuA(i_A):
    return 20.0 * np.log10(np.maximum(np.abs(i_A), 1e-300) * 1e6)

def calculate_nm_currents(f_Hz, p=PARAMS):
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
    z_dump1_c = z_cap_esr_esl(s, p["DUMP1_C_F"], p["DUMP1_ESR_ohm"], p["DUMP1_ESL_H"])
    y_shunt_3 = 1.0 / z_cx1 + 1.0 / (p["DUMP1_R_ohm"] + z_dump1_c)

    z_ldm2_dm = 2.0 * z_inductor_with_parallel_cap(s, p["LDM2_L_each_H"], p["LDM2_DCR_each_ohm"], p["LDM2_C_PA_each_F"])
    z_cx2 = z_cap_esr_esl(s, p["CX2_C_F"], p["CX2_ESR_ohm"], p["CX2_ESL_H"])
    z_dump2_c = z_cap_esr_esl(s, p["DUMP2_C_F"], p["DUMP2_ESR_ohm"], p["DUMP2_ESL_H"])
    y_shunt_4 = 1.0 / z_cx2 + 1.0 / (p["DUMP2_R_ohm"] + z_dump2_c)

    l_cmc_leak_each = p["CMC_L_each_H"] * (1.0 - p["CMC_K"])
    z_cmc_dm = 2.0 * z_inductor_with_parallel_cap(s, l_cmc_leak_each, p["CMC_DCR_each_ohm"], p["CMC_C_PA_each_F"])

    z_cy_each = z_cap_esr_esl(s, p["CY_each_F"], p["CY_ESR_each_ohm"], p["CY_ESL_each_H"])
    y_shunt_5 = s * p["CMC_C_LINE_F"] + 1.0 / (2.0 * z_cy_each) + 1.0 / p["DUMMY_R_ohm"]

    y_left_1 = y_shunt_1 + 1.0 / z_line_ps
    y_branch_2_to_1 = 1.0 / (z_line_input + 1.0 / y_left_1)
    y_left_2 = y_shunt_2 + y_branch_2_to_1
    y_branch_3_to_2 = 1.0 / (z_ldm1_dm + 1.0 / y_left_2)
    y_left_3 = y_shunt_3 + y_branch_3_to_2
    y_branch_4_to_3 = 1.0 / (z_ldm2_dm + 1.0 / y_left_3)
    y_left_4 = y_shunt_4 + y_branch_4_to_3
    y_branch_5_to_4 = 1.0 / (z_cmc_dm + 1.0 / y_left_4)
    y_left_5 = y_shunt_5 + y_branch_5_to_4

    v5 = p["I_DM_SOURCE_A"] / y_left_5
    i_cmc_dm = v5 * y_branch_5_to_4
    v4 = i_cmc_dm / y_left_4
    i_ldm2_dm = v4 * y_branch_4_to_3
    v3 = i_ldm2_dm / y_left_3
    i_ldm1_dm = v3 * y_branch_3_to_2

    i_hot = i_ldm1_dm
    i_rtn = -i_ldm1_dm
    i_nm = (i_hot - i_rtn) / 2.0

    return {"frequency_Hz": f_Hz, "I_HOT_A": i_hot, "I_RTN_A": i_rtn, "I_NM_A": i_nm}

def read_ltspice_raw(path: Path):
    b = path.read_bytes()
    marker = "Binary:".encode("utf-16le")
    idx = b.find(marker)
    enc = "utf-16le"
    if idx < 0:
        marker = b"Binary:"
        idx = b.find(marker)
        enc = "latin1"
    if idx < 0:
        raise ValueError("Binary marker not found.")
    header = b[:idx + len(marker)].decode(enc, errors="ignore")
    start = idx + len(marker)
    if enc == "utf-16le":
        while b[start:start + 2] in (b"\r\x00", b"\n\x00"):
            start += 2
    else:
        while b[start:start + 1] in (b"\r", b"\n"):
            start += 1
    nvars = int(re.search(r"No\. Variables:\s*(\d+)", header).group(1))
    npoints = int(re.search(r"No\. Points:\s*(\d+)", header).group(1))
    names = []
    in_vars = False
    for line in header.splitlines():
        if line.startswith("Variables:"):
            in_vars = True
            continue
        if in_vars and line.strip():
            parts = line.split()
            if len(parts) >= 3 and parts[0].isdigit():
                names.append(parts[1])
    arr = np.frombuffer(b, dtype="<f8", count=npoints * nvars * 2, offset=start)
    vals = arr.reshape(npoints, nvars, 2)
    vals = vals[:, :, 0] + 1j * vals[:, :, 1]
    return names, vals[:, 0].real, vals

def save_results(out_dir, results):
    f = results["frequency_Hz"]
    ih = results["I_HOT_A"]
    ir = results["I_RTN_A"]
    inm = results["I_NM_A"]
    csv_path = out_dir / "nm_current_source_python_results.csv"
    data = np.column_stack([
        f, np.abs(ih), np.abs(ir), np.abs(inm),
        to_dbuA(ih), to_dbuA(ir), to_dbuA(inm),
        np.real(ih), np.imag(ih), np.real(ir), np.imag(ir), np.real(inm), np.imag(inm),
        np.angle(ih, deg=True), np.angle(ir, deg=True), np.angle(inm, deg=True),
    ])
    np.savetxt(csv_path, data, delimiter=",", header="frequency_Hz,I_HOT_abs_A,I_RTN_abs_A,I_NM_abs_A,I_HOT_dBuA,I_RTN_dBuA,I_NM_dBuA,I_HOT_real_A,I_HOT_imag_A,I_RTN_real_A,I_RTN_imag_A,I_NM_real_A,I_NM_imag_A,I_HOT_phase_deg,I_RTN_phase_deg,I_NM_phase_deg", comments="")

    plt.figure(figsize=(10, 6))
    plt.semilogx(f, to_dbuA(ih), label="I_HOT")
    plt.semilogx(f, to_dbuA(ir), label="I_RTN")
    plt.semilogx(f, to_dbuA(inm), label="I_NM=(I_HOT-I_RTN)/2")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA]")
    plt.title("Normal-mode current at LDM1 input plane")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "nm_current_source_python_current.png", dpi=180)
    plt.close()

    plt.figure(figsize=(10, 5))
    plt.semilogx(f, np.angle(ih, deg=True), label="I_HOT phase")
    plt.semilogx(f, np.angle(ir, deg=True), label="I_RTN phase")
    plt.semilogx(f, np.angle(inm, deg=True), label="I_NM phase")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Phase [deg]")
    plt.title("Normal-mode current phase")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "nm_current_source_python_phase.png", dpi=180)
    plt.close()

def compare_with_raw(out_dir):
    raw_path = out_dir / "461C_FILTER(4).raw"
    if not raw_path.exists():
        return None

    names, freq, vals = read_ltspice_raw(raw_path)
    idx = {name: i for i, name in enumerate(names)}
    ih_lt = vals[:, idx["I(VPROBE_HOT)"]]
    ir_lt = vals[:, idx["I(VPROBE_RTN)"]]
    inm_lt = (ih_lt - ir_lt) / 2.0

    res = calculate_nm_currents(freq)
    ih_py = res["I_HOT_A"]
    ir_py = res["I_RTN_A"]
    inm_py = res["I_NM_A"]
    err_dB = to_dbuA(inm_py) - to_dbuA(inm_lt)

    with (out_dir / "nm_ltspice_vs_python_current_comparison.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["frequency_Hz","LT_I_HOT_abs_A","LT_I_RTN_abs_A","LT_I_NM_abs_A","PY_I_HOT_abs_A","PY_I_RTN_abs_A","PY_I_NM_abs_A","LT_I_NM_dBuA","PY_I_NM_dBuA","error_NM_dB"])
        for row in zip(freq, np.abs(ih_lt), np.abs(ir_lt), np.abs(inm_lt), np.abs(ih_py), np.abs(ir_py), np.abs(inm_py), to_dbuA(inm_lt), to_dbuA(inm_py), err_dB):
            w.writerow(row)

    with (out_dir / "nm_ltspice_vs_python_current_summary.csv").open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["metric","value"])
        w.writerow(["raw_points", len(freq)])
        w.writerow(["max_abs_NM_error_dB", float(np.max(np.abs(err_dB)))])
        w.writerow(["rms_NM_error_dB", float(np.sqrt(np.mean(err_dB**2)))])

    plt.figure(figsize=(10, 6))
    plt.semilogx(freq, to_dbuA(ih_lt), label="LTspice I_HOT")
    plt.semilogx(freq, to_dbuA(ir_lt), label="LTspice I_RTN")
    plt.semilogx(freq, to_dbuA(inm_lt), label="LTspice I_NM")
    plt.semilogx(freq, to_dbuA(inm_py), "--", label="Python I_NM")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Current [dBµA]")
    plt.title("Normal-mode current: LTspice vs Python")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_dir / "nm_ltspice_vs_python_current_overlay.png", dpi=180)
    plt.close()

    plt.figure(figsize=(10, 5))
    plt.semilogx(freq, err_dB)
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Python - LTspice [dB]")
    plt.title("Normal-mode current difference")
    plt.tight_layout()
    plt.savefig(out_dir / "nm_ltspice_vs_python_current_error.png", dpi=180)
    plt.close()

    return float(np.max(np.abs(err_dB))), float(np.sqrt(np.mean(err_dB**2)))

def main():
    out_dir = Path(__file__).resolve().parent
    f = np.logspace(np.log10(PARAMS["f_start_Hz"]), np.log10(PARAMS["f_stop_Hz"]), int(PARAMS["n_points"]))
    results = calculate_nm_currents(f)
    save_results(out_dir, results)
    cmp_result = compare_with_raw(out_dir)

    print("Output: nm_current_source_python_results.csv")
    print("Output: nm_current_source_python_current.png")
    print("Output: nm_current_source_python_phase.png")
    if cmp_result is not None:
        print("LTspice comparison max/RMS error [dB]:", cmp_result)

    for ff in [100, 400, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8]:
        i = int(np.argmin(np.abs(results["frequency_Hz"] - ff)))
        print(f"{results['frequency_Hz'][i]:.6g} Hz: |I_HOT|={abs(results['I_HOT_A'][i]):.9g} A, |I_RTN|={abs(results['I_RTN_A'][i]):.9g} A, |I_NM|={abs(results['I_NM_A'][i]):.9g} A, I_NM={to_dbuA(results['I_NM_A'][i]):.6f} dBµA")

if __name__ == "__main__":
    main()
