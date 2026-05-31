#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Check correspondence between fixed LTspice RAW result and the Python DM model.

Input files expected in the same folder:
    461C_FILTER(2).raw
    test_system_emi_filter_dm_response.py

Output files:
    ltspice_dcr_corrected_vs_python_dm_overlay.png
    ltspice_dcr_corrected_vs_python_dm_error.png
    ltspice_dcr_corrected_vs_python_dm_comparison.csv
    ltspice_dcr_corrected_vs_python_dm_summary.csv
"""
from pathlib import Path
import csv
import re
import importlib.util
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


def read_ltspice_raw(path):
    raw_path = Path(path)
    b = raw_path.read_bytes()
    marker = "Binary:".encode("utf-16le")
    idx = b.find(marker)
    if idx < 0:
        raise ValueError("Binary marker not found. This parser expects LTspice UTF-16 binary RAW.")

    header = b[:idx + len(marker)].decode("utf-16le", errors="ignore")
    start = idx + len(marker)
    while b[start:start + 2] in (b"\r\x00", b"\n\x00"):
        start += 2

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

    expected = start + npoints * nvars * 2 * 8
    if expected != len(b):
        raise ValueError(f"Unexpected RAW size. expected={expected}, actual={len(b)}")

    arr = np.frombuffer(b, dtype="<f8", count=npoints * nvars * 2, offset=start)
    vals = arr.reshape(npoints, nvars, 2)
    vals = vals[:, :, 0] + 1j * vals[:, :, 1]
    freq = vals[:, 0].real
    return header, names, freq, vals


def load_python_model(model_path):
    spec = importlib.util.spec_from_file_location("dm_model", model_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    folder = Path(__file__).resolve().parent
    raw_path = folder / "461C_FILTER(2).raw"
    model_path = folder / "test_system_emi_filter_dm_response.py"

    header, names, freq, vals = read_ltspice_raw(raw_path)
    idx = {name: i for i, name in enumerate(names)}

    required = ["V(load_1)", "V(load_2)", "V(ps_1)", "V(ps_2)"]
    missing = [name for name in required if name not in idx]
    if missing:
        raise KeyError(f"Required traces are missing in RAW: {missing}")

    vload = vals[:, idx["V(load_1)"]] - vals[:, idx["V(load_2)"]]
    vsrc = vals[:, idx["V(ps_1)"]] - vals[:, idx["V(ps_2)"]]
    h_lt = vload / vsrc

    model = load_python_model(model_path)
    h_py = model.calc_dm_response(freq, Rs=0.0)

    db_lt = 20.0 * np.log10(np.abs(h_lt))
    db_py = 20.0 * np.log10(np.abs(h_py))
    err = db_py - db_lt

    comparison_csv = folder / "ltspice_dcr_corrected_vs_python_dm_comparison.csv"
    summary_csv = folder / "ltspice_dcr_corrected_vs_python_dm_summary.csv"
    overlay_png = folder / "ltspice_dcr_corrected_vs_python_dm_overlay.png"
    error_png = folder / "ltspice_dcr_corrected_vs_python_dm_error.png"

    with comparison_csv.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            "frequency_Hz",
            "ltspice_gain_dB",
            "python_gain_dB",
            "error_python_minus_ltspice_dB",
            "abs_H_ltspice",
            "abs_H_python",
        ])
        for row in zip(freq, db_lt, db_py, err, np.abs(h_lt), np.abs(h_py)):
            w.writerow(row)

    with summary_csv.open("w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["metric", "value"])
        w.writerow(["n_points", len(freq)])
        w.writerow(["max_abs_error_dB", float(np.max(np.abs(err)))])
        w.writerow(["rms_error_dB", float(np.sqrt(np.mean(err ** 2)))])
        imax = int(np.argmax(np.abs(err)))
        w.writerow(["max_error_frequency_Hz", float(freq[imax])])
        w.writerow(["ltspice_gain_at_max_error_dB", float(db_lt[imax])])
        w.writerow(["python_gain_at_max_error_dB", float(db_py[imax])])
        for ff in [100, 400, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8]:
            i = int(np.argmin(np.abs(freq - ff)))
            w.writerow([f"ltspice_gain_{ff:g}_Hz_dB", float(db_lt[i])])
            w.writerow([f"python_gain_{ff:g}_Hz_dB", float(db_py[i])])
            w.writerow([f"error_{ff:g}_Hz_dB", float(err[i])])

    plt.figure(figsize=(10, 6))
    plt.semilogx(freq, db_lt, label="LTspice RAW: V(load_1,load_2)/V(ps_1,ps_2)")
    plt.semilogx(freq, db_py, "--", label="Python model: ROUT = 0 ohm")
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("|Vload/Vsource| [dB]")
    plt.title("DCR-corrected LTspice vs Python DM model")
    plt.legend()
    plt.tight_layout()
    plt.savefig(overlay_png, dpi=180)

    plt.figure(figsize=(10, 5))
    plt.semilogx(freq, err)
    plt.grid(True, which="both")
    plt.xlabel("Frequency [Hz]")
    plt.ylabel("Python - LTspice [dB]")
    plt.title("Difference of DM frequency response")
    plt.tight_layout()
    plt.savefig(error_png, dpi=180)

    print("RAW points:", len(freq))
    print("Max abs error [dB]:", float(np.max(np.abs(err))))
    print("RMS error [dB]:", float(np.sqrt(np.mean(err ** 2))))
    print("Overlay:", overlay_png)
    print("Error plot:", error_png)
    print("Comparison CSV:", comparison_csv)
    print("Summary CSV:", summary_csv)


if __name__ == "__main__":
    main()
