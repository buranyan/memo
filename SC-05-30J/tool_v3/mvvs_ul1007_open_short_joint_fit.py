#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MVVS 0.5sq 2芯ケーブル
遠端開放・遠端短絡の同時フィッティング

同じ5 mケーブルを、同じ測定治具・校正条件で
  1) 遠端開放
  2) 遠端短絡
した2つのWaveForms CSVを同時に使用し、差動線路定数を推定します。

前提
----
- 2つのCSVは同じケーブル長・治具・校正条件。
- 開放側は理想開放として扱う。
- 短絡側は遠端短絡部のR・Lを含む。
- ケーブルの直流ループ抵抗は、one_core_resistance_ohm_per_m から計算。
- Z0、VF、損失、共通治具R/L、短絡端R/Lを同時フィット。
- 開放端容量まで同時に自由化するとVFとの相関が強く一意性が落ちるため、
  既定では0 pF固定。必要なら --fit-open-capacitance を指定できます。

必要ファイル
------------
同じフォルダに以下を置いてください。
- mvvs_ul1007_mtl_analysis_v2.py
- このスクリプト

実行例
------
python mvvs_ul1007_open_short_joint_fit.py \
    Zm_2C_normal_20260712.csv \
    Zm_2C_normal_short_20260712.csv \
    --length-m 5 \
    --output-dir fit_5m_open_short
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
from dataclasses import asdict, dataclass, replace
from pathlib import Path
from typing import Iterable

import matplotlib.pyplot as plt
from matplotlib import rcParams
import numpy as np

# 日本語ラベル用。環境にない場合はmatplotlibの既定フォントへフォールバックします。
rcParams['font.family'] = ['Noto Sans CJK JP', 'DejaVu Sans']
rcParams['axes.unicode_minus'] = False
from scipy.optimize import least_squares

try:
    from mvvs_ul1007_mtl_analysis_v2 import (
        AnalysisParameters,
        C0,
        differential_line_model,
        read_measured_impedance_csv,
    )
except ImportError as exc:
    raise SystemExit(
        "mvvs_ul1007_mtl_analysis_v2.py を同じフォルダに置いてください。"
    ) from exc


@dataclass(frozen=True)
class JointFitResult:
    parameters_open: AnalysisParameters
    parameters_short: AnalysisParameters
    open_frequency_hz: np.ndarray
    open_measured_ohm: np.ndarray
    open_modeled_ohm: np.ndarray
    short_frequency_hz: np.ndarray
    short_measured_ohm: np.ndarray
    short_modeled_ohm: np.ndarray
    open_fit_min_frequency_hz: float
    short_fit_min_frequency_hz: float
    fit_max_frequency_hz: float
    optimizer_cost: float
    optimizer_optimality: float
    optimizer_evaluations: int


def _metric(measured: np.ndarray, modeled: np.ndarray) -> dict[str, float]:
    magnitude_error_db = 20.0 * np.log10(
        np.maximum(np.abs(modeled), 1e-300)
        / np.maximum(np.abs(measured), 1e-300)
    )
    phase_error_deg = np.degrees(np.angle(modeled / measured))
    return {
        "magnitude_rmse_db": float(np.sqrt(np.mean(magnitude_error_db**2))),
        "phase_rmse_deg": float(np.sqrt(np.mean(phase_error_deg**2))),
        "maximum_absolute_magnitude_error_db": float(
            np.max(np.abs(magnitude_error_db))
        ),
        "maximum_absolute_phase_error_deg": float(
            np.max(np.abs(phase_error_deg))
        ),
    }


def _line_derived(parameters: AnalysisParameters) -> dict[str, float]:
    z0 = parameters.differential_z0_ohm
    vf = parameters.differential_velocity_factor
    length = parameters.length_m
    c_per_m = 1.0 / (z0 * C0 * vf)
    l_per_m = z0 / (C0 * vf)
    delay = length / (C0 * vf)
    return {
        "line_capacitance_pF_per_m": c_per_m * 1e12,
        "line_inductance_nH_per_m": l_per_m * 1e9,
        "one_way_delay_ns": delay * 1e9,
        "quarter_wave_frequency_MHz": 1.0 / (4.0 * delay) / 1e6,
        "half_wave_frequency_MHz": 1.0 / (2.0 * delay) / 1e6,
        "dc_cable_loop_resistance_ohm": (
            2.0
            * parameters.one_core_resistance_ohm_per_m
            * parameters.length_m
        ),
    }


def fit_open_short(
    open_csv: Path,
    short_csv: Path,
    length_m: float,
    open_fit_min_frequency_hz: float = 5.0e4,
    short_fit_min_frequency_hz: float = 1.0e3,
    fit_max_frequency_hz: float = 25.0e6,
    fit_open_capacitance: bool = False,
) -> JointFitResult:
    open_frequency_all, open_impedance_all = read_measured_impedance_csv(open_csv)
    short_frequency_all, short_impedance_all = read_measured_impedance_csv(short_csv)

    open_mask = (
        (open_frequency_all >= open_fit_min_frequency_hz)
        & (open_frequency_all <= fit_max_frequency_hz)
        & (np.abs(open_impedance_all) > 0.0)
    )
    short_mask = (
        (short_frequency_all >= short_fit_min_frequency_hz)
        & (short_frequency_all <= fit_max_frequency_hz)
        & (np.abs(short_impedance_all) > 0.0)
    )

    fo = open_frequency_all[open_mask]
    zo = open_impedance_all[open_mask]
    fs = short_frequency_all[short_mask]
    zs = short_impedance_all[short_mask]

    if fo.size < 30 or fs.size < 30:
        raise ValueError("フィッティング範囲内の測定点が不足しています。")

    base = AnalysisParameters(
        length_m=length_m,
        start_frequency_hz=float(min(open_frequency_all.min(), short_frequency_all.min())),
        stop_frequency_hz=float(max(open_frequency_all.max(), short_frequency_all.max())),
        number_of_points=int(max(open_frequency_all.size, short_frequency_all.size)),
        differential_dc_attenuation_override_np=-1.0,
        differential_open_end_capacitance_pF=0.0,
    )

    # x:
    # [0] Z0 [ohm]
    # [1] VF
    # [2] sqrt(f) loss [Np/sqrt(MHz)]
    # [3] tan delta
    # [4] common fixture R [ohm]
    # [5] common fixture L [nH]
    # [6] short-end R [ohm]
    # [7] short-end L [nH]
    # [8] open-end C [pF], optional
    x0 = [75.0, 0.548, 0.001, 0.08, 0.0, 34.0, 0.11, 15.0]
    lower = [40.0, 0.45, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    upper = [110.0, 0.65, 0.10, 0.15, 2.0, 500.0, 1.0, 500.0]

    if fit_open_capacitance:
        x0.append(2.0)
        lower.append(0.0)
        # Small physical end capacitance only. A large free value causes
        # a strong VF/Cend correlation and loses uniqueness.
        upper.append(15.0)

    def build_parameters(x: np.ndarray, termination: str) -> AnalysisParameters:
        c_open = float(x[8]) if fit_open_capacitance else 0.0
        return replace(
            base,
            differential_termination=termination,
            differential_z0_ohm=float(x[0]),
            differential_velocity_factor=float(x[1]),
            differential_extra_loss_np_per_sqrt_mhz=float(x[2]),
            differential_tan_delta=float(x[3]),
            fixture_series_resistance_ohm=float(x[4]),
            fixture_series_inductance_nH=float(x[5]),
            differential_short_resistance_ohm=float(x[6]),
            differential_short_inductance_nH=float(x[7]),
            differential_open_end_capacitance_pF=c_open,
        )

    def residual(x: np.ndarray) -> np.ndarray:
        po = build_parameters(x, "open")
        ps = build_parameters(x, "short")
        mo = differential_line_model(po, fo)
        ms = differential_line_model(ps, fs)

        # 各測定条件が点数の差で過度に重み付けされないよう、
        # それぞれ sqrt(N) で正規化する。
        open_mag = np.log(
            np.maximum(np.abs(mo), 1e-300)
            / np.maximum(np.abs(zo), 1e-300)
        ) / math.sqrt(fo.size)
        open_phase = 0.8 * np.angle(mo / zo) / math.sqrt(fo.size)
        short_mag = np.log(
            np.maximum(np.abs(ms), 1e-300)
            / np.maximum(np.abs(zs), 1e-300)
        ) / math.sqrt(fs.size)
        short_phase = 0.8 * np.angle(ms / zs) / math.sqrt(fs.size)

        return np.concatenate([open_mag, open_phase, short_mag, short_phase])

    optimized = least_squares(
        residual,
        np.asarray(x0, dtype=float),
        bounds=(np.asarray(lower), np.asarray(upper)),
        loss="soft_l1",
        f_scale=0.01,
        x_scale="jac",
        max_nfev=10000,
    )

    parameters_open = build_parameters(optimized.x, "open")
    parameters_short = build_parameters(optimized.x, "short")

    return JointFitResult(
        parameters_open=parameters_open,
        parameters_short=parameters_short,
        open_frequency_hz=fo,
        open_measured_ohm=zo,
        open_modeled_ohm=differential_line_model(parameters_open, fo),
        short_frequency_hz=fs,
        short_measured_ohm=zs,
        short_modeled_ohm=differential_line_model(parameters_short, fs),
        open_fit_min_frequency_hz=float(fo.min()),
        short_fit_min_frequency_hz=float(fs.min()),
        fit_max_frequency_hz=float(min(fo.max(), fs.max())),
        optimizer_cost=float(optimized.cost),
        optimizer_optimality=float(optimized.optimality),
        optimizer_evaluations=int(optimized.nfev),
    )


def _write_comparison(
    path: Path,
    frequency: np.ndarray,
    measured: np.ndarray,
    modeled: np.ndarray,
) -> None:
    magnitude_error_db = 20.0 * np.log10(
        np.maximum(np.abs(modeled), 1e-300)
        / np.maximum(np.abs(measured), 1e-300)
    )
    phase_error_deg = np.degrees(np.angle(modeled / measured))

    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "frequency_hz",
                "measured_real_ohm",
                "measured_imag_ohm",
                "measured_magnitude_ohm",
                "measured_phase_deg",
                "model_real_ohm",
                "model_imag_ohm",
                "model_magnitude_ohm",
                "model_phase_deg",
                "magnitude_error_db",
                "phase_error_deg",
            ]
        )
        for f, zm, zc, edb, eph in zip(
            frequency,
            measured,
            modeled,
            magnitude_error_db,
            phase_error_deg,
            strict=True,
        ):
            writer.writerow(
                [
                    f"{f:.15g}",
                    f"{zm.real:.15g}",
                    f"{zm.imag:.15g}",
                    f"{abs(zm):.15g}",
                    f"{np.degrees(np.angle(zm)):.15g}",
                    f"{zc.real:.15g}",
                    f"{zc.imag:.15g}",
                    f"{abs(zc):.15g}",
                    f"{np.degrees(np.angle(zc)):.15g}",
                    f"{edb:.15g}",
                    f"{eph:.15g}",
                ]
            )


def _save_plot(
    frequency: np.ndarray,
    measured: np.ndarray,
    modeled: np.ndarray,
    title: str,
    ylabel: str,
    path: Path,
    phase: bool,
    show: bool,
) -> None:
    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    if phase:
        ax.semilogx(
            frequency,
            np.degrees(np.angle(measured)),
            label="実測",
        )
        ax.semilogx(
            frequency,
            np.degrees(np.angle(modeled)),
            label="同時フィット",
        )
        ax.set_ylim(-180.0, 180.0)
        ax.set_yticks(np.arange(-180.0, 181.0, 45.0))
        ax.set_yticks(np.arange(-180.0, 181.0, 15.0), minor=True)
    else:
        ax.loglog(frequency, np.abs(measured), label="実測")
        ax.loglog(frequency, np.abs(modeled), label="同時フィット")

    ax.set_title(title)
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel(ylabel)
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    ax.legend()
    fig.tight_layout()
    fig.savefig(path, dpi=180)
    if show:
        plt.show()
    plt.close(fig)


def export_result(result: JointFitResult, output_dir: Path, show: bool) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    (output_dir / "fitted_5m_open_config.json").write_text(
        json.dumps(
            asdict(result.parameters_open),
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    (output_dir / "fitted_5m_short_config.json").write_text(
        json.dumps(
            asdict(result.parameters_short),
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    open_metrics = _metric(
        result.open_measured_ohm,
        result.open_modeled_ohm,
    )
    short_metrics = _metric(
        result.short_measured_ohm,
        result.short_modeled_ohm,
    )

    summary = {
        "assumption": (
            "開放・短絡は同じ5 mケーブル、同じ治具、同じ校正条件。"
            "開放端は理想開放、短絡端R/Lをフィット。"
        ),
        "fit_ranges_hz": {
            "open": [
                result.open_fit_min_frequency_hz,
                result.fit_max_frequency_hz,
            ],
            "short": [
                result.short_fit_min_frequency_hz,
                result.fit_max_frequency_hz,
            ],
        },
        "fit_points": {
            "open": int(result.open_frequency_hz.size),
            "short": int(result.short_frequency_hz.size),
        },
        "shared_fitted_parameters": {
            "differential_z0_ohm": result.parameters_open.differential_z0_ohm,
            "differential_velocity_factor": (
                result.parameters_open.differential_velocity_factor
            ),
            "differential_extra_loss_np_per_sqrt_mhz": (
                result.parameters_open.differential_extra_loss_np_per_sqrt_mhz
            ),
            "differential_tan_delta_effective": (
                result.parameters_open.differential_tan_delta
            ),
            "fixture_series_resistance_ohm": (
                result.parameters_open.fixture_series_resistance_ohm
            ),
            "fixture_series_inductance_nH": (
                result.parameters_open.fixture_series_inductance_nH
            ),
        },
        "short_end_fitted_parameters": {
            "differential_short_resistance_ohm": (
                result.parameters_short.differential_short_resistance_ohm
            ),
            "differential_short_inductance_nH": (
                result.parameters_short.differential_short_inductance_nH
            ),
        },
        "open_end_capacitance_pF": (
            result.parameters_open.differential_open_end_capacitance_pF
        ),
        "derived_line_parameters": _line_derived(result.parameters_open),
        "open_fit_metrics": open_metrics,
        "short_fit_metrics": short_metrics,
        "optimizer": {
            "cost": result.optimizer_cost,
            "optimality": result.optimizer_optimality,
            "function_evaluations": result.optimizer_evaluations,
        },
        "note": (
            "tanδとsqrt(f)損失はケーブル単体だけでなく、"
            "測定治具・編組・補正残差を含む有効損失値です。"
        ),
    }

    (output_dir / "fit_open_short_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    _write_comparison(
        output_dir / "fit_open_comparison.csv",
        result.open_frequency_hz,
        result.open_measured_ohm,
        result.open_modeled_ohm,
    )
    _write_comparison(
        output_dir / "fit_short_comparison.csv",
        result.short_frequency_hz,
        result.short_measured_ohm,
        result.short_modeled_ohm,
    )

    _save_plot(
        result.open_frequency_hz,
        result.open_measured_ohm,
        result.open_modeled_ohm,
        "5 m・遠端開放：実測と同時フィット",
        "|Z| [Ω]",
        output_dir / "fit_open_magnitude.png",
        phase=False,
        show=show,
    )
    _save_plot(
        result.open_frequency_hz,
        result.open_measured_ohm,
        result.open_modeled_ohm,
        "5 m・遠端開放：位相",
        "位相 [deg]",
        output_dir / "fit_open_phase.png",
        phase=True,
        show=show,
    )
    _save_plot(
        result.short_frequency_hz,
        result.short_measured_ohm,
        result.short_modeled_ohm,
        "5 m・遠端短絡：実測と同時フィット",
        "|Z| [Ω]",
        output_dir / "fit_short_magnitude.png",
        phase=False,
        show=show,
    )
    _save_plot(
        result.short_frequency_hz,
        result.short_measured_ohm,
        result.short_modeled_ohm,
        "5 m・遠端短絡：位相",
        "位相 [deg]",
        output_dir / "fit_short_phase.png",
        phase=True,
        show=show,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="同一ケーブルの遠端開放・短絡CSVを同時フィット"
    )
    parser.add_argument("open_csv", type=Path)
    parser.add_argument("short_csv", type=Path)
    parser.add_argument("--length-m", type=float, default=5.0)
    parser.add_argument(
        "--open-fit-min-frequency-hz",
        type=float,
        default=5.0e4,
    )
    parser.add_argument(
        "--short-fit-min-frequency-hz",
        type=float,
        default=1.0e3,
    )
    parser.add_argument(
        "--fit-max-frequency-hz",
        type=float,
        default=25.0e6,
    )
    parser.add_argument(
        "--fit-open-capacitance",
        action="store_true",
        help="小さい開放端容量0～15 pFもフィットする。",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("fit_5m_open_short"),
    )
    parser.add_argument("--show", action="store_true")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    args = build_parser().parse_args(
        list(argv) if argv is not None else None
    )
    try:
        result = fit_open_short(
            open_csv=args.open_csv,
            short_csv=args.short_csv,
            length_m=args.length_m,
            open_fit_min_frequency_hz=args.open_fit_min_frequency_hz,
            short_fit_min_frequency_hz=args.short_fit_min_frequency_hz,
            fit_max_frequency_hz=args.fit_max_frequency_hz,
            fit_open_capacitance=args.fit_open_capacitance,
        )
        export_result(result, args.output_dir, args.show)

        p = result.parameters_open
        derived = _line_derived(p)
        open_metrics = _metric(
            result.open_measured_ohm,
            result.open_modeled_ohm,
        )
        short_metrics = _metric(
            result.short_measured_ohm,
            result.short_modeled_ohm,
        )

        print("開放・短絡の同時フィッティングが完了しました。")
        print(f"出力先: {args.output_dir.resolve()}")
        print(f"Z0 = {p.differential_z0_ohm:.9g} Ω")
        print(f"VF = {p.differential_velocity_factor:.9g}")
        print(
            "C' = "
            f"{derived['line_capacitance_pF_per_m']:.9g} pF/m"
        )
        print(
            "L' = "
            f"{derived['line_inductance_nH_per_m']:.9g} nH/m"
        )
        print(
            "短絡端 R = "
            f"{result.parameters_short.differential_short_resistance_ohm:.9g} Ω"
        )
        print(
            "短絡端 L = "
            f"{result.parameters_short.differential_short_inductance_nH:.9g} nH"
        )
        print(
            "開放: "
            f"{open_metrics['magnitude_rmse_db']:.4g} dB, "
            f"{open_metrics['phase_rmse_deg']:.4g} deg"
        )
        print(
            "短絡: "
            f"{short_metrics['magnitude_rmse_db']:.4g} dB, "
            f"{short_metrics['phase_rmse_deg']:.4g} deg"
        )
        return 0
    except (ValueError, FileNotFoundError, OSError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
