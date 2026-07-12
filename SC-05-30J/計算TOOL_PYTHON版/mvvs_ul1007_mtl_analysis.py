#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MVVS 0.5sq×2芯 + UL1007 AWG20
差動／3導体同相インピーダンス解析

Excelファイル
「MVVS05_UL1007_3導体MTL_差動同相解析_報告書付き」
の計算内容を Python で実装したものです。

主な機能
--------
1. 1 kHz～25 MHz の対数周波数掃引
2. 差動モード入力インピーダンス ZDM
3. 3導体同相モデルの入力インピーダンス ZCM
4. 内部モード／外部モードの分布定数と RLGC 行列
5. CSV、JSON、テキスト、PNG グラフ出力
6. 実測CSVとの比較（任意）
7. JSON設定ファイルによるパラメータ変更

必要パッケージ
--------------
numpy
matplotlib

使用例
------
既定値で実行:
    python mvvs_ul1007_mtl_analysis.py

出力先を指定:
    python mvvs_ul1007_mtl_analysis.py --output-dir results

グラフを画面表示:
    python mvvs_ul1007_mtl_analysis.py --show

既定設定JSONを書き出す:
    python mvvs_ul1007_mtl_analysis.py --write-default-config config.json

設定JSONを読み込む:
    python mvvs_ul1007_mtl_analysis.py --config config.json

実測値と比較:
    python mvvs_ul1007_mtl_analysis.py \
        --measured-dm measured_dm.csv \
        --measured-cm measured_cm.csv

実測CSVの列:
    frequency_hz,magnitude_ohm,phase_deg
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Iterable, Mapping

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import FuncFormatter, LogLocator, NullFormatter
from matplotlib import font_manager


def configure_japanese_font() -> None:
    """利用可能な日本語フォントをMatplotlibへ設定する。"""
    candidates = [
        "Yu Gothic",
        "YuGothic",
        "Meiryo",
        "Noto Sans CJK JP",
        "Noto Sans JP",
        "IPAexGothic",
        "IPAGothic",
        "TakaoGothic",
    ]
    available = {font.name for font in font_manager.fontManager.ttflist}
    selected = next((name for name in candidates if name in available), None)
    if selected is not None:
        plt.rcParams["font.family"] = selected
    plt.rcParams["axes.unicode_minus"] = False


C0 = 299_792_458.0
EPS0 = 8.854_187_817e-12
MU0 = 4.0e-7 * math.pi


@dataclass
class AnalysisParameters:
    """Excel「入力・概要」シートに相当する入力パラメータ。"""

    # 線路・形状
    length_m: float = 6.0
    mvvs_outer_diameter_mm: float = 6.2
    shield_outer_diameter_mm: float = 4.3
    core_insulation_outer_diameter_mm: float = 1.9
    ul1007_outer_diameter_mm: float = 1.8
    ul1007_conductor_diameter_mm: float = 0.95

    # 同相内部モード P-S
    internal_capacitance_pf_per_m: float = 200.0
    internal_velocity_factor: float = 0.620
    internal_tan_delta: float = 0.030
    internal_extra_loss_np_per_sqrt_mhz: float = 0.0030

    # 同相外部モード S-G
    external_effective_relative_permittivity: float = 2.20
    external_tan_delta: float = 0.020
    external_extra_loss_np_per_sqrt_mhz: float = 0.0040

    # 導体抵抗
    one_core_resistance_ohm_per_m: float = 0.0378
    shield_resistance_ohm_per_m: float = 0.0250
    gnd_resistance_ohm_per_m: float = 0.0350

    # 遠端コンデンサ1
    capacitor1_uF: float = 10.0
    capacitor1_esr_ohm: float = 0.050
    capacitor1_esl_nH: float = 20.0

    # 遠端コンデンサ2
    capacitor2_uF: float = 10.0
    capacitor2_esr_ohm: float = 0.050
    capacitor2_esl_nH: float = 20.0

    # 差動モード
    differential_z0_ohm: float = 100.0
    differential_velocity_factor: float = 0.640
    differential_tan_delta: float = 0.025
    differential_extra_loss_np_per_sqrt_mhz: float = 0.0030
    differential_short_resistance_ohm: float = 0.010
    differential_short_inductance_nH: float = 10.0

    # 測定治具の共通直列寄生
    fixture_series_resistance_ohm: float = 0.0
    fixture_series_inductance_nH: float = 0.0

    # 周波数掃引
    start_frequency_hz: float = 1.0e3
    stop_frequency_hz: float = 25.0e6
    number_of_points: int = 501

    def validate(self) -> None:
        positive = [
            "length_m",
            "mvvs_outer_diameter_mm",
            "shield_outer_diameter_mm",
            "core_insulation_outer_diameter_mm",
            "ul1007_outer_diameter_mm",
            "ul1007_conductor_diameter_mm",
            "internal_capacitance_pf_per_m",
            "internal_velocity_factor",
            "external_effective_relative_permittivity",
            "differential_z0_ohm",
            "differential_velocity_factor",
            "capacitor1_uF",
            "capacitor2_uF",
            "start_frequency_hz",
            "stop_frequency_hz",
        ]
        for name in positive:
            if getattr(self, name) <= 0:
                raise ValueError(f"{name} は正の値である必要があります。")

        if self.stop_frequency_hz <= self.start_frequency_hz:
            raise ValueError("stop_frequency_hz は start_frequency_hz より大きくしてください。")
        if self.number_of_points < 2:
            raise ValueError("number_of_points は2以上にしてください。")
        if self.internal_velocity_factor > 1.0 or self.differential_velocity_factor > 1.0:
            raise ValueError("速度係数は1以下にしてください。")

        d = 0.5 * (self.mvvs_outer_diameter_mm + self.ul1007_outer_diameter_mm)
        a = 0.5 * self.shield_outer_diameter_mm
        b = 0.5 * self.ul1007_conductor_diameter_mm
        if d <= a + b:
            raise ValueError(
                "外部モードの2円柱近似で導体が重なっています。"
                "外径・シールド径・GND導体径を確認してください。"
            )


@dataclass(frozen=True)
class DerivedParameters:
    center_distance_mm: float
    shield_to_gnd_conductor_min_distance_mm: float
    external_chi: float

    internal_capacitance_f_per_m: float
    internal_inductance_h_per_m: float
    internal_z0_ohm: float
    internal_velocity_factor: float
    internal_one_way_delay_s: float
    internal_quarter_wave_hz: float

    external_capacitance_f_per_m: float
    external_inductance_h_per_m: float
    external_z0_ohm: float
    external_velocity_factor: float
    external_one_way_delay_s: float
    external_quarter_wave_hz: float

    pair_bundle_resistance_ohm_per_m: float
    internal_loop_resistance_ohm: float
    external_loop_resistance_ohm: float

    capacitor_parallel_low_frequency_uF: float
    capacitor_parallel_equal_esr_ohm: float
    capacitor_parallel_equal_esl_nH: float
    estimated_low_frequency_series_resonance_hz: float

    differential_one_way_delay_s: float
    differential_quarter_wave_hz: float
    differential_half_wave_hz: float


@dataclass(frozen=True)
class SweepResult:
    frequency_hz: np.ndarray
    impedance_ohm: np.ndarray

    @property
    def real_ohm(self) -> np.ndarray:
        return np.real(self.impedance_ohm)

    @property
    def imag_ohm(self) -> np.ndarray:
        return np.imag(self.impedance_ohm)

    @property
    def magnitude_ohm(self) -> np.ndarray:
        return np.abs(self.impedance_ohm)

    @property
    def phase_deg(self) -> np.ndarray:
        return np.degrees(np.angle(self.impedance_ohm))


@dataclass(frozen=True)
class AnalysisResult:
    parameters: AnalysisParameters
    derived: DerivedParameters
    differential: SweepResult
    common_mode: SweepResult
    common_mode_load_ohm: np.ndarray
    common_mode_external_impedance_ohm: np.ndarray
    r_matrix_ohm_per_m: np.ndarray
    l_matrix_h_per_m: np.ndarray
    g_matrix_s_per_m_at_1mhz: np.ndarray
    c_matrix_f_per_m: np.ndarray


def derive_parameters(p: AnalysisParameters) -> DerivedParameters:
    p.validate()

    length = p.length_m

    # シールドケーブルとGND線の外被が密着するときの中心間距離。
    center_distance_mm = 0.5 * (p.mvvs_outer_diameter_mm + p.ul1007_outer_diameter_mm)
    shield_to_gnd_min_mm = (
        center_distance_mm
        - 0.5 * p.shield_outer_diameter_mm
        - 0.5 * p.ul1007_conductor_diameter_mm
    )

    # 外部モード: 異径2円柱線路。
    d = center_distance_mm * 1e-3
    a = 0.5 * p.shield_outer_diameter_mm * 1e-3
    b = 0.5 * p.ul1007_conductor_diameter_mm * 1e-3
    chi_argument = (d * d - a * a - b * b) / (2.0 * a * b)
    chi = math.acosh(chi_argument)

    c1 = p.internal_capacitance_pf_per_m * 1e-12
    vf1 = p.internal_velocity_factor
    z01 = 1.0 / (C0 * vf1 * c1)
    l1 = z01 / (C0 * vf1)
    td1 = length / (C0 * vf1)

    eps_eff = p.external_effective_relative_permittivity
    c2 = 2.0 * math.pi * EPS0 * eps_eff / chi
    l2 = MU0 * chi / (2.0 * math.pi)
    z02 = math.sqrt(l2 / c2)
    vf2 = 1.0 / math.sqrt(eps_eff)
    td2 = length / (C0 * vf2)

    pair_resistance = 0.5 * p.one_core_resistance_ohm_per_m
    internal_loop_r = (
        pair_resistance + p.shield_resistance_ohm_per_m
    ) * length
    external_loop_r = (
        p.shield_resistance_ohm_per_m + p.gnd_resistance_ohm_per_m
    ) * length

    c1_uF = p.capacitor1_uF
    c2_uF = p.capacitor2_uF
    c_parallel_uF = c1_uF + c2_uF

    # Excelと同じく、2個が同一値の場合のESR/ESL並列値を表示。
    equal_esr = (
        0.5 * p.capacitor1_esr_ohm
        if math.isclose(p.capacitor1_esr_ohm, p.capacitor2_esr_ohm)
        else math.nan
    )
    equal_esl = (
        0.5 * p.capacitor1_esl_nH
        if math.isclose(p.capacitor1_esl_nH, p.capacitor2_esl_nH)
        else math.nan
    )

    # Excelの概算式:
    # fr = 1/(2π√(((L1+L2)l + ESL_eq) C_eq))
    esl_for_estimate_h = 0.0 if math.isnan(equal_esl) else equal_esl * 1e-9
    total_l_estimate_h = (l1 + l2) * length + esl_for_estimate_h
    estimated_fr = 1.0 / (
        2.0 * math.pi * math.sqrt(total_l_estimate_h * c_parallel_uF * 1e-6)
    )

    td_d = length / (C0 * p.differential_velocity_factor)

    return DerivedParameters(
        center_distance_mm=center_distance_mm,
        shield_to_gnd_conductor_min_distance_mm=shield_to_gnd_min_mm,
        external_chi=chi,
        internal_capacitance_f_per_m=c1,
        internal_inductance_h_per_m=l1,
        internal_z0_ohm=z01,
        internal_velocity_factor=vf1,
        internal_one_way_delay_s=td1,
        internal_quarter_wave_hz=1.0 / (4.0 * td1),
        external_capacitance_f_per_m=c2,
        external_inductance_h_per_m=l2,
        external_z0_ohm=z02,
        external_velocity_factor=vf2,
        external_one_way_delay_s=td2,
        external_quarter_wave_hz=1.0 / (4.0 * td2),
        pair_bundle_resistance_ohm_per_m=pair_resistance,
        internal_loop_resistance_ohm=internal_loop_r,
        external_loop_resistance_ohm=external_loop_r,
        capacitor_parallel_low_frequency_uF=c_parallel_uF,
        capacitor_parallel_equal_esr_ohm=equal_esr,
        capacitor_parallel_equal_esl_nH=equal_esl,
        estimated_low_frequency_series_resonance_hz=estimated_fr,
        differential_one_way_delay_s=td_d,
        differential_quarter_wave_hz=1.0 / (4.0 * td_d),
        differential_half_wave_hz=1.0 / (2.0 * td_d),
    )


def _safe_atanh_ratio(resistance_ohm: float, z0_ohm: float) -> float:
    """Excelの ATANH(MIN(0.999999, Rloop/Z0)) に対応。"""
    ratio = min(0.999999, max(0.0, resistance_ohm / z0_ohm))
    return math.atanh(ratio)


def _distributed_loss(
    frequency_hz: np.ndarray,
    electrical_angle_rad: np.ndarray,
    dc_attenuation_np: float,
    extra_loss_np_per_sqrt_mhz: float,
    tan_delta: float,
) -> np.ndarray:
    """
    Excelで使用している減衰モデル。

    A = A0 + k√(f/MHz) + θ tanδ / 2
    """
    return (
        dc_attenuation_np
        + extra_loss_np_per_sqrt_mhz * np.sqrt(frequency_hz / 1e6)
        + 0.5 * tan_delta * electrical_angle_rad
    )


def capacitor_impedance(
    omega_rad_s: np.ndarray,
    capacitance_uF: float,
    esr_ohm: float,
    esl_nH: float,
) -> np.ndarray:
    c = capacitance_uF * 1e-6
    l = esl_nH * 1e-9
    return esr_ohm + 1j * (omega_rad_s * l - 1.0 / (omega_rad_s * c))


def parallel(z1: np.ndarray, z2: np.ndarray) -> np.ndarray:
    return z1 * z2 / (z1 + z2)


def transmission_line_input_impedance(
    z0_ohm: float,
    gamma_length: np.ndarray,
    load_impedance_ohm: np.ndarray,
) -> np.ndarray:
    """
    損失線路の入力インピーダンス。

    Zin = Z0 (ZL + Z0 tanh(γl)) / (Z0 + ZL tanh(γl))
    """
    t = np.tanh(gamma_length)
    return z0_ohm * (load_impedance_ohm + z0_ohm * t) / (
        z0_ohm + load_impedance_ohm * t
    )


def calculate_differential(
    p: AnalysisParameters,
    d: DerivedParameters,
    frequency_hz: np.ndarray,
) -> SweepResult:
    omega = 2.0 * math.pi * frequency_hz

    theta = omega * d.differential_one_way_delay_s
    # Excel: ATANH((2*r_core*l)/Z0d)
    dc_loop_r = 2.0 * p.one_core_resistance_ohm_per_m * p.length_m
    a0 = _safe_atanh_ratio(dc_loop_r, p.differential_z0_ohm)
    attenuation = _distributed_loss(
        frequency_hz,
        theta,
        a0,
        p.differential_extra_loss_np_per_sqrt_mhz,
        p.differential_tan_delta,
    )

    gamma_l = attenuation + 1j * theta
    z_load = (
        p.differential_short_resistance_ohm
        + 1j * omega * p.differential_short_inductance_nH * 1e-9
    )

    z_in = transmission_line_input_impedance(
        p.differential_z0_ohm,
        gamma_l,
        z_load,
    )
    z_fixture = (
        p.fixture_series_resistance_ohm
        + 1j * omega * p.fixture_series_inductance_nH * 1e-9
    )
    return SweepResult(frequency_hz, z_in + z_fixture)


def calculate_common_mode(
    p: AnalysisParameters,
    d: DerivedParameters,
    frequency_hz: np.ndarray,
) -> tuple[SweepResult, np.ndarray, np.ndarray]:
    """
    3導体同相モデルのモード解。

    物理導体:
        P = 2芯束
        S = 編組シールド
        G = UL1007 GND線

    モード:
        v1 = VP - VS, i1 = IP                 内部モード P-S
        v2 = VS - VG, i2 = IP + IS = -IG     外部モード S-G

    境界条件:
        近端: VSG(0)=0
        遠端: IS(l)=0
        遠端P-G負荷: 10 µF×2 の並列

    遠端から見た外部モード:
        Zext = Z02 tanh(γ2 l)

    内部モードの実効負荷:
        ZLeff = ZC + Zext
    """
    omega = 2.0 * math.pi * frequency_hz

    # 内部モード
    theta1 = omega * d.internal_one_way_delay_s
    a01 = _safe_atanh_ratio(
        d.internal_loop_resistance_ohm,
        d.internal_z0_ohm,
    )
    attenuation1 = _distributed_loss(
        frequency_hz,
        theta1,
        a01,
        p.internal_extra_loss_np_per_sqrt_mhz,
        p.internal_tan_delta,
    )
    gamma1_l = attenuation1 + 1j * theta1

    # 外部モード
    theta2 = omega * d.external_one_way_delay_s
    a02 = _safe_atanh_ratio(
        d.external_loop_resistance_ohm,
        d.external_z0_ohm,
    )
    attenuation2 = _distributed_loss(
        frequency_hz,
        theta2,
        a02,
        p.external_extra_loss_np_per_sqrt_mhz,
        p.external_tan_delta,
    )
    gamma2_l = attenuation2 + 1j * theta2

    zc1 = capacitor_impedance(
        omega,
        p.capacitor1_uF,
        p.capacitor1_esr_ohm,
        p.capacitor1_esl_nH,
    )
    zc2 = capacitor_impedance(
        omega,
        p.capacitor2_uF,
        p.capacitor2_esr_ohm,
        p.capacitor2_esl_nH,
    )
    z_load = parallel(zc1, zc2)

    # 近端S-G短絡線路を遠端から見る。
    z_external = d.external_z0_ohm * np.tanh(gamma2_l)
    z_effective_load = z_load + z_external

    z_in = transmission_line_input_impedance(
        d.internal_z0_ohm,
        gamma1_l,
        z_effective_load,
    )
    z_fixture = (
        p.fixture_series_resistance_ohm
        + 1j * omega * p.fixture_series_inductance_nH * 1e-9
    )
    return SweepResult(frequency_hz, z_in + z_fixture), z_load, z_external


def calculate_rlgc_matrices(
    p: AnalysisParameters,
    d: DerivedParameters,
    frequency_hz: float = 1.0e6,
) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    """
    3導体系座標 V=[VP-VG, VS-VG]^T, I=[IP, IS]^T の RLGC 行列。

    L = [[L1+L2, L2],
         [L2,    L2]]

    C = [[C1,   -C1],
         [-C1, C1+C2]]

    R = [[r1+r2, r2],
         [r2,    r2]]

    G = [[g1,   -g1],
         [-g1, g1+g2]]
    """
    l1 = d.internal_inductance_h_per_m
    l2 = d.external_inductance_h_per_m
    c1 = d.internal_capacitance_f_per_m
    c2 = d.external_capacitance_f_per_m

    r1 = d.pair_bundle_resistance_ohm_per_m + p.shield_resistance_ohm_per_m
    r2 = p.shield_resistance_ohm_per_m + p.gnd_resistance_ohm_per_m

    omega = 2.0 * math.pi * frequency_hz
    g1 = omega * c1 * p.internal_tan_delta
    g2 = omega * c2 * p.external_tan_delta

    r = np.array([[r1 + r2, r2], [r2, r2]], dtype=float)
    l = np.array([[l1 + l2, l2], [l2, l2]], dtype=float)
    g = np.array([[g1, -g1], [-g1, g1 + g2]], dtype=float)
    c = np.array([[c1, -c1], [-c1, c1 + c2]], dtype=float)
    return r, l, g, c


def run_analysis(parameters: AnalysisParameters) -> AnalysisResult:
    parameters.validate()
    derived = derive_parameters(parameters)
    frequency_hz = np.logspace(
        math.log10(parameters.start_frequency_hz),
        math.log10(parameters.stop_frequency_hz),
        parameters.number_of_points,
    )

    differential = calculate_differential(parameters, derived, frequency_hz)
    common_mode, z_load, z_external = calculate_common_mode(
        parameters,
        derived,
        frequency_hz,
    )
    r, l, g, c = calculate_rlgc_matrices(parameters, derived)

    return AnalysisResult(
        parameters=parameters,
        derived=derived,
        differential=differential,
        common_mode=common_mode,
        common_mode_load_ohm=z_load,
        common_mode_external_impedance_ohm=z_external,
        r_matrix_ohm_per_m=r,
        l_matrix_h_per_m=l,
        g_matrix_s_per_m_at_1mhz=g,
        c_matrix_f_per_m=c,
    )


def engineering_frequency(value: float, _position: float | None = None) -> str:
    if value <= 0:
        return ""
    if value >= 1e6:
        scaled = value / 1e6
        return f"{scaled:g}M"
    if value >= 1e3:
        scaled = value / 1e3
        return f"{scaled:g}k"
    return f"{value:g}"


def configure_log_frequency_axis(ax: plt.Axes) -> None:
    ax.set_xscale("log")
    ax.set_xlim(left=1e3)
    ax.xaxis.set_major_locator(LogLocator(base=10.0))
    ax.xaxis.set_minor_locator(
        LogLocator(base=10.0, subs=np.arange(2.0, 10.0) * 0.1)
    )
    ax.xaxis.set_major_formatter(FuncFormatter(engineering_frequency))
    ax.xaxis.set_minor_formatter(NullFormatter())
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")


def save_impedance_magnitude_plot(
    result: SweepResult,
    title: str,
    output_path: Path,
    show: bool,
) -> None:
    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.loglog(result.frequency_hz, result.magnitude_ohm)
    ax.set_title(title)
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel("|Z| [Ω]")
    configure_log_frequency_axis(ax)
    ax.yaxis.set_major_locator(LogLocator(base=10.0))
    ax.yaxis.set_minor_locator(
        LogLocator(base=10.0, subs=np.arange(2.0, 10.0) * 0.1)
    )
    ax.yaxis.set_minor_formatter(NullFormatter())
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    fig.tight_layout()
    fig.savefig(output_path, dpi=180)
    if show:
        plt.show()
    plt.close(fig)


def save_phase_plot(
    result: SweepResult,
    title: str,
    output_path: Path,
    show: bool,
) -> None:
    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.semilogx(result.frequency_hz, result.phase_deg)
    ax.set_title(title)
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel("位相 [deg]")
    ax.set_ylim(-180.0, 180.0)
    configure_log_frequency_axis(ax)
    ax.set_yticks(np.arange(-180.0, 181.0, 45.0))
    ax.set_yticks(np.arange(-180.0, 181.0, 15.0), minor=True)
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    fig.tight_layout()
    fig.savefig(output_path, dpi=180)
    if show:
        plt.show()
    plt.close(fig)


def write_sweep_csv(
    path: Path,
    result: SweepResult,
    extra_columns: Mapping[str, np.ndarray] | None = None,
) -> None:
    extra_columns = extra_columns or {}
    headers = [
        "frequency_hz",
        "impedance_real_ohm",
        "impedance_imag_ohm",
        "magnitude_ohm",
        "phase_deg",
    ] + list(extra_columns)

    arrays: list[np.ndarray] = [
        result.frequency_hz,
        result.real_ohm,
        result.imag_ohm,
        result.magnitude_ohm,
        result.phase_deg,
    ] + [np.asarray(v) for v in extra_columns.values()]

    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        for row in zip(*arrays, strict=True):
            writer.writerow([f"{float(value):.15g}" for value in row])


def write_derived_csv(path: Path, d: DerivedParameters) -> None:
    units = {
        "center_distance_mm": "mm",
        "shield_to_gnd_conductor_min_distance_mm": "mm",
        "external_chi": "-",
        "internal_capacitance_f_per_m": "F/m",
        "internal_inductance_h_per_m": "H/m",
        "internal_z0_ohm": "ohm",
        "internal_velocity_factor": "-",
        "internal_one_way_delay_s": "s",
        "internal_quarter_wave_hz": "Hz",
        "external_capacitance_f_per_m": "F/m",
        "external_inductance_h_per_m": "H/m",
        "external_z0_ohm": "ohm",
        "external_velocity_factor": "-",
        "external_one_way_delay_s": "s",
        "external_quarter_wave_hz": "Hz",
        "pair_bundle_resistance_ohm_per_m": "ohm/m",
        "internal_loop_resistance_ohm": "ohm",
        "external_loop_resistance_ohm": "ohm",
        "capacitor_parallel_low_frequency_uF": "uF",
        "capacitor_parallel_equal_esr_ohm": "ohm",
        "capacitor_parallel_equal_esl_nH": "nH",
        "estimated_low_frequency_series_resonance_hz": "Hz",
        "differential_one_way_delay_s": "s",
        "differential_quarter_wave_hz": "Hz",
        "differential_half_wave_hz": "Hz",
    }
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["parameter", "value", "unit"])
        for item in fields(d):
            value = getattr(d, item.name)
            writer.writerow([item.name, f"{value:.15g}", units.get(item.name, "")])


def matrix_text(name: str, matrix: np.ndarray, scale: float, unit: str) -> str:
    m = matrix * scale
    return (
        f"{name} [{unit}]\n"
        f"[[{m[0, 0]:.9g}, {m[0, 1]:.9g}],\n"
        f" [{m[1, 0]:.9g}, {m[1, 1]:.9g}]]\n"
    )


def write_summary(path: Path, result: AnalysisResult) -> None:
    d = result.derived
    text = [
        "MVVS 0.5sq×2芯 + UL1007 AWG20",
        "差動／3導体同相インピーダンス解析",
        "",
        "■ 主な派生値",
        f"ケーブル-GND線中心間距離 = {d.center_distance_mm:.6g} mm",
        f"シールド-GND導体最短距離 = {d.shield_to_gnd_conductor_min_distance_mm:.6g} mm",
        f"内部モード Z01 = {d.internal_z0_ohm:.9g} Ω",
        f"内部モード L1 = {d.internal_inductance_h_per_m*1e9:.9g} nH/m",
        f"内部モード C1 = {d.internal_capacitance_f_per_m*1e12:.9g} pF/m",
        f"内部モード 片道遅延 = {d.internal_one_way_delay_s*1e9:.9g} ns",
        f"内部モード 1/4波長 = {d.internal_quarter_wave_hz/1e6:.9g} MHz",
        f"外部モード Z02 = {d.external_z0_ohm:.9g} Ω",
        f"外部モード L2 = {d.external_inductance_h_per_m*1e9:.9g} nH/m",
        f"外部モード C2 = {d.external_capacitance_f_per_m*1e12:.9g} pF/m",
        f"外部モード 片道遅延 = {d.external_one_way_delay_s*1e9:.9g} ns",
        f"外部モード 1/4波長 = {d.external_quarter_wave_hz/1e6:.9g} MHz",
        f"差動モード 1/4波長 = {d.differential_quarter_wave_hz/1e6:.9g} MHz",
        f"差動モード 1/2波長 = {d.differential_half_wave_hz/1e6:.9g} MHz",
        f"低周波概算直列共振 = {d.estimated_low_frequency_series_resonance_hz/1e3:.9g} kHz",
        "",
        "■ 3導体RLGC行列",
        matrix_text("R", result.r_matrix_ohm_per_m, 1.0, "Ω/m"),
        matrix_text("L", result.l_matrix_h_per_m, 1e9, "nH/m"),
        matrix_text("G @ 1 MHz", result.g_matrix_s_per_m_at_1mhz, 1e6, "µS/m"),
        matrix_text("C", result.c_matrix_f_per_m, 1e12, "pF/m"),
        "■ 数式",
        "差動・内部モード入力インピーダンス:",
        "Zin = Z0 (ZL + Z0 tanh(γl)) / (Z0 + ZL tanh(γl))",
        "",
        "同相3導体モード:",
        "Zext = Z02 tanh(γ2 l)",
        "ZLeff = (ZC1 || ZC2) + Zext",
        "ZCM = Z01 (ZLeff + Z01 tanh(γ1 l)) / (Z01 + ZLeff tanh(γ1 l))",
    ]
    path.write_text("\n".join(text), encoding="utf-8")


def read_measured_csv(path: Path) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    frequencies: list[float] = []
    magnitudes: list[float] = []
    phases: list[float] = []

    with path.open("r", newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        required = {"frequency_hz", "magnitude_ohm", "phase_deg"}
        if reader.fieldnames is None or not required.issubset(reader.fieldnames):
            raise ValueError(
                f"{path} の列は frequency_hz,magnitude_ohm,phase_deg としてください。"
            )
        for row_number, row in enumerate(reader, start=2):
            try:
                f = float(row["frequency_hz"])
                m = float(row["magnitude_ohm"])
                ph = float(row["phase_deg"])
            except (TypeError, ValueError) as exc:
                raise ValueError(f"{path} の {row_number} 行目を数値化できません。") from exc
            if f <= 0 or m < 0:
                raise ValueError(f"{path} の {row_number} 行目の値が不正です。")
            frequencies.append(f)
            magnitudes.append(m)
            phases.append(ph)

    order = np.argsort(frequencies)
    return (
        np.asarray(frequencies)[order],
        np.asarray(magnitudes)[order],
        np.asarray(phases)[order],
    )


def interpolate_model(
    model: SweepResult,
    measured_frequency_hz: np.ndarray,
) -> tuple[np.ndarray, np.ndarray]:
    log_model_f = np.log10(model.frequency_hz)
    log_measured_f = np.log10(measured_frequency_hz)

    model_mag = 10.0 ** np.interp(
        log_measured_f,
        log_model_f,
        np.log10(np.maximum(model.magnitude_ohm, 1e-300)),
    )
    unwrapped_phase = np.unwrap(np.radians(model.phase_deg))
    phase_rad = np.interp(log_measured_f, log_model_f, unwrapped_phase)
    phase_deg = (np.degrees(phase_rad) + 180.0) % 360.0 - 180.0
    return model_mag, phase_deg


def write_comparison(
    input_path: Path,
    model: SweepResult,
    output_csv: Path,
    output_magnitude_png: Path,
    output_phase_png: Path,
    title_prefix: str,
    show: bool,
) -> None:
    f, measured_mag, measured_phase = read_measured_csv(input_path)
    model_mag, model_phase = interpolate_model(model, f)
    magnitude_error_percent = 100.0 * (model_mag - measured_mag) / np.maximum(
        measured_mag, 1e-300
    )
    phase_error_deg = model_phase - measured_phase

    with output_csv.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "frequency_hz",
                "measured_magnitude_ohm",
                "measured_phase_deg",
                "model_magnitude_ohm",
                "model_phase_deg",
                "magnitude_error_percent",
                "phase_error_deg",
            ]
        )
        for values in zip(
            f,
            measured_mag,
            measured_phase,
            model_mag,
            model_phase,
            magnitude_error_percent,
            phase_error_deg,
            strict=True,
        ):
            writer.writerow([f"{float(value):.15g}" for value in values])

    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.loglog(f, measured_mag, marker="o", linestyle="none", label="実測")
    ax.loglog(model.frequency_hz, model.magnitude_ohm, label="モデル")
    ax.set_title(f"{title_prefix} 振幅比較")
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel("|Z| [Ω]")
    configure_log_frequency_axis(ax)
    ax.yaxis.set_minor_locator(
        LogLocator(base=10.0, subs=np.arange(2.0, 10.0) * 0.1)
    )
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_magnitude_png, dpi=180)
    if show:
        plt.show()
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.semilogx(f, measured_phase, marker="o", linestyle="none", label="実測")
    ax.semilogx(model.frequency_hz, model.phase_deg, label="モデル")
    ax.set_title(f"{title_prefix} 位相比較")
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel("位相 [deg]")
    ax.set_ylim(-180.0, 180.0)
    configure_log_frequency_axis(ax)
    ax.set_yticks(np.arange(-180.0, 181.0, 45.0))
    ax.set_yticks(np.arange(-180.0, 181.0, 15.0), minor=True)
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_phase_png, dpi=180)
    if show:
        plt.show()
    plt.close(fig)


def load_parameters(path: Path) -> AnalysisParameters:
    data = json.loads(path.read_text(encoding="utf-8"))
    known = {item.name for item in fields(AnalysisParameters)}
    unknown = sorted(set(data) - known)
    if unknown:
        raise ValueError(f"設定JSONに未知の項目があります: {', '.join(unknown)}")
    parameters = AnalysisParameters(**data)
    parameters.validate()
    return parameters


def write_default_config(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(asdict(AnalysisParameters()), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def export_all(
    result: AnalysisResult,
    output_dir: Path,
    show: bool,
    measured_dm: Path | None,
    measured_cm: Path | None,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    (output_dir / "used_parameters.json").write_text(
        json.dumps(asdict(result.parameters), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_derived_csv(output_dir / "derived_parameters.csv", result.derived)
    write_summary(output_dir / "analysis_summary.txt", result)

    write_sweep_csv(
        output_dir / "differential_results.csv",
        result.differential,
    )
    write_sweep_csv(
        output_dir / "common_mode_results.csv",
        result.common_mode,
        {
            "load_magnitude_ohm": np.abs(result.common_mode_load_ohm),
            "load_phase_deg": np.degrees(np.angle(result.common_mode_load_ohm)),
            "external_impedance_magnitude_ohm": np.abs(
                result.common_mode_external_impedance_ohm
            ),
            "external_impedance_phase_deg": np.degrees(
                np.angle(result.common_mode_external_impedance_ohm)
            ),
        },
    )

    save_impedance_magnitude_plot(
        result.differential,
        "差動モード |ZDM|",
        output_dir / "differential_magnitude.png",
        show,
    )
    save_phase_plot(
        result.differential,
        "差動モード 位相",
        output_dir / "differential_phase.png",
        show,
    )
    save_impedance_magnitude_plot(
        result.common_mode,
        "同相3導体 |ZCM|",
        output_dir / "common_mode_magnitude.png",
        show,
    )
    save_phase_plot(
        result.common_mode,
        "同相3導体 位相",
        output_dir / "common_mode_phase.png",
        show,
    )

    if measured_dm is not None:
        write_comparison(
            measured_dm,
            result.differential,
            output_dir / "comparison_differential.csv",
            output_dir / "comparison_differential_magnitude.png",
            output_dir / "comparison_differential_phase.png",
            "差動モード",
            show,
        )
    if measured_cm is not None:
        write_comparison(
            measured_cm,
            result.common_mode,
            output_dir / "comparison_common_mode.csv",
            output_dir / "comparison_common_mode_magnitude.png",
            output_dir / "comparison_common_mode_phase.png",
            "同相3導体",
            show,
        )


def print_console_summary(result: AnalysisResult, output_dir: Path) -> None:
    d = result.derived
    print("解析が完了しました。")
    print(f"出力先: {output_dir.resolve()}")
    print(f"内部モード Z01: {d.internal_z0_ohm:.6g} Ω")
    print(f"外部モード Z02: {d.external_z0_ohm:.6g} Ω")
    print(f"内部モード 1/4波長: {d.internal_quarter_wave_hz/1e6:.6g} MHz")
    print(f"外部モード 1/4波長: {d.external_quarter_wave_hz/1e6:.6g} MHz")
    print(f"差動モード 1/4波長: {d.differential_quarter_wave_hz/1e6:.6g} MHz")
    print(f"差動モード 1/2波長: {d.differential_half_wave_hz/1e6:.6g} MHz")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="MVVS 0.5sq×2芯 + UL1007 AWG20 の差動／3導体同相解析"
    )
    parser.add_argument(
        "--config",
        type=Path,
        help="入力パラメータJSON。省略時はExcelと同じ既定値。",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("mvvs_ul1007_results"),
        help="結果出力フォルダ。",
    )
    parser.add_argument(
        "--show",
        action="store_true",
        help="保存と同時にグラフを画面表示。",
    )
    parser.add_argument(
        "--measured-dm",
        type=Path,
        help="差動実測CSV。",
    )
    parser.add_argument(
        "--measured-cm",
        type=Path,
        help="同相実測CSV。",
    )
    parser.add_argument(
        "--write-default-config",
        type=Path,
        metavar="PATH",
        help="既定値の設定JSONを書き出して終了。",
    )
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    try:
        if args.write_default_config is not None:
            write_default_config(args.write_default_config)
            print(f"既定設定を書き出しました: {args.write_default_config}")
            return 0

        configure_japanese_font()

        parameters = (
            load_parameters(args.config)
            if args.config is not None
            else AnalysisParameters()
        )

        for measured_path in (args.measured_dm, args.measured_cm):
            if measured_path is not None and not measured_path.exists():
                raise FileNotFoundError(f"実測CSVが見つかりません: {measured_path}")

        result = run_analysis(parameters)
        export_all(
            result=result,
            output_dir=args.output_dir,
            show=args.show,
            measured_dm=args.measured_dm,
            measured_cm=args.measured_cm,
        )
        print_console_summary(result, args.output_dir)
        return 0

    except (ValueError, FileNotFoundError, OSError, json.JSONDecodeError) as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
