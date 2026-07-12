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
8. 差動遠端条件 short / open / load
9. WaveForms CSVからの遠端開放自動フィッティング

必要パッケージ
--------------
numpy
matplotlib
scipy

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

実測CSV:
    一般CSVとDigilent WaveForms CSVの双方に対応。
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
from dataclasses import asdict, dataclass, fields, replace
from pathlib import Path
from typing import Iterable, Mapping

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.ticker import FuncFormatter, LogLocator, NullFormatter
from matplotlib import font_manager
from scipy.optimize import least_squares
from scipy.signal import find_peaks


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

    # 差動線路の遠端条件。short / open / load を指定する。
    differential_termination: str = "short"
    # load 選択時の直列RLC。C=0 のときは R+L のみ。
    differential_load_resistance_ohm: float = 50.0
    differential_load_inductance_nH: float = 0.0
    differential_load_capacitance_pF: float = 0.0
    # open 選択時の任意の遠端浮遊容量。0なら理想開放。
    differential_open_end_capacitance_pF: float = 0.0
    # 0以上なら直流減衰 A0 [Np] を直接指定。負値なら導体抵抗から計算。
    differential_dc_attenuation_override_np: float = -1.0
    # 遅延の対数周波数依存。0ならExcelと同じ非分散モデル。
    differential_delay_log_slope_per_decade: float = 0.0

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
        if self.differential_termination not in {"short", "open", "load"}:
            raise ValueError("differential_termination は short / open / load のいずれかです。")
        nonnegative = [
            "differential_load_resistance_ohm",
            "differential_load_inductance_nH",
            "differential_load_capacitance_pF",
            "differential_open_end_capacitance_pF",
            "fixture_series_resistance_ohm",
            "fixture_series_inductance_nH",
        ]
        for name in nonnegative:
            if getattr(self, name) < 0:
                raise ValueError(f"{name} は0以上にしてください。")
        if not -0.25 <= self.differential_delay_log_slope_per_decade <= 0.25:
            raise ValueError("differential_delay_log_slope_per_decade は -0.25～0.25 としてください。")

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


def differential_line_model(
    p: AnalysisParameters,
    frequency_hz: np.ndarray,
) -> np.ndarray:
    """差動線路を short / open / load の遠端条件で計算する。"""
    frequency_hz = np.asarray(frequency_hz, dtype=float)
    omega = 2.0 * math.pi * frequency_hz

    base_theta = omega * p.length_m / (C0 * p.differential_velocity_factor)
    log_frequency = np.log10(np.maximum(frequency_hz, 1e-300) / 1.0e6)
    theta = base_theta * (
        1.0 + p.differential_delay_log_slope_per_decade * log_frequency
    )
    if np.any(theta <= 0):
        raise ValueError("遅延分散係数により電気長が0以下になりました。")

    if p.differential_dc_attenuation_override_np >= 0.0:
        a0 = p.differential_dc_attenuation_override_np
    else:
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

    if p.differential_termination == "open":
        c_open = p.differential_open_end_capacitance_pF * 1e-12
        if c_open > 0.0:
            z_load = 1.0 / (1j * omega * c_open)
            z_in = transmission_line_input_impedance(
                p.differential_z0_ohm, gamma_l, z_load
            )
        else:
            # 理想開放の極限 ZL→∞
            z_in = p.differential_z0_ohm / np.tanh(gamma_l)
    elif p.differential_termination == "short":
        z_load = (
            p.differential_short_resistance_ohm
            + 1j * omega * p.differential_short_inductance_nH * 1e-9
        )
        z_in = transmission_line_input_impedance(
            p.differential_z0_ohm, gamma_l, z_load
        )
    else:
        z_load = (
            p.differential_load_resistance_ohm
            + 1j * omega * p.differential_load_inductance_nH * 1e-9
        )
        c_load = p.differential_load_capacitance_pF * 1e-12
        if c_load > 0.0:
            z_load = z_load + 1.0 / (1j * omega * c_load)
        z_in = transmission_line_input_impedance(
            p.differential_z0_ohm, gamma_l, z_load
        )

    z_fixture = (
        p.fixture_series_resistance_ohm
        + 1j * omega * p.fixture_series_inductance_nH * 1e-9
    )
    return z_in + z_fixture


def calculate_differential(
    p: AnalysisParameters,
    d: DerivedParameters,
    frequency_hz: np.ndarray,
) -> SweepResult:
    # d は他の派生値とのAPI互換性のため受け取る。
    del d
    return SweepResult(frequency_hz, differential_line_model(p, frequency_hz))


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


def read_measured_impedance_csv(path: Path) -> tuple[np.ndarray, np.ndarray]:
    """
    一般CSVまたはDigilent WaveForms Impedance Analyzer CSVを読む。

    対応例:
      frequency_hz,magnitude_ohm,phase_deg
      Frequency (Hz),Trace th (deg),Trace |Z| (Ohm),Trace Rs (Ohm),Trace Xs (Ohm)
    """
    lines = path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    header_index = None
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped and not stripped.startswith("#") and "," in stripped:
            header_index = index
            break
    if header_index is None:
        raise ValueError(f"{path} にCSVヘッダーが見つかりません。")

    reader = csv.DictReader(lines[header_index:])
    if reader.fieldnames is None:
        raise ValueError(f"{path} のCSVヘッダーを解釈できません。")

    normalized = {
        name.strip().lower().replace(" ", "").replace("_", ""): name
        for name in reader.fieldnames
    }

    def find_name(*candidates: str) -> str | None:
        for candidate in candidates:
            key = candidate.lower().replace(" ", "").replace("_", "")
            if key in normalized:
                return normalized[key]
        return None

    frequency_name = find_name("frequency_hz", "Frequency(Hz)")
    magnitude_name = find_name("magnitude_ohm", "Trace|Z|(Ohm)")
    phase_name = find_name("phase_deg", "Traceth(deg)")
    real_name = find_name("impedance_real_ohm", "TraceRs(Ohm)")
    imag_name = find_name("impedance_imag_ohm", "TraceXs(Ohm)")

    if frequency_name is None:
        raise ValueError(f"{path} に周波数列がありません。")
    if (real_name is None or imag_name is None) and (
        magnitude_name is None or phase_name is None
    ):
        raise ValueError(
            f"{path} に複素インピーダンス、または振幅・位相の列がありません。"
        )

    frequencies: list[float] = []
    impedances: list[complex] = []
    for row_number, row in enumerate(reader, start=header_index + 2):
        try:
            f = float(row[frequency_name])
            if real_name is not None and imag_name is not None:
                z = complex(float(row[real_name]), float(row[imag_name]))
            else:
                magnitude = float(row[magnitude_name])
                phase_rad = math.radians(float(row[phase_name]))
                z = magnitude * complex(math.cos(phase_rad), math.sin(phase_rad))
        except (TypeError, ValueError, KeyError) as exc:
            raise ValueError(f"{path} の {row_number} 行目を数値化できません。") from exc
        if f <= 0 or not np.isfinite(f) or not np.isfinite(z.real) or not np.isfinite(z.imag):
            continue
        frequencies.append(f)
        impedances.append(z)

    if not frequencies:
        raise ValueError(f"{path} に有効な測定点がありません。")
    order = np.argsort(frequencies)
    return np.asarray(frequencies)[order], np.asarray(impedances, dtype=complex)[order]


def read_measured_csv(path: Path) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    frequency, impedance = read_measured_impedance_csv(path)
    return frequency, np.abs(impedance), np.degrees(np.angle(impedance))


@dataclass(frozen=True)
class DifferentialFitResult:
    parameters: AnalysisParameters
    measured_frequency_hz: np.ndarray
    measured_impedance_ohm: np.ndarray
    modeled_impedance_ohm: np.ndarray
    fit_min_frequency_hz: float
    fit_max_frequency_hz: float
    magnitude_rmse_db: float
    phase_rmse_deg: float
    optimizer_cost: float
    optimizer_optimality: float
    optimizer_evaluations: int


def _estimate_open_fit_initial_values(
    frequency_hz: np.ndarray,
    impedance_ohm: np.ndarray,
    length_m: float,
) -> tuple[float, float, float]:
    magnitude = np.abs(impedance_ohm)
    log_magnitude = np.log(np.maximum(magnitude, 1e-300))
    minima, _ = find_peaks(-log_magnitude, prominence=0.35)
    maxima, _ = find_peaks(log_magnitude, prominence=0.35)

    usable_minima = [i for i in minima if frequency_hz[i] >= 5e5]
    if usable_minima:
        minimum_index = usable_minima[0]
    else:
        minimum_index = int(np.argmin(magnitude))
    first_min_frequency = frequency_hz[minimum_index]
    vf_initial = float(np.clip(4.0 * length_m * first_min_frequency / C0, 0.40, 0.80))

    later_maxima = [i for i in maxima if i > minimum_index]
    if later_maxima:
        maximum_index = later_maxima[0]
        z0_initial = math.sqrt(magnitude[minimum_index] * magnitude[maximum_index])
    else:
        z0_initial = 65.0
    z0_initial = float(np.clip(z0_initial, 25.0, 140.0))

    ratio = float(np.clip(magnitude[minimum_index] / z0_initial, 1e-5, 0.95))
    a0_initial = float(np.arctanh(ratio))
    return z0_initial, vf_initial, a0_initial


def fit_differential_open_measurement(
    base_parameters: AnalysisParameters,
    measured_csv: Path,
    length_m: float,
    fit_min_frequency_hz: float = 5.0e4,
    fit_max_frequency_hz: float = 25.0e6,
    fit_dispersion: bool = False,
) -> DifferentialFitResult:
    """遠端開放の実測値から Z0、VF、損失、治具直列寄生を自動同定する。"""
    frequency_all, impedance_all = read_measured_impedance_csv(measured_csv)
    mask = (
        (frequency_all >= fit_min_frequency_hz)
        & (frequency_all <= fit_max_frequency_hz)
        & (np.abs(impedance_all) > 0.0)
    )
    frequency = frequency_all[mask]
    measured = impedance_all[mask]
    if frequency.size < 30:
        raise ValueError("フィッティング周波数範囲内の測定点が不足しています。")

    z0_initial, vf_initial, a0_initial = _estimate_open_fit_initial_values(
        frequency, measured, length_m
    )

    # x = Z0, VF, A0, k√f, tanδ, Rfixture, Lfixture, [delay slope]
    x0 = [z0_initial, vf_initial, a0_initial, 0.001, 0.05, 0.5, 20.0]
    lower = [20.0, 0.40, 0.0, 0.0, 0.0, 0.0, 0.0]
    upper = [150.0, 0.80, 1.0, 0.5, 0.20, 30.0, 300.0]
    if fit_dispersion:
        x0.append(-0.01)
        lower.append(-0.15)
        upper.append(0.15)

    def parameters_from_vector(x: np.ndarray) -> AnalysisParameters:
        slope = float(x[7]) if fit_dispersion else 0.0
        return replace(
            base_parameters,
            length_m=float(length_m),
            start_frequency_hz=float(frequency_all.min()),
            stop_frequency_hz=float(frequency_all.max()),
            number_of_points=int(frequency_all.size),
            differential_termination="open",
            differential_z0_ohm=float(x[0]),
            differential_velocity_factor=float(x[1]),
            differential_dc_attenuation_override_np=float(x[2]),
            differential_extra_loss_np_per_sqrt_mhz=float(x[3]),
            differential_tan_delta=float(x[4]),
            fixture_series_resistance_ohm=float(x[5]),
            fixture_series_inductance_nH=float(x[6]),
            differential_delay_log_slope_per_decade=slope,
        )

    def residual(x: np.ndarray) -> np.ndarray:
        modeled = differential_line_model(parameters_from_vector(x), frequency)
        magnitude_residual = np.log(
            np.maximum(np.abs(modeled), 1e-300)
            / np.maximum(np.abs(measured), 1e-300)
        )
        phase_residual = np.angle(modeled / measured)
        return np.concatenate([magnitude_residual, 0.8 * phase_residual])

    optimized = least_squares(
        residual,
        np.asarray(x0, dtype=float),
        bounds=(np.asarray(lower), np.asarray(upper)),
        loss="soft_l1",
        f_scale=0.15,
        max_nfev=10000,
        x_scale="jac",
    )
    fitted_parameters = parameters_from_vector(optimized.x)
    modeled = differential_line_model(fitted_parameters, frequency)
    magnitude_error_db = 20.0 * np.log10(
        np.maximum(np.abs(modeled), 1e-300)
        / np.maximum(np.abs(measured), 1e-300)
    )
    phase_error_deg = np.degrees(np.angle(modeled / measured))

    return DifferentialFitResult(
        parameters=fitted_parameters,
        measured_frequency_hz=frequency,
        measured_impedance_ohm=measured,
        modeled_impedance_ohm=modeled,
        fit_min_frequency_hz=float(frequency.min()),
        fit_max_frequency_hz=float(frequency.max()),
        magnitude_rmse_db=float(np.sqrt(np.mean(magnitude_error_db**2))),
        phase_rmse_deg=float(np.sqrt(np.mean(phase_error_deg**2))),
        optimizer_cost=float(optimized.cost),
        optimizer_optimality=float(optimized.optimality),
        optimizer_evaluations=int(optimized.nfev),
    )


def export_differential_fit(
    result: DifferentialFitResult,
    output_dir: Path,
    show: bool,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    config_path = output_dir / "fitted_5m_open_config.json"
    config_path.write_text(
        json.dumps(asdict(result.parameters), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    measured = result.measured_impedance_ohm
    modeled = result.modeled_impedance_ohm
    with (output_dir / "fit_comparison.csv").open(
        "w", newline="", encoding="utf-8-sig"
    ) as file:
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
        for f, zm, zc in zip(
            result.measured_frequency_hz, measured, modeled, strict=True
        ):
            writer.writerow(
                [
                    f"{f:.15g}",
                    f"{zm.real:.15g}",
                    f"{zm.imag:.15g}",
                    f"{abs(zm):.15g}",
                    f"{math.degrees(math.atan2(zm.imag, zm.real)):.15g}",
                    f"{zc.real:.15g}",
                    f"{zc.imag:.15g}",
                    f"{abs(zc):.15g}",
                    f"{math.degrees(math.atan2(zc.imag, zc.real)):.15g}",
                    f"{20.0 * math.log10(max(abs(zc), 1e-300) / max(abs(zm), 1e-300)):.15g}",
                    f"{math.degrees(np.angle(zc / zm)):.15g}",
                ]
            )

    summary = {
        "fit_range_hz": [
            result.fit_min_frequency_hz,
            result.fit_max_frequency_hz,
        ],
        "number_of_fit_points": int(result.measured_frequency_hz.size),
        "magnitude_rmse_db": result.magnitude_rmse_db,
        "phase_rmse_deg": result.phase_rmse_deg,
        "optimizer_cost": result.optimizer_cost,
        "optimizer_optimality": result.optimizer_optimality,
        "optimizer_evaluations": result.optimizer_evaluations,
        "fitted_parameters": {
            "length_m": result.parameters.length_m,
            "differential_z0_ohm": result.parameters.differential_z0_ohm,
            "differential_velocity_factor": result.parameters.differential_velocity_factor,
            "differential_dc_attenuation_override_np": result.parameters.differential_dc_attenuation_override_np,
            "differential_extra_loss_np_per_sqrt_mhz": result.parameters.differential_extra_loss_np_per_sqrt_mhz,
            "differential_tan_delta": result.parameters.differential_tan_delta,
            "fixture_series_resistance_ohm": result.parameters.fixture_series_resistance_ohm,
            "fixture_series_inductance_nH": result.parameters.fixture_series_inductance_nH,
            "differential_delay_log_slope_per_decade": result.parameters.differential_delay_log_slope_per_decade,
        },
    }
    (output_dir / "fit_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.loglog(result.measured_frequency_hz, np.abs(measured), label="実測")
    ax.loglog(result.measured_frequency_hz, np.abs(modeled), label="フィット")
    ax.set_title("5 m 2芯間・遠端開放 インピーダンス振幅")
    ax.set_xlabel("周波数 [Hz]")
    ax.set_ylabel("|Z| [Ω]")
    configure_log_frequency_axis(ax)
    ax.yaxis.set_minor_locator(LogLocator(base=10.0, subs=np.arange(2.0, 10.0) * 0.1))
    ax.grid(which="major", linewidth=0.8)
    ax.grid(which="minor", linewidth=0.4, linestyle=":")
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_dir / "fit_magnitude.png", dpi=180)
    if show:
        plt.show()
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10.0, 5.8))
    ax.semilogx(
        result.measured_frequency_hz,
        np.degrees(np.angle(measured)),
        label="実測",
    )
    ax.semilogx(
        result.measured_frequency_hz,
        np.degrees(np.angle(modeled)),
        label="フィット",
    )
    ax.set_title("5 m 2芯間・遠端開放 位相")
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
    fig.savefig(output_dir / "fit_phase.png", dpi=180)
    if show:
        plt.show()
    plt.close(fig)


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
        "--differential-termination",
        choices=["short", "open", "load"],
        help="通常解析時の差動遠端条件を上書き。",
    )
    parser.add_argument(
        "--fit-differential-open",
        type=Path,
        metavar="CSV",
        help="遠端開放の実測CSVを読み込み、自動フィッティングして終了。",
    )
    parser.add_argument(
        "--fit-length-m",
        type=float,
        default=5.0,
        help="開放フィットに使う線路長 [m]。既定値5 m。",
    )
    parser.add_argument(
        "--fit-min-frequency-hz",
        type=float,
        default=5.0e4,
        help="開放フィットの下限周波数 [Hz]。",
    )
    parser.add_argument(
        "--fit-max-frequency-hz",
        type=float,
        default=25.0e6,
        help="開放フィットの上限周波数 [Hz]。",
    )
    parser.add_argument(
        "--fit-dispersion",
        action="store_true",
        help="開放フィットで対数周波数依存の遅延分散も同定。",
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
        if args.differential_termination is not None:
            parameters = replace(
                parameters,
                differential_termination=args.differential_termination,
            )

        if args.fit_differential_open is not None:
            if not args.fit_differential_open.exists():
                raise FileNotFoundError(
                    f"実測CSVが見つかりません: {args.fit_differential_open}"
                )
            fitted = fit_differential_open_measurement(
                base_parameters=parameters,
                measured_csv=args.fit_differential_open,
                length_m=args.fit_length_m,
                fit_min_frequency_hz=args.fit_min_frequency_hz,
                fit_max_frequency_hz=args.fit_max_frequency_hz,
                fit_dispersion=args.fit_dispersion,
            )
            export_differential_fit(fitted, args.output_dir, args.show)
            print("遠端開放フィッティングが完了しました。")
            print(f"出力先: {args.output_dir.resolve()}")
            print(f"Z0 = {fitted.parameters.differential_z0_ohm:.6g} Ω")
            print(f"VF = {fitted.parameters.differential_velocity_factor:.6g}")
            print(f"振幅RMSE = {fitted.magnitude_rmse_db:.6g} dB")
            print(f"位相RMSE = {fitted.phase_rmse_deg:.6g} deg")
            return 0

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
