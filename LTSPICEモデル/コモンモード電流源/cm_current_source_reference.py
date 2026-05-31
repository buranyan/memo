#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Common-mode current-source excitation reference calculation.

Two current sources of 0.5 A are connected from DUMMY HOT/RTN to chassis.
The circuit is assumed symmetric, so a common-mode equivalent ladder is used.
"""
from pathlib import Path
import numpy as np
import matplotlib.pyplot as plt


def z_cap_esr_esl(s, C, ESR, ESL):
    return ESR + s*ESL + 1/(s*C)


def z_ind_parallel_c(s, L, R, Cpa):
    zrl = R + s*L
    return 1/(1/zrl + s*Cpa) if Cpa > 0 else zrl


def calc(f, I_each=0.5):
    w = 2*np.pi*f
    s = 1j*w
    # shunt at node 1: LINE_PS C_CHA + feedthrough caps
    Y_line_ps_ccha = s * 2 * 50e-12
    Z_through_each = z_cap_esr_esl(s, 10e-6, 20e-3, 10e-9)
    Y_through = 2 / Z_through_each
    Y1 = Y_line_ps_ccha + Y_through

    # line input series and shunt
    Z_line_input = (20e-3 + s*1e-6) / 2
    Y2 = s * 2 * 50e-12

    # LDM1/LDM2 CM series
    Z_LDM1 = z_ind_parallel_c(s, 68e-6, 20e-3, 5e-12) / 2
    Z_LDM2 = z_ind_parallel_c(s, 33e-6, 20e-3, 5e-12) / 2

    # CMC common-mode equivalent
    Lcm = 3e-3
    k = 0.999
    Z_CMC = z_ind_parallel_c(s, Lcm*(1+k), 100e-3, 5e-12) / 2

    # Y caps at load node
    Z_Y_each = z_cap_esr_esl(s, 2.2e-9, 20e-3, 10e-9)
    Y5 = 2 / Z_Y_each

    # Recursive left-looking admittance
    Yleft1 = Y1
    Yb2 = 1 / (Z_line_input + 1/Yleft1)
    Yleft2 = Y2 + Yb2
    Yb3 = 1 / (Z_LDM1 + 1/Yleft2)
    Yleft3 = Yb3
    Yb4 = 1 / (Z_LDM2 + 1/Yleft3)
    Yleft4 = Yb4
    Yb5 = 1 / (Z_CMC + 1/Yleft4)
    Yleft5 = Y5 + Yb5

    I_total = 2*I_each
    V5 = I_total / Yleft5
    I_CMC = V5 * Yb5
    V4 = I_CMC / Yleft4
    I_LDM2 = V4 * Yb4
    V3 = I_LDM2 / Yleft3
    I_LDM1_cm = V3 * Yb3
    I_hot = I_LDM1_cm/2
    I_rtn = I_LDM1_cm/2
    I_sum = I_hot + I_rtn
    return I_hot, I_rtn, I_sum


def main():
    out = Path(__file__).resolve().parent
    f = np.logspace(2,8,1201)
    Ih, Ir, Icm = calc(f, 0.5)
    data = np.column_stack([
        f,
        np.abs(Ih),
        np.abs(Ir),
        np.abs(Icm),
        20*np.log10(np.abs(Ih)*1e6),
        20*np.log10(np.abs(Ir)*1e6),
        20*np.log10(np.abs(Icm)*1e6),
        np.angle(Icm, deg=True),
    ])
    csv = out/'cm_current_source_reference.csv'
    np.savetxt(csv, data, delimiter=',', header='frequency_Hz,I_HOT_abs_A,I_RTN_abs_A,I_CM_abs_A,I_HOT_dBuA,I_RTN_dBuA,I_CM_dBuA,I_CM_phase_deg', comments='')
    png = out/'cm_current_source_reference.png'
    plt.figure(figsize=(10,6))
    plt.semilogx(f, data[:,4], label='I_HOT at LDM1 input')
    plt.semilogx(f, data[:,5], label='I_RTN at LDM1 input')
    plt.semilogx(f, data[:,6], label='I_CM = I_HOT + I_RTN')
    plt.grid(True, which='both')
    plt.xlabel('Frequency [Hz]')
    plt.ylabel('Current [dBµA]')
    plt.title('Common-mode current at LDM1 input plane, two 0.5 A noise sources')
    plt.legend()
    plt.tight_layout()
    plt.savefig(png, dpi=180)
    print('CSV:', csv)
    print('PNG:', png)
    for ff in [100,400,1e3,10e3,100e3,1e6,10e6,100e6]:
        ih, ir, icm = calc(np.array([ff]), 0.5)
        print(f'{ff:>10.0f} Hz : I_HOT={abs(ih[0]):.6e} A, I_RTN={abs(ir[0]):.6e} A, I_CM={abs(icm[0]):.6e} A, I_CM={20*np.log10(abs(icm[0])*1e6):.3f} dBuA')

if __name__ == '__main__':
    main()
