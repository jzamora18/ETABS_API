import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def coef_import(use_group):
    if use_group == "I":
        return 1.0
    elif use_group == "II":
        return 1.1
    elif use_group == "III":
        return 1.25
    elif use_group == "IV":
        return 1.5
    else:
        raise ValueError("Group of use no valid")


def plot_spectrum(
    Aa,
    Av,
    Fa,
    Fv,
    T0,
    Tc,
    Tl,
    importance_factor,
    R,
    spectrum_type="design",
):
    T = np.linspace(0, 8, 500)
    Sa = np.zeros_like(T)
    Sa1 = np.full_like(T, np.nan)  # Use NaN for no values

    # Design spectrum calculation based on NSR-10
    for i, t in enumerate(T):
        if t < Tc:
            Sa[i] = 2.5 * Aa * Fa * importance_factor
        elif t < Tl:
            Sa[i] = 1.2 * Av * Fv * importance_factor / t
        else:
            Sa[i] = 1.2 * Av * Fv * Tl * importance_factor / t**2

    plt.plot(T, Sa / R, label=f"Spectrum type: {spectrum_type}")

    return T, Sa


def k_value(t):
    if t <= 0.5:
        return 1
    elif 0.5 < t <= 2.5:
        return 0.75 + 0.5 * t
    elif t > 2.5:
        return 2
