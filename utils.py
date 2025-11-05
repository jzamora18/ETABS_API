import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def coef_import(GrupoUso):
    if GrupoUso == "I":
        return 1.0
    elif GrupoUso == "II":
        return 1.1
    elif GrupoUso == "III":
        return 1.25
    elif GrupoUso == "IV":
        return 1.5
    else:
        raise ValueError("Grupo de uso no válido")


def plot_espectro_nsr10(
    Aa,
    Av,
    Fa,
    Fv,
    T0,
    Tc,
    Tl,
    CoefImportancia,
    R,
    analisis_dinamico=False,
    tipo_espectro="diseño",
):
    T = np.linspace(0, 8, 500)
    Sa = np.zeros_like(T)
    Sa1 = np.full_like(T, np.nan)  # Use NaN for no values

    # Espectro de Diseño NSR-10
    for i, t in enumerate(T):
        if t < Tc:
            Sa[i] = 2.5 * Aa * Fa * CoefImportancia
        elif t < Tl:
            Sa[i] = 1.2 * Av * Fv * CoefImportancia / t
        else:
            Sa[i] = 1.2 * Av * Fv * Tl * CoefImportancia / t**2

    plt.plot(T, Sa / R, label=f"Espectro de {tipo_espectro}")

    return T, Sa


def k_value(t):
    if t <= 0.5:
        return 1
    elif 0.5 < t <= 2.5:
        return 0.75 + 0.5 * t
    elif t > 2.5:
        return 2
