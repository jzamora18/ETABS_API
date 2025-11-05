"""
Código para agilizar el analisis en ETABS mediante su API.
Las unidades por defecto son kN, m, C (SI).

"""

# 0. Importar librerias

import os
import sys
import pandas as pd
import numpy as np
import comtypes.client
import ctypes

import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

from utils import * 

# 1. Valores de entrada

g = 9.8067  # Gravedad m/s²
Norma = "NSR-10"
Suelo = "E"
GrupoUso = "I"
CoefImportancia = coef_import(GrupoUso)
Aa = 0.25
Av = 0.20
R = 7.0 # Coeficiente de reducción por ductilidad
H_total = 8.1
H_i = [3.1, 5.0]



if Norma == "NSR-10":    
    df_Fa = pd.read_excel("Tables NSR 10.xlsx", sheet_name="Tabla A.2.4-3_Fa", index_col=0)
    df_Fv = pd.read_excel("Tables NSR 10.xlsx", sheet_name="Tabla A.2.4-4_Fv", index_col=0)

    Fa = df_Fa.loc[Suelo, Aa]
    Fv = df_Fv.loc[Suelo, Av]

    T0 = 0.10 * (Av*Fv)/(Aa*Fa)
    Tc = 0.48 * (Av*Fv)/(Aa*Fa)
    Tl = 2.4*Fv

elif Norma == "MZSC-14":
    """
    Impementar lectura de tablas para MZSC-14
    """

    #df_Fa = pd.read_excel("Tables NSR 10.xlsx", sheet_name="MZSC-Espectro T corto", index_col=0)
    #df_Fv = pd.read_excel("Tables NSR 10.xlsx", sheet_name="MZSC-Espectro T largo", index_col=0)
    
    pass


plt.figure(figsize=(8, 5))

T, Sa = plot_espectro_nsr10(Aa, Av, Fa, Fv, T0, Tc, Tl, CoefImportancia, 1, tipo_espectro='diseño')
T, Sa_diseno = plot_espectro_nsr10(Aa, Av, Fa, Fv, T0, Tc, Tl, CoefImportancia, R, tipo_espectro='diseño reducido')

plt.xlabel('Periodo T [s]')
plt.ylabel('Aceleración Espectral Sa [g]')
plt.title('Espectro de Diseño Sísmico')
plt.grid(True)
plt.legend()
plt.gca().yaxis.set_major_locator(ticker.MultipleLocator(0.1))
plt.gca().xaxis.set_major_locator(ticker.MultipleLocator(0.5))
plt.xlim(0, 8)
plt.ylim(0, 1.1)
plt.show()

T_Sa_df = pd.DataFrame({'T': T, 'Sa': Sa, 'Sa_diseno': Sa_diseno})

Ta_df = pd.read_excel("Tables NSR 10.xlsx", sheet_name="Sistema estructural", index_col=0)

Sistema_estructural = "Pórticos de Concreto Resistente a Momento"
Ct = Ta_df.loc[Sistema_estructural, 'Ct']
a = Ta_df.loc[Sistema_estructural, 'a']

Ta = float(Ct * (H_total ** a))
Cu = max(1.2, 1.75-1.2*Av*Fv)
CuTa = Cu * Ta

# Directorio y nombre del modelo ETABS
ModelName = "Test.edb"

# Basic options to control ETABS instance

AttachToInstance = True  # False si el programa no está abierto

# Con respecto a qué instancia de ETABS usar
SpecifyPath = False
ProgramPath = R"C:\Program Files\Computers and Structures\ETABS 22\ETABS.exe"


APIPath = R"C:\Users\juanj\Documents\0. CODING\4_ETABS_API\Model"
if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass

ModelPath = APIPath + os.sep + ModelName

# create API helper object
helper = comtypes.client.CreateObject("ETABSv1.Helper")
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)

if AttachToInstance:
    # attach to a running instance of ETABS
    try:
        # get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject")
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.")
        sys.exit(-1)
else:
    if SpecifyPath:
        try:
            #'create an instance of the ETABS object from the specified path
            myETABSObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program from " + ProgramPath)
            sys.exit(-1)
    else:
        try:
            # create an instance of the ETABS object from the latest installed ETABS
            myETABSObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")
        except (OSError, comtypes.COMError):
            print("Cannot start a new instance of the program.")
            sys.exit(-1)

    # start ETABS application
    myETABSObject.ApplicationStart()

# create SapModel object
SapModel = myETABSObject.SapModel

#initialize model
kN_m_C = 6  # Units kN, m, C
kgf_m_C = 8 # Units kgf, m, C
Ton_m_C = 12 # Units Ton, m, C
ret = SapModel.SetPresentUnits(kN_m_C)

def save_and_unlock():
    ret = SapModel.File.Save(ModelPath)
    ret = SapModel.SetModelIsLocked(False)

def save_and_run():
    ret = SapModel.File.Save(ModelPath)
    ret = SapModel.Analyze.RunAnalysis()

save_and_unlock()

spectrum_name = f"Espectro Prueba"

ret = SapModel.Func.GetNameList()

if spectrum_name in ret[1]:
    print(f"El espectro '{spectrum_name}' ya existe.")

else:
    T_values = T_Sa_df['T'].values
    Sa_values = T_Sa_df['Sa_diseno'].values

    # Prepare data for ETABS
    periods, accelerations = zip(*sorted(zip(T_values, Sa_values)))

    db_tables = SapModel.DatabaseTables
    table_key = "Functions - Response Spectrum - User Defined"
    ret01 = db_tables.GetAllFieldsInTable(table_key)

    fields = ["Name", "Period", "Value", "DampingRatio"]

    # Create table data - each record is a complete row
    table_data = []
    for period, accel in zip(periods, accelerations):
        table_data.append([
            spectrum_name,      # Name
            str(period),        # Period sec
            str(accel),         # Value
            "0.05",             # Damping Ratio (5%)
        ])

    flattened_data = [item for row in table_data for item in row]

    ret02 = db_tables.SetTableForEditingArray( table_key, 1, fields, len(table_data), flattened_data)
    ret03 = db_tables.ApplyEditedTables(True)

save_and_unlock()

spectrum_name = f"Espectro Prueba"

ret = SapModel.Func.GetNameList()

if spectrum_name in ret[1]:
    print(f"El espectro '{spectrum_name}' ya existe.")

else:
    T_values = T_Sa_df['T'].values
    Sa_values = T_Sa_df['Sa_diseno'].values

    # Prepare data for ETABS
    periods, accelerations = zip(*sorted(zip(T_values, Sa_values)))

    db_tables = SapModel.DatabaseTables
    table_key = "Functions - Response Spectrum - User Defined"
    ret01 = db_tables.GetAllFieldsInTable(table_key)

    fields = ["Name", "Period", "Value", "DampingRatio"]

    # Create table data - each record is a complete row
    table_data = []
    for period, accel in zip(periods, accelerations):
        table_data.append([
            spectrum_name,      # Name
            str(period),        # Period sec
            str(accel),         # Value
            "0.05",             # Damping Ratio (5%)
        ])

    flattened_data = [item for row in table_data for item in row]

    ret02 = db_tables.SetTableForEditingArray( table_key, 1, fields, len(table_data), flattened_data)
    ret03 = db_tables.ApplyEditedTables(True)

def extract_values(name):
    table = SapModel.DatabaseTables.GetTableForDisplayArray(name, GroupName="")
    df = pd.DataFrame(np.array_split(table[4], table[3]))
    df.columns = table[2] # Ajustar los nombres de las columnas
    return df

save_and_run()

# 1. Masa del modelo
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetComboSelectedForOutput("MASA", Selected=True)

table_mass = "Story Forces"
table_mass_df = extract_values(table_mass)
table_mass_df = table_mass_df[
    (table_mass_df["Location"] == "Bottom") & (table_mass_df["OutputCase"] == "MASA")
].reset_index(drop=True)

table_mass_df


columns_to_drop = ["OutputCase", "CaseType", "StepType", "StepNumber", "Location", "VX", "VY", "T", "MX", "MY"]
table_mass_df = table_mass_df.drop(columns=columns_to_drop)

# Ensure P is numeric
table_mass_df["P"] = pd.to_numeric(table_mass_df["P"], errors="coerce")

# Calculate P_i as desired
table_mass_df["Masa_i"] = table_mass_df["P"].diff()
table_mass_df.loc[0, "Masa_i"] = table_mass_df.loc[0, "P"]

table_mass_df

# 2. Modal Participating Mass Ratios
SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
SapModel.Results.Setup.SetCaseSelectedForOutput("Modal")

table_periods = "Modal Participating Mass Ratios"
table_periods_df = extract_values(table_periods)
table_periods_df.head()

table_periods_df["SumUX"] = pd.to_numeric(table_periods_df["SumUX"], errors="coerce")
tx_row = table_periods_df.loc[table_periods_df["SumUX"] > 0.9].iloc[0]
tx = float(tx_row["Period"])

table_periods_df["SumUY"] = pd.to_numeric(table_periods_df["SumUY"], errors="coerce")
ty_row = table_periods_df.loc[table_periods_df["SumUY"] > 0.9].iloc[0]
ty = float(ty_row["Period"])

# 3. Fuerza Horizontal Equivalente
tx = min(tx, CuTa)
ty = min(ty, CuTa)

Sa_x = T_Sa_df.loc[T_Sa_df['T'] >= tx, 'Sa_diseno'].iloc[0]
Sa_y = T_Sa_df.loc[T_Sa_df['T'] >= ty, 'Sa_diseno'].iloc[0]

kx = k_value(tx)
ky = k_value(ty)

FHE_df = table_mass_df.copy().drop(columns=["P"])
FHE_df["Masa_acum_i"] = FHE_df["Masa_i"].cumsum()
FHE_df["h_i"] = [3.1, 5]
FHE_df["Masa_acum_i * h_i ^ k"] = FHE_df["Masa_acum_i"] * FHE_df["h_i"]**kx

Masa_total = FHE_df["Masa_i"].sum()
Masa_factor = FHE_df["Masa_acum_i * h_i ^ k"].sum()

FHE_df["Cv"] = FHE_df["Masa_acum_i * h_i ^ k"] / Masa_factor
Cv_total = FHE_df["Cv"].sum().round(3)

if Cv_total != 1:
    print("Revisar cálculos, masas y factores. Cv_total != 1")

Vb_X_FHE = Sa_x * Masa_total
Vb_Y_FHE = Sa_y * Masa_total

FHE_df

coef_regularidad = 0.8
print(f'Vb_FHE = {coef_regularidad}*Vb_FHE')
Vb_X_FHE *= coef_regularidad
Vb_Y_FHE *= coef_regularidad

# Dynamic forces
ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
ret = SapModel.Results.Setup.SetComboSelectedForOutput("MASA", Selected=True)

table_base_reaction = "Base Reactions"
table_base_reaction_df = extract_values(table_base_reaction)

columns_to_drop = ["StepNumber","CaseType", "StepType", "MX", "MY", "MZ", "X", "Y", "Z"]
table_base_reaction_df = table_base_reaction_df.drop(columns=columns_to_drop)

Vb_X_din = float(table_base_reaction_df.loc[table_base_reaction_df["OutputCase"] == "SX", "FX"].iloc[0])
Vb_Y_din = float(table_base_reaction_df.loc[table_base_reaction_df["OutputCase"] == "SY", "FY"].iloc[0])

factor_ajuste_X = max(1, Vb_X_FHE / Vb_X_din)
factor_ajuste_Y = max(1, Vb_Y_FHE / Vb_Y_din)

print(f"Cortante basal FHE X: {Vb_X_FHE:.2f} kN, Cortante basal dinámico X: {Vb_X_din:.2f} kN, Factor de ajuste: {factor_ajuste_X:.2f}")
print(f"Cortante basal FHE Y: {Vb_Y_FHE:.2f} kN, Cortante basal dinámico Y: {Vb_Y_din:.2f} kN, Factor de ajuste: {factor_ajuste_Y:.2f}")

save_and_unlock()
# ret = SapModel.Analyze.GetCaseStatus()
# print(ret[1])

ret = SapModel.LoadCases.ResponseSpectrum.GetLoads("SX")
print(ret)

ret = SapModel.LoadCases.ResponseSpectrum.SetLoads(
    "SX",
    1,
    ('U1',),
    (spectrum_name,),
    (g*factor_ajuste_X,),  # SF
    ('Global',),  # CSys
    (0.0,),  # Ang
)

ret = SapModel.LoadCases.ResponseSpectrum.GetLoads("SX")
print(ret)

print("="*70)

ret = SapModel.LoadCases.ResponseSpectrum.GetLoads("SY")
print(ret)

ret = SapModel.LoadCases.ResponseSpectrum.SetLoads(
    "SY",
    1,
    ('U2',),
    (spectrum_name,),
    (g*factor_ajuste_Y,),  # SF
    ('Global',),  # CSys
    (0.0,),  # Ang
)

ret = SapModel.LoadCases.ResponseSpectrum.GetLoads("SY")
print(ret)

print("="*70)

save_and_run()

ret = SapModel.SelectObj.Group("Columnas")

ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput
ret = SapModel.Results.Setup.SetCaseSelectedForOutput("SX")

table_drifts = "Joint Drifts"
table_drifts_df = extract_values(table_drifts)
table_drifts_df = table_drifts_df[(table_drifts_df["OutputCase"] == "SX") | (table_drifts_df["OutputCase"] == "SY")].reset_index(drop=True)
table_drifts_df

