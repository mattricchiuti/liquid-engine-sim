import math

import matplotlib.pyplot as plt                   # plotting
import numpy as np                                # arrays and math on whole time histories
import openpyxl                                   # reading input.xlsx
from CoolProp.CoolProp import PropsSI             # N2O properties: PropsSI(output, in1, v1, in2, v2, "NitrousOxide")
from rocketcea.cea_obj import CEA_Obj             # NASA CEA: c*, chamber temp, gamma, MW vs. Pc and O/F
from scipy.optimize import brentq                 # root finding, e.g. the chamber pressure / mass flow balance

INPUT_FILE = "input.xlsx"

# Column letters for each input panel on the sheet: parameter name, then (value, unit) pairs
PANELS = [
    {"name": "C", "imperial": ("D", "E"), "metric": ("F", "G")},
    {"name": "L", "imperial": ("M", "N"), "metric": ("O", "P")},
]

# Unit written in the spreadsheet -> function converting a value in that unit to SI
TO_SI = {
    "": lambda x: x,                          # unitless (counts, Cd, Cv, efficiencies)
    "s": lambda x: x,
    "1/s": lambda x: x,
    "m": lambda x: x,
    "Pa": lambda x: x,
    "kg/m3": lambda x: x,
    "in": lambda x: x * 0.0254,               # -> m
    "ft": lambda x: x * 0.3048,               # -> m
    "psi": lambda x: x * 6894.757,            # -> Pa
    "F": lambda x: (x - 32) / 1.8 + 273.15,   # -> K
    "deg": math.radians,                      # -> rad
    "lbm/ft3": lambda x: x * 16.01846,        # -> kg/m^3
    "cP": lambda x: x * 0.001,                # -> Pa*s
}


# ---------- inputs ----------
def load_inputs(path=INPUT_FILE):
    """Read the 'input' sheet.

    Returns (inputs, sweeps):
      inputs: {parameter name: value in SI}
      sweeps: list of active sweeps, each {"parameter": name, "values": [values in SI]}

    For each parameter, uses whichever column (Imperial or Metric) holds a
    typed value rather than a formula, preferring Imperial if both do; if both
    are formulas, uses the calculated value. Converts using the unit written
    next to the value. An unrecognized unit is an error.
    """
    # data_only=True gives the values formulas last calculated to; the second
    # copy keeps the formula text, so typed values can be told apart from formulas
    ws = openpyxl.load_workbook(path, data_only=True)["input"]
    ws_formulas = openpyxl.load_workbook(path)["input"]
    inputs = {"Oxidizer": "N2O"}  # the tank model only handles self-pressurizing N2O
    units = {}  # unit each parameter was read in, for converting sweep values

    def is_typed_number(cell_ref):
        return isinstance(ws_formulas[cell_ref].value, (int, float))

    # Fuel name: the dropdown to the right of the "Fuel Properties" heading
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Fuel Properties":
                inputs["Fuel"] = ws.cell(cell.row, cell.column + 1).value

    # Sweep block: rows under the "Parameter | min | max | step | 0/1" header
    sweep_rows = []
    for row in range(1, ws.max_row + 1):
        if ws[f"C{row}"].value == "Parameter" and ws[f"D{row}"].value == "min":
            r = row + 1
            while isinstance(ws[f"G{r}"].value, (int, float)):  # each sweep row has a 0/1 toggle
                sweep_rows.append(r)
                r += 1

    unknown_units = []
    for col in PANELS:
        for row in range(1, ws.max_row + 1):
            if col["name"] == "C" and row in sweep_rows:
                continue
            label = ws[f"{col['name']}{row}"].value
            columns = (col["imperial"], col["metric"])
            # A typed value wins; if both columns are formulas (e.g. Oxidizer Length), use the calculated one
            source = next((c for c in columns if is_typed_number(f"{c[0]}{row}")), None) \
                or next((c for c in columns if isinstance(ws[f"{c[0]}{row}"].value, (int, float))), None)
            if not isinstance(label, str) or source is None:
                continue  # section titles, column headers, blank rows
            value = ws[f"{source[0]}{row}"].value
            unit = (ws[f"{source[1]}{row}"].value or "").strip()
            label = label.strip()
            if label in inputs:
                raise ValueError(f"Duplicate parameter name {label!r} in {path}")
            if unit not in TO_SI:
                unknown_units.append(f"{label} ({unit!r}, row {row})")
                continue
            inputs[label] = TO_SI[unit](value)
            units[label] = unit

    if unknown_units:
        raise ValueError(f"No SI conversion for the unit of: " + "; ".join(unknown_units))

    sweeps = []
    for row in sweep_rows:
        if ws[f"G{row}"].value != 1:
            continue  # sweep turned off
        name = (ws[f"C{row}"].value or "").strip()
        lo, hi, step = (ws[f"{c}{row}"].value for c in "DEF")
        if name not in inputs:
            raise ValueError(f"Sweep in row {row} is on but its parameter {name!r} isn't an input")
        if not all(isinstance(v, (int, float)) for v in (lo, hi, step)) or step <= 0 or hi < lo:
            raise ValueError(f"Sweep in row {row} ({name}) needs numbers with min <= max and step > 0")
        n = round((hi - lo) / step) + 1  # both ends included
        values = [TO_SI[units[name]](lo + i * step) for i in range(n)]
        sweeps.append({"parameter": name, "values": values})

    return inputs, sweeps


def area(diameter):
    return math.pi / 4 * (diameter ** 2)

def initial_conditions(inputs):

    tank_ID = inputs["Tank OD"] - 2 * inputs["Tank Wall Thickness"]
    tank_area = math.pi / 4 * tank_ID ** 2

    T_ox = inputs["Nitrous Temperature"]

    rho_ox = PropsSI("D", "T", T_ox, "Q", 0, inputs["Oxidizer"])
    V_ox = inputs["Oxidizer Length"] * tank_area
    m_ox = rho_ox * V_ox

    rho_fuel = inputs["Fuel Density"]
    V_fuel = inputs["Fuel Length"] * tank_area
    m_fuel = rho_fuel * V_fuel

    P_tank0 = PropsSI("P", "T", T_ox, "Q", 0, inputs["Oxidizer"])

    A_t = area(inputs["Throat Diameter"])
    A_e = area(inputs["Exit Diameter"])
    A_oxfeed = area(inputs["Ox Line Diameter"])
    A_fuelfeed = area(inputs["Fuel Line Diameter"])

    CdA_ox = inputs["No. Oxidizer Holes"] * area(inputs["Oxidizer Hole Diameter"]) * inputs["Oxidizer Cd"]
    CdA_fuel = inputs["No. Fuel Holes"] * area(inputs["Fuel Hole Diameter"]) * inputs["Fuel Cd"]

    return {
        "T_ox": T_ox,
        "rho_ox": rho_ox,
        "rho_fuel": rho_fuel,
        "V_ox": V_ox,
        "V_fuel": V_fuel,
        "m_ox": m_ox,
        "m_fuel": m_fuel,
        "P_tank0": P_tank0,
        "A_t": A_t,
        "A_e": A_e,
        "A_oxfeed": A_oxfeed,
        "A_fuelfeed": A_fuelfeed,
        "CdA_ox": CdA_ox,
        "CdA_fuel": CdA_fuel,
    }


if __name__ == "__main__":
    inputs, sweeps = load_inputs()
    for name, value in inputs.items():
        print(f"{name:28} {value:.6g}" if isinstance(value, float) else f"{name:28} {value}")

    print()
    for sweep in sweeps:
        print(f"Sweep {sweep['parameter']}: {len(sweep['values'])} values, "
              f"{sweep['values'][0]:.6g} to {sweep['values'][-1]:.6g} (SI)")
    if not sweeps:
        print("No sweeps active")

    ic = initial_conditions(inputs)
    print()
    for name, value in ic.items():
        print(f"{name:28} {value:.6g}")
