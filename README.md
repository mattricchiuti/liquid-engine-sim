# Liquid Engine Sim

A time-stepping performance model of a self-pressurized **N2O / liquid-fuel** rocket motor
(isopropanol by default), written in Python. Given the motor's geometry and propellant
conditions, it predicts the thrust curve, chamber and feed pressures, mass flows, O/F,
Isp and injector stiffness over the whole burn, and can sweep any input across a range.

It started as a replacement for the HalfCatSim spreadsheet and keeps a few of its
conventions (blowdown model, startup ramp), but models the injector, feed system,
combustion and vapor phase in more detail.

## Quick start

Requires Python 3.12+ (developed on 3.14).

```bash
pip install -r requirements.txt
python sim.py
```

1. Edit the inputs in **`input.xlsx`** (LibreOffice Calc or Excel) and **save** it.
2. Run `sim.py`. Results are written to `outputs/` and open in your browser.

The spreadsheet can stay open while you run the sim; Python reads the saved file.

## Inputs (`input.xlsx`)

Everything lives on the `input` sheet, in two panels (simulation setup and geometry on
the left, propellant properties, efficiencies and feed losses on the right).

- **Units:** type each value in either the Imperial or the Metric column. The sim uses
  whichever column holds a typed number (not a formula) and converts it to SI using the
  unit written beside it. Unknown units stop the run with an error naming the parameter.
- **Names matter:** parameters are found by the name in their label cell, so rows can be
  moved freely, but renaming a label renames the input.
- **Decimals, not percents:** efficiencies and discharge coefficients are fractions
  (C* efficiency `0.65`, not `65`).
- **Formulas:** calculated cells (e.g. oxidizer and fuel column lengths) are read as their
  last calculated value. They are *not* recalculated during sweeps (see Limitations).

### Sweeps

The **Sweep Setup** block at the top has two rows:

| Parameter | min | max | step | 0/1 |
|---|---|---|---|---|
| pick from the dropdown | in that parameter's units | | | 1 = on |

- One active row runs the sim for each value (both ends included).
- Two active rows run every combination (a grid).
- The dropdown's list lives on the `lists` sheet; keep it in sync if you rename inputs.

## Outputs (`outputs/`)

Interactive pages (hover for values, drag to zoom, double-click to reset, click legend
entries to hide/show series). They work offline and can be shared as single files.

| File | When | Contents |
|---|---|---|
| `results.html` | single run | thrust, pressures (tanks, manifolds, chamber), injector stiffness, mass flows, O/F, Isp, propellant left, c*, chamber temperature, nozzle exit pressure vs. time |
| `sweep.html` | one sweep | curves over time for every case (thrust, Pc, O/F, both stiffnesses) + summary metrics vs. the swept value |
| `sweep.html` | two sweeps | heat maps of total impulse, Isp, peak thrust, burn time, minimum stiffness |
| `sweep_results.csv` | any sweep | one row per case: swept values and summary numbers |

The console prints progress: percent of propellant used for a single run, and
`[n/N  xx%, ~time left]` per case for sweeps. Plots are dark mode by default; set
`THEME = "light"` in `sim.py` for light.

## What it models

Each time step (10 ms, with 1 ms steps during startup):

1. **Tank pressures** — HalfCat's empirical blowdown: N2O tank pressure falls linearly with
   oxidizer mass from its saturation pressure to `Decay Constant (liquid)` × P0, then
   decays exponentially once the liquid is gone (below 10% of the starting mass). The fuel
   tank is driven by a piston: fuel pressure = oxidizer pressure − piston loss.
2. **Feed lines** — friction (Darcy, Haaland friction factor) plus fittings as a sum of
   K-factors (ball valve from its Cv, elbows, tank entrance), treated as resistances in series.
3. **Injectors**
   - *Fuel:* single-phase incompressible orifice flow, solved in closed form with the line.
   - *Oxidizer, liquid phase:* the **NHNE** (non-homogeneous, non-equilibrium, Dyer) model,
     blending single-phase incompressible and homogeneous-equilibrium flow (with choking),
     using CoolProp N2O properties. The manifold pressure is found by bisection so the line
     and injector flows match.
   - *Oxidizer, vapor phase:* compressible gas orifice flow (choked or subsonic), using an
     effective isentropic exponent k = ρa²/P from CoolProp, since saturated N2O vapor is
     far from ideal.
4. **Chamber** — chamber pressure is the value where injector flow in equals nozzle flow
   out (Pc·At / c*), found by bisection. Combustion properties (c*, Tc, MW, γ) come from
   **NASA CEA** (via RocketCEA), precomputed into a lookup table over chamber pressure and
   O/F and interpolated. C* efficiency scales c* by η and chamber temperature by η².
5. **Nozzle and thrust** — isentropic expansion to the exit area ratio, nozzle efficiency on
   exhaust velocity, conical divergence loss (1 + cos α)/2, and pressure thrust
   (Pe − Pa)·Ae. Flags possible flow separation when Pe < 0.4·Pa.
6. **Startup** — propellant flow is assumed established before ignition; chamber pressure
   ramps linearly from ambient to its balanced value over the first 10 ms (after HalfCat).

The burn ends when either tank is empty, the fuel tank pressure reaches ambient, or the
sim duration is reached.

## Assumptions and limitations

- **Tank model is empirical.** The blowdown curve uses HalfCat's fitted decay constants,
  not a thermodynamic model of the boiling N2O; tank temperature isn't tracked. Its accuracy
  away from the conditions it was fitted for (e.g. very cold or warm N2O) is uncertain.
- **Oxidizer section is assumed full of liquid** at the start (no vapor space).
- **Liquid-to-vapor switch is instantaneous**, so the thrust step at that point is sharper
  than in a real motor.
- **Fuel-rich vapor tail:** O/F drops below 1 in the tail; equilibrium CEA is optimistic
  there, and the nozzle may separate, so tail impulse is approximate.
- **Quasi-steady chamber:** no chamber filling dynamics or combustion instability modelling.
  Injector stiffness is reported so stability margins can be judged.
- **Spreadsheet formulas aren't recalculated in sweeps.** Sweeping an input that a sheet
  formula depends on (e.g. tank length → fuel length) won't update the formula's result.
  Only directly typed inputs appear in the sweep dropdown.
- **Not validated against test data.** Cd and C* efficiency in particular are estimates;
  a cold-flow test (Cd) and a static fire (C* efficiency) would pin them down.

## Code layout

Everything is in **`sim.py`**, in order: input loading, initial conditions, tank model,
feed and injector models, CEA table and combustion, chamber balance, nozzle, the time loop
(`run`), sweeps, and plotting. Each physics piece is a separate function, so it can be
tested on its own or swapped for a better model.

Handy constants near the top: `GAS_PHASE_FRACTION`, `OF_MIN`/`OF_MAX`, the CEA table grid,
`SEPARATION_RATIO`, `STARTUP_TIME`/`STARTUP_DT`; near the plotting code: `THEME`,
`OPEN_IN_BROWSER`.

## Notes

- RocketCEA keeps its working files in `C:\UserData\RocketCEA` when the Windows user folder
  path contains a space; it prints a notice about this on every run. Harmless.
- CoolProp has no N2O viscosity model, so liquid N2O viscosity is a spreadsheet input and
  vapor viscosity is a constant in the code (it only affects line friction).

## References

- J. Dyer et al., "Modeling Feed System Flow Physics for Self-Pressurizing Propellants,"
  AIAA 2007-5702.
- B. J. Solomon, "Engineering Model to Calculate Mass Flow Rate of a Two-Phase Saturated
  Fluid Through an Injector Orifice," M.S. thesis, Utah State University, 2011.
- B. S. Waxman, "An Investigation of Injectors for Use with High Vapor Pressure Propellants
  with Applications to Hybrid Rockets," Ph.D. dissertation, Stanford University, 2014.
- HalfCatSim v1.3.10 (spreadsheet the model was originally benchmarked against).
- NASA CEA via [RocketCEA](https://rocketcea.readthedocs.io/); N2O properties via
  [CoolProp](http://www.coolprop.org/).
