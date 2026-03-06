# PLAXIS Output Export Tool

This repository contains a Windows-based export tool for extracting structural and node-response results from the PLAXIS Output API and packaging them into Excel workbooks and PNG plots.

The project is built for repetitive post-processing of dynamic analyses where phases must be grouped by loading direction, structural elements must be compared across many earthquakes, and node response histories must be converted into spectra in a consistent format.

## What the tool does

The software provides two main workflows:

1. Structural moment export
2. Node multi-phase spectrum export

Both workflows connect to the PLAXIS Output remote scripting server, read selected phases and objects, and generate ready-to-review deliverables.

## Main capabilities

### Structural moment workflow

- Loads phases directly from the PLAXIS Output API
- Splits candidate phases into X and Y groups using regex patterns
- Lets the user manually finalize X and Y phase selections
- Loads structural objects from the model:
  - Embedded beams
  - Plate elements
- Supports two separate plate groups
- Supports optional "merge as single profile" behavior for each plate group
- Reads `MEnvelopeMax2D` and `MEnvelopeMin2D` for every selected object and phase
- Normalizes depth from top to bottom
- Averages `M+` and `M-` independently by direction
- Writes Excel sheets in both long and wide formats
- Embeds Excel charts directly in the workbook
- Exports PNG charts with configurable DPI

### Node spectrum workflow

- Loads CurvePoints directly from the PLAXIS Output API
- Runs across many selected dynamic phases
- Reads acceleration time histories for each selected node and phase
- Uses the dynamic time vector from PLAXIS for spectrum calculation
- Separates X and Y phase groups
- Computes per-phase spectra and mean spectra
- Writes long-format and wide-format Excel sheets
- Embeds Excel charts directly in the workbook
- Exports PNG plots with configurable DPI
- Optionally saves per-phase node time-history CSV files into phase-named subfolders

## Typical use case

This tool is intended for cases such as:

- Reviewing pile moment distributions for many earthquake records
- Comparing wall and plate response by direction
- Exporting node acceleration histories and response spectra for selected monitoring points
- Producing a single workbook that can be reviewed without re-querying PLAXIS

## Repository contents

- `plaxis_export_gui.py`
  - Tkinter GUI for the main API-based workflow
- `export_plaxis_data.py`
  - Core data access, post-processing, Excel writing, and plot generation
- `run_plaxis_multiphase_cli.py`
  - Command-line wrapper for batch execution

## Requirements

### Environment

- Windows
- Python 3
- PLAXIS Output with remote scripting enabled
- Access to the correct Output port and password

### Python packages

Install the required packages with:

```powershell
pip install numpy pandas openpyxl matplotlib plxscripting pywinauto
```

Notes:

- `pywinauto` is only needed for legacy GUI-copy helper code that still exists in the repository.
- The main structural and node export workflows use the Output API.

## GUI usage

Start the GUI with:

```powershell
python plaxis_export_gui.py
```

### GUI flow

1. Enter the PLAXIS Output host, port, password, and output workbook paths.
2. Define X and Y phase regex patterns.
3. Click `Load Phases`.
4. Review and adjust the X and Y phase selections manually.
5. Click `Load Structural Objects` if you want moment export.
6. Click `Load CurvePoints` if you want node spectrum export.
7. Select the required piles, plate groups, and/or nodes.
8. Set spectrum parameters and PNG DPI as needed.
9. Run either:
   - `Run Structural Moment Analysis`
   - `Run Node Spectrum Analysis`

### GUI options worth noting

- `PNG DPI`
  - Controls PNG export resolution
- `Merge as single profile`
  - Treats the selected plate group as one continuous profile
- `Save node time histories into phase subfolders`
  - Writes CSV time histories under a folder next to the node workbook

Example time-history folder layout:

```text
PLAXIS_multiphase_node_results_time_history/
  DD2_X_20030501002708_1201_H2/
    Node_24388_22_90_20_95.csv
```

The phase folder name is based on the visible phase name, not the internal suffix such as `[Phase_6]`.

## CLI usage

The CLI supports two modes:

- `node`
- `structural`

### Node export example

```powershell
python run_plaxis_multiphase_cli.py node ^
  --host localhost ^
  --port 10000 ^
  --password "YOUR_PASSWORD" ^
  --x-regex "^DD2_X_.*" ^
  --y-regex "^DD2_Y_.*" ^
  --out "C:\temp\node_results.xlsx" ^
  --result-type "Soil.Ax" ^
  --time-col "DynamicTime" ^
  --damping 0.05 ^
  --period-start 0.01 ^
  --period-end 3.0 ^
  --period-step 0.01 ^
  --plot-dpi 180 ^
  --curvepoint-regex "Node 4490" ^
  --save-node-timehistory-subfolders
```

### Structural export example

```powershell
python run_plaxis_multiphase_cli.py structural ^
  --host localhost ^
  --port 10000 ^
  --password "YOUR_PASSWORD" ^
  --x-regex "^DD2_X_.*" ^
  --y-regex "^DD2_Y_.*" ^
  --out "C:\temp\structural_results.xlsx" ^
  --plot-dpi 180 ^
  --pile-regex "EmbeddedBeam_" ^
  --plate1-regex "Plate_" ^
  --plate1-merge-single-profile
```

## Output structure

The tool intentionally writes analysis outputs into separate workbooks:

- one workbook for structural moment export
- one workbook for node spectrum export

### Structural workbook

Typical sheets:

- `Phases`
- `Selections`
- `MomentRawLong`
- `MomentAvgByDir`
- `MomentWide_*`
- `_Status`

`MomentWide_*` sheets are arranged in object-based blocks so depth and moment columns stay aligned and readable.

### Node workbook

Typical sheets:

- `Phases`
- `Selections`
- `NodeTimeHistoryLong`
- `NodeSpectrumLong`
- `NodeSpectrumMean`
- `Spec_*`
- `SpecPhase_*`
- `SpecMean_*`
- `_Status`

For wide spectrum sheets, the first column is `Period_s` and the following columns are phase-based response values. Excel charts are written into the workbook as native charts, not just linked images.

## Plot and Excel behavior

- PNG charts use compact legends to reduce wasted plot area
- Excel charts use numeric axes with automatic axis scaling
- Structural charts are written from aligned x-y data blocks rather than sparse mixed columns
- Node and structural plots are exported both as workbook charts and as PNG files

## Public API functions

The main reusable entry points in `export_plaxis_data.py` are:

- `list_phases_api(host, port, password)`
- `list_structural_objects_api(host, port, password)`
- `run_structural_moment_export(args, logger=print)`
- `run_node_multiphase_spectrum_export(args, logger=print)`

These functions are intended for reuse from the GUI, the CLI, or custom automation.

## Limitations and assumptions

- Phase direction grouping is regex-driven first, then manually confirmed by the user
- Structural summaries are based on averaged `M+` and `M-`, not absolute envelopes
- Plate groups are manually defined; there is no automatic wall/foundation classification
- The tool expects PLAXIS Output API access to be available on the selected port

## Recommended review order

When checking a new export, start with:

1. `Selections`
2. `Phases`
3. `_Status`

This makes it easier to verify that the correct phases, objects, and nodes were exported before reviewing charts.
