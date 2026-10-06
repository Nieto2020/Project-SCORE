# Copilot Instructions — Project-SCORE

Analysis scripts for **Catch Consulting's** bimonthly labor-market survey. The scripts read the
Qualtrics export `BASE HISTORICA.xlsx` and produce console reports and an Excel column catalog.
There is no package, no build step, and no application runtime — the deliverable is the printed report.

## Commands

Requires a Python interpreter with `pandas` and `openpyxl` installed globally (verified against
pandas 3.0.6 / openpyxl 3.1.5). There is no virtualenv, no `requirements.txt`, and no lockfile.
Run both scripts **from the repository root** — input and output filenames are hardcoded as relative paths.

```powershell
python informe_base_historica.py   # main deliverable: prints the 9-section report to stdout
python test.py                     # rewrites columnas_por_tipo.xlsx (column catalog by dtype)
```

Notes:

* There are **no automated tests** and no test framework configured (no pytest, no test files).
  `test.py` is an exploratory script, not a suite — running it regenerates an artifact, it does not
  verify anything. Do not describe it as a test run.
* No linter, formatter, or type checker is configured (no ruff/black/flake8/mypy, no config files).
* To validate a change, run `python informe_base_historica.py` and read the printed sections for the
  affected area — that output is the only regression signal available.

### Excel MCP server

[.vscode/mcp.json](../.vscode/mcp.json) registers the `excel` MCP server (`excel-mcp-server`, installed
with `pip install --user excel-mcp-server`) confined to this workspace via `--allow-dir
${workspaceFolder}` and running with `--read-only`, which registers only the 7 non-mutating tools
(`describe_workbook`, `describe_sheet`, `list_workbooks`, `read_range`, `find_cells`, `export_workbook`,
`read_vba`). It is for **inspection only** — use it to look at `BASE HISTORICA.xlsx` structure
(e.g. `describe_workbook`, `describe_sheet`) instead of writing a throwaway pandas script.

Use the MCP server to *look*; use pandas to *compute*. Reading 40 companies × 7 sheets × ~1150 columns
cell by cell through `read_range` is impractical, and `read_range` returns dates as ISO 8601 strings,
which would lose the dtypes that `test.py` catalogs. All computation belongs in
[informe_base_historica.py](../informe_base_historica.py) or [test.py](../test.py).

## Architecture

Two **standalone scripts with no shared module and no cross-imports**. Helpers (`find_col`, `limpiar`,
`to_num`, `fmt_num`, `_normalizar_texto`, `binario_si_no`) exist only in
[informe_base_historica.py](../informe_base_historica.py); [test.py](../test.py) is self-contained with
its own `_nombre_hoja_valido`. Factor out a helper only if a third script needs it.

### Data layout

`BASE HISTORICA.xlsx` is a wide Qualtrics export:

* **One worksheet per bimonthly period** (currently 7: `JUL-AGO_2024` … `JUL-AGO_2025`).
* ~1142–1169 columns × 40 companies + 1 Qualtrics sub-header row.
* Periods are discovered dynamically from `xl.sheet_names` — no period list is hardcoded, so adding a
  worksheet automatically includes it in every aggregation.
* Company identifier column is `Compañía`. Company names are **already anonymized** (`COMPAÑIA 1` …
  `COMPAÑIA 40`); the committed data contains no real client names.
* Row 0 of each sheet is a Qualtrics sub-header containing `Open-Ended Response` / `Response`;
  `limpiar()` drops it along with rows whose `Compañía` is blank. This yields a balanced panel of
  40 companies across 7 periods (280 records).

### Column resolution is by phrase, never by position

Headers are long Spanish survey prompts (e.g. `Por favor, capture el Salario Diario Inicial ofrecido
a su personal operativo...`). Which column holds a given metric is resolved at runtime by
case-insensitive substring match:

```python
find_col(df, "salario diario inicial")   # -> first column whose name contains the phrase
```

The logical-key → phrase maps `METRICAS`, `CATEGORICAS`, `PRESTACIONES_AMBIENTE` and
`BONO_PERMANENCIA_FRASE` at the top of [informe_base_historica.py](../informe_base_historica.py) are the
single source of truth. Each phrase currently matches **exactly one** column. Positions are **not**
stable: later sheets append unnamed trailing columns (`Unnamed: 1142.1`, `Unnamed: 1158`, …), so
positional indexing would silently break. Add a new survey question to the relevant map instead of
inlining a phrase at the call site.

### Transform pipeline

1. `limpiar(df)` — remove the Qualtrics sub-header row and blank-company rows; strip whitespace.
2. `to_num(serie)` — numeric coercion: strip thousands commas, map `""`, `"-"`, `nan`, `None`, `N/A`,
   `No Aplica`, `Desconocido` to `NaN`.
3. `_normalizar_texto(valor)` — lowercase and strip accents (NFKD) **before any categorical
   comparison**, so `"Sí"`/`"Si"` and `"Excelente"` compare reliably. `binario_si_no` (Sí→1, No→0, else
   NaN) and `relacion_excelente` (Excelente→1, Buena→0, else NaN) build on it.

### KPI construction ([calcular_kpi_empresas](../informe_base_historica.py:166))

Pools every period into one long frame, computes per-company ratios guarded against zero headcount
(`bajas / headcount.where(headcount > 0)`), aggregates each company by the **mean across periods**,
then normalizes components to 0–100 with `pct_rank`. Direction is corrected by inverting
`rotacion` and `ausentismo` via `100 - pct_rank(...)`. Finally:

```python
kpi_ambiente  = mean(salario, capacitacion_pc, prestaciones, relacion_sindical)
kpi_retencion = mean(rotacion_inv, ausentismo_inv, bono_permanencia, salario, capacitacion_pc)
kpi_general   = (kpi_ambiente + kpi_retencion) / 2
```

Missing `prestaciones` answers count as "not offered" (`.fillna(0)`); missing bono/relación stay `NaN`.

### Report flow

`main()` calls `configurar_consola()` **first** to force UTF-8 stdout/stderr, so Spanish accents render
correctly in the Windows terminal. It then runs `cargar_periodos` → `calcular_metricas` →
`calcular_kpi_empresas` and prints 9 numbered sections in priority order through the
`seccion(titulo, prioridad)` helper (sections: volumen y cobertura, calidad de datos, headcount,
salarios, contratación/rotación/ausentismo, capacitación, sindicalización/CCT, tabla resumen, KPI por
empresa).

## Conventions

* **Spanish throughout** — identifiers, docstrings, comments, and all console output. Match it.
* The **console output is the deliverable**. Format every number with `fmt_num` (thousands separators,
  `"-"` for `NaN`, optional decimals) or explicit f-string formats; never dump raw pandas objects
  without `.to_string(index=False)` and recast columns to display strings first.
* Constants and phrase dictionaries live at the top of the module in `CAPS` as a mapping, not inline.
* Both source files contain accented Spanish literals; the main script declares
  `# -*- coding: utf-8 -*-` at the top — keep source files UTF-8 encoded.
* Excel worksheet names must go through `_nombre_hoja_valido` (replace `[]:*?/\`, cap at 31 chars).
* Use `regex=False` for literal `str.replace` calls (pandas 3.x behavior).
* **Known inconsistency to keep in sync**: section labels hardcode `"7 bimestres"` / `"7 mediciones"`
  in f-strings while the real period count is derived from the workbook. If a sheet is added or
  removed, update those literals; the numbers are not computed.
