# Shared KiCad release workflow

This directory contains the reusable post-processing used by project-local
KiCad Jobsets. Projects retain their own `Outputs.kicad_jobset`; the shared
formatter owns common naming, assembly PDF handling, BOM generation and 3D PDF
pad-material correction.

Set `KICAD_LIB_ROOT` to the absolute path of this repository and ensure a Python
3.10+ interpreter is available on `PATH` as `python`. Execute jobs use a
shell-neutral Python launcher:

```text
python -c "import os,runpy; from pathlib import Path; runpy.run_path(str(Path(os.environ['KICAD_LIB_ROOT']) / 'release-workflow' / 'run_finalize.py'),run_name='__main__')" --project-root "${KIPRJMOD}" --kind <kind>
```

The launcher reads `KICAD_LIB_ROOT` through Python's environment API and joins
the remaining path with `pathlib`, so the same Jobset command works with Windows,
Linux and macOS path conventions without `%VAR%` or `$VAR` shell expansion.
`${KIPRJMOD}` remains a KiCad-provided project variable. Restart KiCad after
adding or changing the environment variable. On systems that expose Python only
as `python3`, provide a `python` alias/symlink or activate a virtual environment
before running KiCad/Prism.

`run_finalize.py` uses only the standard library. On Windows it starts
`py -3.14 finalize_outputs.py` with the original arguments, keeping third-party
packages out of KiCad's bundled Python 3.11. On Linux and macOS it starts
`finalize_outputs.py` with `sys.executable`, so Prism keeps using its existing
worker interpreter. The launcher's process exit status is the finalizer's exit
status.

`kind` is `prepare-fabrication`, `prepare-assembly`, `assembly`, or `fabrication`. The explicit project
root identifies the calling `.kicad_pro` and `Outputs.kicad_jobset`; KiCad
provides the temporary destination in `JOBSET_OUTPUT_WORK_PATH`.

Projects using this workflow follow these conventions:

- exactly one `.kicad_pro` in the project root;
- an `Outputs.kicad_jobset` in that root;
- project text variables `ProjectTitle`, `ProjectPCBRevision`, and
  `ProjectPCBARevision`;
- the standard `_work` intermediate filenames expected by the formatter;
- the shared `drawing-sheets/alex-generic-pcb.kicad_wks`,
  `drawing-sheets/alex-generic-sch.kicad_wks`, and
  `drawing-sheets/alex-generic-pcba.kicad_wks` worksheets.

Destination IDs and native export jobs remain project-specific. Native KiCad
exports should run first, with the corresponding formatter operation as the
final Execute Command job. Projects use the revision convention `Rx.y`: `x` is
the copper/design revision, while `y` is a BOM/documentation-only revision that
does not change the copper design. Fabrication outputs use the PCB revision and
assembly outputs use the PCBA revision.

Preparation jobs create exclusive disposable worksheet copies beneath
`${JOBSET_OUTPUT_WORK_PATH}/_work/`. Native schematic, fabrication PDF and
documentation-Gerber jobs consume those copies, and successful finalization
deletes them. Canonical worksheets remain only in this repository; projects do
not keep independent copies. The assembly preparation step uses the `sch` sheet
for the schematic and the `pcba` sheet for Top/Bottom assembly drawings. It also
resolves an empty KiCad variant to the display label `No Variant`; named variants
retain their names.

Install and test with:

```text
python -m pip install -r release-workflow/requirements.txt
python -m unittest discover -s release-workflow -p test_finalize_outputs.py -v
```

Install dependencies once in the interpreter used by the release worker, not
during every release. For a containerized Prism worker, mount this repository
(preferably read-only), set `KICAD_LIB_ROOT` to the container path, and install
the requirements into the worker's Python environment. The workflow does not
automatically install packages, download the library, or fall back to embedded
worksheet copies.

## Output and naming conventions

The formatter names final files from `ProjectTitle`, the relevant PCB or PCBA
revision, the selected variant where applicable, and one timestamp captured per
destination. Timestamps use `(YYYY-MM-DD HH-MM)` in the worker's local timezone.
Separate destinations may have different timestamps. Repeated runs within one
minute reuse names; later timestamps accumulate.

Native exports are staged under the Jobset's `_work/` directory, then renamed
and organized by the final formatter job. A typical project publishes:

```text
Releases/
  Fabrication/
    Checks/
    <title> R<PCB rev> FAB - No Variant (<timestamp>).pdf
    <title> R<PCB rev> STEP - No Variant (<timestamp>).step
    <title> R<PCB rev> CAM/
      Gerbers and native .gbrjob
      Drill/
  Assembly/
    Checks/
    <variant>/
      <title> R<PCBA rev> ASY - <variant> (<timestamp>).pdf
      <title> R<PCBA rev> BOM - <variant> (<timestamp>).xlsx
      <title> R<PCBA rev> SCH - <variant> (<timestamp>).pdf
      <title> R<PCBA rev> 3DPCB - <variant> (<timestamp>).pdf
      <title> R<PCBA rev> Pick Place - <variant> (<timestamp>).csv
```

Projects should version release directories rather than ignore them, so outputs
correspond to a known source revision. Run a release from the intended
synchronized revision without unrelated working-tree changes: release systems
may stage the checkout while synchronizing generated outputs.

## Assembly PDF, BOM, and variants

Assembly preparation creates disposable schematic and Top and, when enabled,
Bottom assembly worksheets. When both views are enabled, finalization produces
one bookmarked assembly PDF with Top followed by a mirrored Bottom view. When
only Top is enabled, the native one-page PDF is published without rewriting it.

`BOM-template.xlsx` is the shared release template. KiCad supplies grouped CSV;
the finalizer fills the template while preserving its title, merged metadata
cells, fonts, fills, column layout, print area, repeating header rows, and
fit-to-width settings. Header metadata comes from `ProjectTitle`, optional
`ProjectPartNumber`, `ProjectPCBARevision`, the destination timestamp, and the
resolved variant. Line number and quantity remain numeric; identifiers and
revisions remain text; values beginning with `=` remain literal text rather
than formulas. Missing exports, unexpected CSV headers, or invalid input fail
finalization.

The 3D-PDF step corrects the named U3D pad material to use KiCad's copper color,
then reopens and verifies the staged PDF. Unexpected pages, annotations,
streams, or material layouts fail the assembly release rather than silently
publishing a questionable document.

The helper currently processes one base assembly set per destination. Supporting
named variants requires explicit native job selectors, separate intermediate
inputs, matching worksheet labels, and corresponding formatter configuration.
Fabrication remains variant-neutral.

## Checks and failure behavior

Project Jobsets should configure ERC and DRC to include errors and warnings and
to fail on errors. Formatter jobs must leave `ignore_exit_code` disabled.
Missing dependencies, incomplete exports, invalid BOMs, unsafe paths, or
unexpected document structures then fail the destination. KiCad does not
collect a destination after a formatter error, and older release assets are not
automatically invalidated by a failed run.

KiCad versions can differ in Jobset behavior. If Jobset DRC reports that
schematic/PCB parity checking was skipped, verify parity directly in KiCad
before release.

Run a complete Jobset or one destination from the project root:

```text
kicad-cli jobset run -f Outputs.kicad_jobset <project>.kicad_pro
kicad-cli jobset run -f Outputs.kicad_jobset --output <destination-id> <project>.kicad_pro
```

Before integrating another project, verify its native layer selection, drill
and placement origins, DNP handling, variant selectors, and destination IDs.
Those choices belong to the project Jobset; this shared workflow deliberately
does not invent them.
