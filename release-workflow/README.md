# Shared KiCad release workflow

This directory contains the reusable post-processing used by project-local
KiCad Jobsets. Projects retain their own `Outputs.kicad_jobset`; the shared
formatter owns common naming, assembly PDF handling, BOM generation and 3D PDF
pad-material correction.

Set `KICAD_LIB_ROOT` to the absolute path of this repository. Execute jobs call:

```text
py -3.14 "%KICAD_LIB_ROOT%/release-workflow/finalize_outputs.py" --project-root "${KIPRJMOD}" --kind <kind>
```

`KICAD_LIB_ROOT` is a user environment variable expanded by `cmd.exe` on
Windows. `${KIPRJMOD}` remains a KiCad-provided project variable. Restart KiCad
after adding or changing the user environment variable. A POSIX worker must use
`$KICAD_LIB_ROOT` instead of `%KICAD_LIB_ROOT%` in its project Jobset commands.

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
