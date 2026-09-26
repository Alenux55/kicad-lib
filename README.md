# KiCad shared library and release tools

This repository contains the reusable KiCad resources used by the author's PCB
projects. Keeping these files here gives projects a common set of symbols,
footprints, 3D models, drawing sheets, generator inputs, and release tooling
without copying the maintenance documentation into every project repository.

## Repository contents

- `alenux.kicad_sym` — shared schematic symbols.
- `alenux.pretty/` — shared footprints.
- `3d_models/` — component models used by the shared footprints.
- `drawing-sheets/` — canonical schematic, PCB fabrication, and PCBA assembly
  worksheets.
- `generator-data/` — source data for generated library parts.
- `release-workflow/` — the shared release formatter, BOM template, dependency
  list, tests, and detailed operating documentation.

## Shared release workflow

PCB projects keep their own `Outputs.kicad_jobset`, because export jobs and
destination IDs are project-specific. Those Jobsets call the tools in
`release-workflow/` to apply the common parts of a release: worksheet
preparation, consistent filenames, assembly-PDF handling, formatted BOM
generation, and 3D-PDF correction.

Projects using the workflow publish separate design/assembly and
manufacturing/fabrication output sets. Released files are intended to be
versioned with the project so the outputs correspond to a known source
revision.

See [`release-workflow/README.md`](release-workflow/README.md) for project
requirements, runtime setup, output conventions, failure behavior, tests, and
the steps needed to integrate a project Jobset.
