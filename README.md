# GLOMBO eDNA metadata pipeline

Python tools for converting a standardised RV Investigator eDNA sampler log into
a Wilderlab sample submission workbook and a draft FAIRe workbook, then adding
Wilderlab sequencing results to the FAIRe workbook.

**Status:** prototype metadata workflow. The outputs require scientific and
metadata review before laboratory submission or publication. The scripts do not
upload data to Wilderlab or any public repository.

## Workflow

| Step | Script | Input | Output |
| --- | --- | --- | --- |
| 1 | `RVI_2_WilderLab.py` | Standard sampler CSV, job JSON, blank Wilderlab workbook | Wilderlab submission XLSX and review CSV |
| 2 | `RVI_2_FAIRe.py` | Same sampler CSV, project JSON, blank FAIRe v1.0.2 FULL workbook | Draft FAIRe XLSX and review CSV |
| 3 | `Wilderlab_2_FAIRe.py` | Draft FAIRe XLSX, Wilderlab results XLSX | Updated FAIRe XLSX, counts CSV and results review CSV |

The DAP-facing input is [schema/edna_sampler_standardised_TEMPLATE.csv](schema/edna_sampler_standardised_TEMPLATE.csv).
It contains a header only. See [docs/CSV_SCHEMA.md](docs/CSV_SCHEMA.md) and
[schema/DAP_field_dictionary.csv](schema/DAP_field_dictionary.csv) for the
field meanings, owners and checks. All files in `examples/` are synthetic.

## Requirements

- Python 3.10 or newer.
- The packages in `requirements.txt`: `openpyxl` for reading Excel and `lxml`
  for updating the workbook XML while retaining the supplied template's sheets
  and other package entries.
- Your own **blank** `FAIRe_checklist_v1.0.2_FULLtemplate.xlsx` and
  `Wilderlab_SampleSubmissionTemplate(1).xlsx`. Place these in `templates/`;
  see [templates/README.md](templates/README.md). The third-party blanks are
  intentionally not included in this public repository.

### Set up in Windows Command Prompt

Extract this repository, open **Command Prompt**, and change to its folder.
For example, change the path below to where you extracted it:

```cmd
cd /d "C:\path\to\glombo-metadata-pipeline"
python -m venv .venv
.\.venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

If your organisation manages your Python installation, you can use the
environment it supplies instead of creating `.venv`. Run all three commands
below from the repository folder. A `More?` prompt after a line ending with
`^` is normal; leave **no spaces after the caret**.

### Run the synthetic example

After putting the two blank external workbooks into `templates/`:

```cmd
python RVI_2_WilderLab.py ^
  --samples "examples\edna_sampler_standardised_MOCK.csv" ^
  --template "templates\Wilderlab_SampleSubmissionTemplate(1).xlsx" ^
  --job "examples\wilderlab_job_MOCK.json" ^
  --out "outputs\Wilderlab_submission_MOCK.xlsx" ^
  --allow-mock
```

```cmd
python RVI_2_FAIRe.py ^
  --samples "examples\edna_sampler_standardised_MOCK.csv" ^
  --template "templates\FAIRe_checklist_v1.0.2_FULLtemplate.xlsx" ^
  --project "examples\faire_project_MOCK.json" ^
  --out "outputs\FAIRe_sampler_MOCK.xlsx" ^
  --allow-mock
```

```cmd
python Wilderlab_2_FAIRe.py ^
  --faire "outputs\FAIRe_sampler_MOCK.xlsx" ^
  --results "examples\Wilderlab_results_MOCK.xlsx" ^
  --out "outputs\FAIRe_with_Wilderlab_MOCK.xlsx" ^
  --allow-mock
```

The final command must start from the **sampler draft** FAIRe workbook, not a
previously populated final workbook. Generated workbooks and review CSVs go
into the ignored `outputs/` directory.

For macOS/Linux, activate the environment with
`source .venv/bin/activate` and use `/` in paths and `\` as the line
continuation character instead of Windows `^`.

### Use real samples

1. Populate a copy of the CSV template with one row per sampling attempt,
   including failed and aborted attempts. Reconcile physical Wilderlab Kit/UIDs
   to sample event IDs and verify collection dates, positions and environment
   labels. Do not substitute the sampler's filter position for a Kit/UID.
2. Make a job JSON and project JSON by copying the synthetic JSON examples and
   replacing **all** demo values with confirmed details. Job-level contact,
   panel, sharing and billing details come from the job JSON, not the sampler
   CSV. Do not commit real JSON, CSV, results or generated workbooks.
3. Run steps 1 and 2 with the real files and omit `--allow-mock`. Review each
   generated `*.review.csv`.
4. After receiving Wilderlab results, run step 3 using the real draft FAIRe
   workbook and its matching results file. The physical UIDs must match
   `metadata.UID`; the sample event IDs must match
   `metadata.ClientSampleID`. The synthetic mock UIDs cannot be joined to a
   real Wilderlab batch.

`Wilderlab_2_FAIRe.py` rejects extra batch UIDs by default. If the results
contain samples from another project, inspect the job scope before using
`--allow-extra-results`. Only after independent verification should
`--accept-uid-only` be used to allow a ClientSampleID mismatch. These
overrides are recorded in the review output.

## What the FAIRe result step fills

- Per-sample `metadata.Assays` codes fill `sampleMetadata.assay_name`. If the
  project-level `assay_name` is blank, it receives the union of reported codes.
- Positive barcode assignments from `full` are joined to taxonomic lineage
  from `aggregated` by TaxID, ScientificName and Rank. Each distinct barcode
  is written once to both `taxaRaw` and `taxaFinal`.
- These two sheets have no sample UID/count columns. Positive counts are kept
  in the accompanying `*.counts_long.csv`, keyed by Kit/UID, sample event ID
  and sequence ID.
- The provider's processed assignments are mirrored into `taxaRaw` because
  unfiltered classifications were not supplied. Remarks mark that limitation.
  Domain, CommonName, Group and assay target are also retained in remarks.
  Taxonomic database identifiers, FASTQ provenance and experiment run metadata
  remain for later confirmation.

The initial FAIRe sampler draft will report `assay_name` as missing when the
project JSON does not supply it. The result step fills the reported assay
codes after analysis.

## Common problems

| Message | What to check |
| --- | --- |
| `ModuleNotFoundError: No module named 'lxml'` | Activate the same environment used to run the scripts, then run `python -m pip install -r requirements.txt`. |
| `collection_date YYYY-MM-DD required` | Inspect the CSV in a text editor. A date must look like `2030-01-15`; Excel may have converted it to a regional format. The error prints the offending value. |
| `FAIRe Kit/UID absent from Wilderlab results` | Check physical UID reconciliation and ensure the result workbook belongs to the same batch. The synthetic mock IDs do not match real example reports. |
| `Data Validation extension is not supported` warning from `openpyxl` | This is a reader warning, not a failed conversion. The writer edits worksheet XML and retains other workbook package entries. Verify the output in Excel. |

## Verify the package

After placing the external blank templates in `templates/`:

```cmd
python -m unittest discover -s tests -v
```

The smoke test executes all three stages in a temporary directory and checks
the mock join, assay fields, both taxon sheets and the count CSV. Without
those templates, it reports a skip.

## Sharing and stewardship

This repository includes only synthetic example records and one synthetic
results fixture. It excludes RV Investigator logs, real sample locations,
real Wilderlab results, generated workbooks and third-party blank templates.
The `.gitignore` helps prevent accidental commits of local workbooks and
outputs; review `git status` before pushing.

No software licence has been assigned here. The project owner should use a
CSIRO-approved licence and confirm third-party template terms before making
the repository public. See [docs/PUBLICATION_NOTES.md](docs/PUBLICATION_NOTES.md).
