# External workbook templates

The scripts are tested with these **unpopulated** input workbooks:

1. `Wilderlab_SampleSubmissionTemplate(1).xlsx`: Wilderlab sample submission
   workbook supplied to the project, with `Job metadata`, `Sample metadata`,
   `Extra metadata` and `Lists` sheets.
2. `FAIRe_checklist_v1.0.2_FULLtemplate.xlsx`: FAIRe checklist v1.0.2 FULL
   workbook with `projectMetadata`, `sampleMetadata`, `taxaRaw` and `taxaFinal`
   sheets, among others.

Obtain the applicable blanks from the template maintainers, check their
distribution terms, and place them in this folder under those names. The
scripts reject populated sample and taxon sheets or changed layouts rather
than overwriting existing records. These third-party files are excluded from the GitHub package
and ignored by Git.
