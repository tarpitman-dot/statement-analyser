# Import performance development note

Synthetic fixtures should be generated from anonymised Details-style rows only; do not use commercial statement data.

Approximate local checks on the development container are intended as regression smoke tests, not CI timing assertions:

| Fixture | Expected path | Notes |
| --- | --- | --- |
| ~5,000 rows | standard mode | Should open quickly and show progress through all stages. |
| ~50,000 rows | large-file threshold edge | Rows are processed in 3,000-row chunks and duplicate detection is deferred. |
| ~100,000 rows | Large File Mode | The worker keeps workbook parsing and row normalisation off the UI thread; browser memory remains subject to SheetJS workbook-opening limits. |

Remaining limitation: Excel files still require SheetJS to open the workbook in memory before worksheet chunks can be extracted. CSV generally has lower workbook overhead for very large statements.
