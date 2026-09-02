# `queries/` — generated Power Query source

`python -m sc.cli discover` decodes each workbook's `DataMashup` blob and writes
one file per query:

```
queries/<workbook-stem>__<source-id>/<QueryName>.m
queries/<workbook-stem>__<source-id>/_Section1.m   # the whole section, verbatim
```

This exists so the query stack is text under version control. Excel stores it as
an opaque base64 blob inside the workbook, which means a query edit that breaks
the pipeline is invisible in a diff and gets debugged from scratch. Once these
files are committed, the next discovery run reports every changed query in
`DISCOVERY.md` under "Power Query changes since the last discovery run".

Empty in a fresh clone — the queries come from your workbooks.

**Do not hand-edit.** These are overwritten on every run. The workbook is still
the place a query gets changed; this directory is the record of it.
