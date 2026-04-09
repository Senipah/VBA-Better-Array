---
title: Contributing
nav_order: 4
---

# Contributing

Thanks for wanting to contribute to *VBA Better Array*!

Before submitting a pull request, please first open a [enhancement](https://github.com/Senipah/VBA-Better-Array/issues/new?assignees=&labels=enhancement&template=feature_request.md&title=) or [bug report](https://github.com/Senipah/VBA-Better-Array/issues/new?assignees=Senipah&labels=bug&template=bug_report.md&title=%5BBUG%5D) as appropriate.

If the issue raised is deemed appropriate and you would like to be assigned to deliver the solution it will then be assigned to you via the submitted ticket.

You should then fork the repository, make your changes and submit a pull request to merge the changes back into the master branch. Upon review, your changes will be merged.

#### Note
Please ensure any changes submitted to the codebase are accompanied with appropriate unit tests.

## Development Workflow

1. Make your changes in `src/` and update/add tests as needed.
2. Rebuild the workbook from source:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/createDevWorkbook.ps1
```
3. Run tests:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/runTests.ps1
```
4. If tests pass:
   - For normal contribution work, commit and push your branch.
   - For release work, run `scripts/build.ps1 <major|minor|patch>` from a clean state.

`scripts/build.ps1` creates release artifacts and then runs `git add --all`, `git commit`, `git tag`, and `git push`.
