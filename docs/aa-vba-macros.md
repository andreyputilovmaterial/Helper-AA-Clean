# AA Workbook — VBA Macro Reference

Reference of the real VBA source behind the three AA workbook buttons that
`run_aa_workbook.py` calls. These snippets were captured from the original
implementation request so the exact procedures (and the message boxes they
show) are easy to look up later.

## Mapping

| Action | Macro | VBA module |
| --- | --- | --- |
| reload | `ListAll` | `Simplified` |
| commit/save | `SaveAll` | `Simplified` |
| validate | `Button1_Click` | `Addon` |

## Action #1 — Reload: `ListAll`

Re-loads the specifications from the data file. Called with no arguments, so
`RJob` is `IsMissing` and the overwrite confirmation is shown; the watcher
clicks **OK** on it.

```vba
Sub ListAll(Optional RJob)

    If IsMissing(RJob) Then
        If MsgBox("Are you sure you want to load specifications?  Your current specifications will be over-written", vbOKCancel, "Overwrite?") = vbCancel Then Exit Sub
    Else
        'No need to prompt for R Job
    End If

    A_List_AnalysisLabels.A_List_AnalysisLabels
    If ActiveSheet.Name = "AnalysisLabels" Then 'Successful
        A_List_ShortNames.A_List_ShortNames
        A_List_AnalysisValues.A_List_AnalysisValues
        'NOTE: New Feature
        ListExclusions
    End If

End Sub
```

## Action #2 — Commit/Save: `SaveAll`

Saves the workbook, validates, and — if validation passes — writes the
specifications back and shows the success message.

- **Success** — `Save Successful` / "Your Analysis Specifications have been
  saved". The script treats success as that dialog **and** the data file (from
  `Main`!C13, captured in pre-step #0) having been modified during the save.
- **Failure** — `Validation Error` / "An error occured while processing your
  specifications.". The script captures the message-box text plus the whole
  `AnalysisValidation` sheet, then fails.

```vba
Sub SaveAll()
    'Added by Adam Spet 2017 - Save the Current Workbook
    ThisWorkbook.Save

    Dim validationMsg

    validationMsg = ValidateBeforeSave()

    If (validationMsg = "") Then
        A_Read_AnalysisLabels.A_Read_AnalysisLabels
        A_Read_ShortNames.A_Read_ShortNames
        A_Read_AnalysisValues.A_Read_AnalysisValues
        'NOTE: New Feature
        SaveExclusionsToMDD
        Sheets("Main").Select
        MsgBox "Your Analysis Specifications have been saved", vbOKOnly, "Save Successful"
    Else
        MsgBox "An error occured while processing your specifications." & vbNewLine & vbNewLine & "Message: " & validationMsg, vbOKOnly, "Validation Error"
    End If

    Sheets("Main").Select

End Sub
```

## Action #3 — Validate: `Button1_Click`

Runs the project-team validation.

- **Success** — `Validation Complete! No Errors found!`.
- **Failure** — the same `Validation Error` dialog described above.

```vba
' Call made by the Project Team Validation Button
Sub Button1_Click()
    If ValidateOnly = True Then
        ' SKIP NEW VALIDATION IF DEFAULT VALIDATION FAILS
    Else
        Addon.AnalysisLabelCharLimit
    End If
End Sub
```

## See also

- `docs/architecture.md` — how `run_aa_workbook.py` calls these macros and
  matches/dismisses their message boxes.
