Attribute VB_Name = "CohortRefresh"
Option Explicit

' Lumora Goods cohort revenue model: refresh and housekeeping macros.
' Import once (Alt+F11 > File > Import File), then save the workbook as .xlsm.
'
' Auto_Open runs every time the workbook opens. It recalculates the model, so the
' Auto cut-off picks up the latest complete month in Historical Data, checks the
' history table, and stamps the refresh time on the Assumptions sheet.

Public Sub Auto_Open()
    RefreshModel
End Sub

Public Sub RefreshModel()
    Dim issues As String
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.CalculateFull
    issues = CheckHistory()
    With ThisWorkbook.Worksheets("Assumptions")
        .Range("LastRefresh").Value = Now
        If Len(issues) = 0 Then
            .Range("HistoryStatus").Value = "OK: no duplicates, month mismatches or blanks"
        Else
            .Range("HistoryStatus").Value = issues
        End If
    End With
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "Refresh stopped: " & Err.Description, vbExclamation, "Cohort model"
End Sub

' Returns an empty string when Historical Data is clean, otherwise a short summary of problems.
Private Function CheckHistory() As String
    Dim lo As ListObject, i As Long, key As String
    Dim seen As Object, dupes As Long, mismatches As Long, blanks As Long
    Dim cCohort As Long, cMonths As Long, cRevenue As Long, cCheck As Long

    Set lo = ThisWorkbook.Worksheets("Historical Data").ListObjects("tblHistory")
    If lo.DataBodyRange Is Nothing Then
        CheckHistory = "Historical Data is empty"
        Exit Function
    End If
    cCohort = lo.ListColumns("Cohort").Index
    cMonths = lo.ListColumns("Months Since First").Index
    cRevenue = lo.ListColumns("Revenue").Index
    cCheck = lo.ListColumns("Check").Index
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To lo.ListRows.Count
        With lo.DataBodyRange
            If IsEmpty(.Cells(i, cRevenue).Value) Then blanks = blanks + 1
            key = CStr(CLng(.Cells(i, cCohort).Value)) & "|" & CStr(.Cells(i, cMonths).Value)
            If seen.Exists(key) Then
                dupes = dupes + 1
            Else
                seen.Add key, True
            End If
            If CStr(.Cells(i, cCheck).Value) <> "OK" Then mismatches = mismatches + 1
        End With
    Next i

    If dupes + mismatches + blanks > 0 Then
        CheckHistory = "Check Historical Data: " & dupes & " duplicate rows, " & _
                       mismatches & " month mismatches, " & blanks & " blank revenue cells"
    End If
End Function

' Adds the month after the last cohort as a new row. Enter its Ad Spend and ROAS afterwards;
' every other column fills in from the table's formulas.
Public Sub AddNextCohort()
    Dim lo As ListObject, lastDate As Date, newRow As ListRow
    Set lo = ThisWorkbook.Worksheets("Cohort Model").ListObjects("tblModel")
    lastDate = lo.ListColumns("Cohort").DataBodyRange.Cells(lo.ListRows.Count, 1).Value
    Set newRow = lo.ListRows.Add
    newRow.Range.Cells(1, lo.ListColumns("Cohort").Index).Value = DateSerial(Year(lastDate), Month(lastDate) + 1, 1)
    Application.Goto newRow.Range.Cells(1, lo.ListColumns("Ad Spend").Index)
End Sub
