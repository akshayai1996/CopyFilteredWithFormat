Option Explicit

'========================================================
' ENTRY POINT – USER RUNS ONLY THIS MACRO
'========================================================
Public Sub CopyFiltered_WithFormat_Final()

    Dim ws As Worksheet
    Dim srcRange As Range, visRange As Range
    Dim destCell As Range
    Dim cell As Range, area As Range
    Dim processedMerges As Object
    Dim destCol As Long

    On Error GoTo ErrHandler   ' 🔒 Global safety net

    Set ws = ActiveSheet

    ' 1️⃣ Validate source selection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select SOURCE (filtered) range first.", vbExclamation
        GoTo Cleanup
    End If

    Set srcRange = Selection

    On Error Resume Next
    Set visRange = srcRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandler

    If visRange Is Nothing Then
        MsgBox "No visible cells found.", vbInformation
        GoTo Cleanup
    End If

    ' 2️⃣ Pick Destination Column (native & scrollable)
    On Error Resume Next
    Set destCell = Application.InputBox( _
        Prompt:="1. Scroll to destination" & vbCrLf & _
                "2. Click ANY cell in the target column", _
        Title:="Select Destination Column", _
        Type:=8)
    On Error GoTo ErrHandler

    If destCell Is Nothing Then GoTo Cleanup
    destCol = destCell.Column

    ' 3️⃣ Performance optimizations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set processedMerges = CreateObject("Scripting.Dictionary")

    ' 4️⃣ Copy loop (merged-safe)
    For Each cell In visRange.Cells

        If cell.MergeCells Then
            Set area = cell.MergeArea

            If Not processedMerges.Exists(area.Address) Then
                processedMerges.Add area.Address, True
                area.Copy

                On Error Resume Next
                With ws.Cells(area.Row, destCol)
                    .PasteSpecial xlPasteAll
                    If Err.Number <> 0 Then
                        Err.Clear
                        MsgBox "Paste failed at row " & area.Row & _
                               ". Destination may be protected or merged.", _
                               vbCritical
                        GoTo Cleanup
                    End If
                    .Resize(area.Rows.Count, area.Columns.Count).Merge
                End With
                On Error GoTo ErrHandler
            End If

        Else
            cell.Copy

            On Error Resume Next
            ws.Cells(cell.Row, destCol).PasteSpecial xlPasteAll
            If Err.Number <> 0 Then
                Err.Clear
                MsgBox "Paste failed at row " & cell.Row & _
                       ". Destination may be protected.", vbCritical
                GoTo Cleanup
            End If
            On Error GoTo ErrHandler
        End If

        ' Optional heartbeat for large datasets
        'If cell.Row Mod 100 = 0 Then
        '    Application.StatusBar = "Processing row: " & cell.Row
        'End If

    Next cell

Cleanup:
    ' 🔄 Always restore Excel state
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False

    ' 🧹 Explicit cleanup (best practice)
    If Not processedMerges Is Nothing Then
        Set processedMerges = Nothing
    End If

    Exit Sub

ErrHandler:
    MsgBox "Unexpected error: " & Err.Description, vbCritical
    Resume Cleanup

End Sub
