Attribute VB_Name = "modPerf"
Option Explicit

'For clsPerf Class:
Public giPerfIndex  As Long
Public gvPerfResults() As Variant
Public giDepth As Long
Public Const gbDebug As Boolean = True

Sub Demo()
    Dim cPerf As clsPerf
    ResetPerformance
    If gbDebug Then
        Set cPerf = New clsPerf
        cPerf.SetRoutine "Demo"
    End If
    Application.OnTime Now, "ReportPerformance"
End Sub

Public Sub ResetPerformance()
    giPerfIndex = 0
    ReDim gvPerfResults(1 To 3, 1 To 1)
End Sub

Public Sub ReportPerformance()
    Dim vNewPerf() As Variant
    Dim iRow As Long
    Dim iCol As Long
    ReDim vNewPerf(1 To UBound(gvPerfResults, 2) + 1, 1 To 3)
    vNewPerf(1, 1) = "Routine"
    vNewPerf(1, 2) = "Started at"
    vNewPerf(1, 3) = "Time taken"
    
    For iRow = 1 To UBound(gvPerfResults, 2)
        For iCol = 1 To 3
            vNewPerf(iRow + 1, iCol) = gvPerfResults(iCol, iRow)
        Next
    Next
    Workbooks.Add
    With ActiveSheet
        .Range("A1").Resize(UBound(vNewPerf, 1), 3).Value = vNewPerf
        .UsedRange.EntireColumn.AutoFit
    End With
    AddPivot
End Sub

Sub AddPivot()
    Dim wks As Worksheet
    Set wks = ActiveSheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                      wks.UsedRange.Address(external:=True), Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:="", TableName:="PerfReport", DefaultVersion:= _
                                      xlPivotTableVersion14
    ActiveSheet.PivotTableWizard TableDestination:=ActiveSheet.Cells(3, 1)
    ActiveSheet.Cells(3, 1).Select
    With ActiveSheet.PivotTables(1)
        With .PivotFields("Routine")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField ActiveSheet.PivotTables(1).PivotFields("Time taken"), "Total Time taken", xlAverage
        .PivotFields("Routine").AutoSort xlDescending, "Total Time taken"
        .AddDataField .PivotFields("Time taken"), "Times called", xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
    End With
End Sub
