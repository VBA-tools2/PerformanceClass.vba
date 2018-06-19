Attribute VB_Name = "modPerf"
Option Explicit

'For 'clsPerf' class:
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
    
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim vNewPerf() As Variant
    Dim iRow As Long
    Dim iCol As Long
    
    
    Application.ScreenUpdating = False
    
    ReDim vNewPerf( _
            LBound(gvPerfResults, 2) To UBound(gvPerfResults, 2) + 1, _
            LBound(gvPerfResults, 1) To UBound(gvPerfResults, 1) _
    )
    vNewPerf(LBound(vNewPerf), 1) = "Routine"
    vNewPerf(LBound(vNewPerf), 2) = "Started at"
    vNewPerf(LBound(vNewPerf), 3) = "Time taken"
    
    For iRow = LBound(gvPerfResults, 2) To UBound(gvPerfResults, 2)
        For iCol = LBound(gvPerfResults, 1) To UBound(gvPerfResults, 1)
            vNewPerf(iRow + 1, iCol) = gvPerfResults(iCol, iRow)
        Next
    Next
    
    Set wkb = Workbooks.Add
    Set wks = wkb.Worksheets(1)
    With wks
        .Name = "RoutineTable"
        .Cells(1, 1).Resize(UBound(vNewPerf, 1), UBound(vNewPerf, 2)).Value = vNewPerf
        .UsedRange.EntireColumn.AutoFit
        
        .ListObjects.Add(xlSrcRange, .UsedRange, , xlYes).Name = _
                "RoutineTable"
    End With
    
    Call AddPivot(wks)
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub AddPivot( _
    Optional ByVal wksSource As Worksheet = Nothing _
)
    
    Dim wkb As Workbook
    Dim wks As Worksheet
    
    
    If wksSource Is Nothing Then
        Set wksSource = ActiveSheet
    End If
    
    Set wkb = wksSource.Parent
    Set wks = wkb.Worksheets.Add
    wks.Name = "RoutinePivot"
    
    wkb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                                      wks.UsedRange.Address(external:=True), Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:="", TableName:="PerfReport", DefaultVersion:= _
                                      xlPivotTableVersion14
    wks.PivotTableWizard TableDestination:=wks.Cells(3, 1)
    wks.Cells(3, 1).Select
    With wks.PivotTables(1)
        With .PivotFields("Routine")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField wks.PivotTables(1).PivotFields("Time taken"), "Total Time taken", xlAverage
        .PivotFields("Routine").AutoSort xlDescending, "Total Time taken"
        .AddDataField .PivotFields("Time taken"), "Times called", xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
    End With
    
End Sub
