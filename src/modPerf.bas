Attribute VB_Name = "modPerf"

Option Explicit

'For 'clsPerf' class:
Public giPerfIndex  As Long
Public gvPerfResults() As Variant
Public giDepth As Long

'==============================================================================
Public Const gbDebug As Boolean = False
'==============================================================================

Public Function MeasureProcedurePerformance( _
    ByVal ProcedureName As String, _
    Optional ByVal IsMaster As Boolean = False _
        ) As clsPerf
    
    If IsMaster Then ResetPerformance
    
    If gbDebug Then
        Dim cPerf As clsPerf
        Set cPerf = New clsPerf
        
        cPerf.SetRoutine ProcedureName
        Set MeasureProcedurePerformance = cPerf
    End If
End Function

Public Sub ResetPerformance()
    giPerfIndex = 0
    ReDim gvPerfResults(1 To 3, 1 To 1)
End Sub

Public Sub ReportPerformance()
    
    If gbDebug = False Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim vNewPerf() As Variant
    ReDim vNewPerf( _
            LBound(gvPerfResults, 2) To UBound(gvPerfResults, 2) + 1, _
            LBound(gvPerfResults, 1) To UBound(gvPerfResults, 1) _
    )
    vNewPerf(LBound(vNewPerf), 1) = "Routine"
    vNewPerf(LBound(vNewPerf), 2) = "Started at"
    vNewPerf(LBound(vNewPerf), 3) = "Time taken"
    
    Dim iRow As Long
    For iRow = LBound(gvPerfResults, 2) To UBound(gvPerfResults, 2)
        Dim iCol As Long
        For iCol = LBound(gvPerfResults, 1) To UBound(gvPerfResults, 1)
            vNewPerf(iRow + 1, iCol) = gvPerfResults(iCol, iRow)
        Next
    Next
    
    Dim wkb As Workbook
    Set wkb = Workbooks.Add
    
    Dim wks As Worksheet
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
    
    If wksSource Is Nothing Then
        Set wksSource = ActiveSheet
    End If
    
    Dim wkb As Workbook
    Set wkb = wksSource.Parent
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets.Add(After:=wkb.Worksheets(wkb.Worksheets.Count))
    wks.Name = "RoutinePivot"
    
    Dim pvtCache As PivotCache
    Set pvtCache = wkb.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=wksSource.UsedRange.Address(External:=True), _
            Version:=xlPivotTableVersion14 _
    )
    
    Dim pvt As PivotTable
    Set pvt = pvtCache.CreatePivotTable( _
            TableDestination:=wks.Cells(3, 1), _
            TableName:="PerfReport", _
            DefaultVersion:=xlPivotTableVersion14 _
    )
    
    With pvt
        With .PivotFields("Routine")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        Dim pvtField As PivotField
        Set pvtField = .AddDataField( _
                .PivotFields("Time taken"), _
                "Average Time taken", _
                xlAverage _
        )
        .PivotFields("Routine").AutoSort _
                xlDescending, _
                pvtField.Name
        .AddDataField _
                pvt.PivotFields("Time taken"), _
                "Times called", _
                xlCount
        .RowAxisLayout xlTabularRow
        .ColumnGrand = False
        .RowGrand = False
    End With
    
End Sub
