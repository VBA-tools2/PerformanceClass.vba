VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPerf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Which item of gvPerfResults "belongs" to this instance?
Dim mlIndex As Long

'When did we start"
Dim mdStartTime As Double

Private Sub Class_Initialize()
    'New item to keep track of, increase #
    glPerfIndex = glPerfIndex + 1
    'store which one is in this class instance
    mlIndex = glPerfIndex
    'Increase the depth to create an illusioin of a call stack
    glDepth = glDepth + 1
    If IsBounded(gvPerfResults) Then
        ReDim Preserve gvPerfResults(1 To 3, 1 To glPerfIndex)
    Else
        ReDim gvPerfResults(1 To 3, 1 To glPerfIndex)
    End If
    'Record when this instance was started
    mdStartTime = dMicroTimer
End Sub

Public Sub SetRoutine(sRoutineName)
    gvPerfResults(1, mlIndex) = String(glDepth * 4, " ") & sRoutineName
End Sub

Private Sub Class_Terminate()
    'Called automatically when the variable pointing to this
    'class' instance goes out of scope
    
    'Outdent the call stack depth
    glDepth = glDepth - 1
    'Record starttime and run-time
    gvPerfResults(2, mlIndex) = mdStartTime
    gvPerfResults(3, mlIndex) = dMicroTimer - mdStartTime
End Sub

Private Function IsBounded(vArray As Variant) As Boolean
    Dim lTest As Long
    On Error Resume Next
    lTest = UBound(vArray)
    IsBounded = (Err.Number = 0)
End Function
