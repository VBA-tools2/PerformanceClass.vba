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

'Which item of 'gvPerfResults' "belongs" to this instance?
Private piIndex As Long
'When did we start?
Private pdStartTime As Double

Private Sub Class_Initialize()
    'New item to keep track of, increase #
    giPerfIndex = giPerfIndex + 1
    'store which one is in this class instance
    piIndex = giPerfIndex
    'Increase the depth to create an illusion of a call stack
    giDepth = giDepth + 1
    If IsBounded(gvPerfResults) Then
        ReDim Preserve gvPerfResults(1 To 3, 1 To giPerfIndex)
    Else
        ReDim gvPerfResults(1 To 3, 1 To giPerfIndex)
    End If
    'Record when this instance was started
    pdStartTime = dMicroTimer
End Sub

Public Sub SetRoutine(sRoutineName)
    gvPerfResults(1, piIndex) = String(giDepth * 4, " ") & sRoutineName
End Sub

Private Sub Class_Terminate()
    'Called automatically when the variable pointing to this
    'class' instance goes out of scope
    
    'Outdent the call stack depth
    giDepth = giDepth - 1
    'Record start time and run-time
    gvPerfResults(2, piIndex) = pdStartTime
    gvPerfResults(3, piIndex) = dMicroTimer - pdStartTime
End Sub

Private Function IsBounded(vArray As Variant) As Boolean
    Dim iTest As Long
    On Error Resume Next
    iTest = UBound(vArray)
    IsBounded = (Err.Number = 0)
End Function