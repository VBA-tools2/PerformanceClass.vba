Attribute VB_Name = "modPerfDemo"

Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems
#End If

Public Sub MasterProcedure()
    Dim cPerf As clsPerf
    Set cPerf = MeasureProcedurePerformance("MasterProcedure", True)
    '-----
    'put the code of this master sub/function here
    Dim i As Long
    For i = 1 To 3
        ClientProcedure
    Next
    '(simulate execution time of some real code)
    Sleep 500
    '-----
    Application.OnTime Now, "modPerf.ReportPerformance"
End Sub

Private Sub ClientProcedure()
    Dim cPerf As clsPerf
    'replace the name of the sub/function
    Set cPerf = MeasureProcedurePerformance("ClientProcedure")
    '-----
    'put the code of this client sub/function here
    '(simulate execution time of some real code)
    Sleep 100
    '-----
End Sub
