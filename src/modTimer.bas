Attribute VB_Name = "modTimer"

Option Explicit
Option Private Module

#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
#End If

Public Function dMicroTimer() As Double
'------------------------------------------------------------------------------
' Procedure : dMicroTimer
' Author    : Charles Williams www.decisionmodels.com
' Created   : 15-June 2007
' Purpose   : High resolution timer
'             Used for speed optimisation
'------------------------------------------------------------------------------
    
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    dMicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    If cyFrequency Then dMicroTimer = cyTicks2 / cyFrequency
End Function
