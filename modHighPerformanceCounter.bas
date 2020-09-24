Attribute VB_Name = "modHighPerformanceCounter"
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long

Private cStartTime As Currency
Private cPerfFreq As Currency
Public Function hpc_Finish() As Long
    Dim cCurrentTime As Currency
    QueryPerformanceCounter cCurrentTime
    cCurrentTime = (cCurrentTime - cStartTime) * 1000
    hpc_Finish = cCurrentTime / cPerfFreq
End Function
Public Sub hpc_Start()
    If QueryPerformanceFrequency(cPerfFreq) = 0 Then
        Debug.Print "High-performance counter not supported!"
    End If
    QueryPerformanceCounter cStartTime
End Sub
