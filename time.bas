Attribute VB_Name = "Module1"
Option Explicit
Public Type TimeF
    H As Integer
    M As Integer
    S As Integer
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation& Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION)

Public Function GMTtoLT(myTime As TimeF) As TimeF
    Dim myTZ As TIME_ZONE_INFORMATION
    Dim lngReturn As Long
    lngReturn = GetTimeZoneInformation(myTZ)
    myTime.H = myTime.H - CInt(myTZ.Bias / 60)
    If lngReturn = 2 Then myTime.H = myTime.H - CInt(myTZ.DaylightBias / 60)
    If myTime.H < 0 Then myTime.H = 24 + myTime.H
    GMTtoLT = myTime
End Function

Public Function IsDaylight() As Boolean
    Dim myDate As Date
    Dim TD As Integer
    Dim DL As Integer
    Dim ST As Integer
    myDate = Date
    TD = Month(myDate) * 31 + Day(myDate)
    Dim myTZ As TIME_ZONE_INFORMATION
    GetTimeZoneInformation myTZ
    DL = myTZ.DaylightDate.wMonth * 31 + myTZ.DaylightDate.wDay
    ST = myTZ.StandardDate.wMonth * 31 + myTZ.StandardDate.wDay
    
    If (TD >= DL) And (TD < ST) Then IsDaylight = True Else IsDaylight = False
End Function

