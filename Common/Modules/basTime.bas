Attribute VB_Name = "basTime"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Function TickToDays(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 24)
  TickToDays = Tick
End Function

Public Function TickToHours(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  TickToHours = Tick
End Function

Public Function TickToMinutes(ByVal Tick As Long) As Long
  Tick = Int(Tick / 1000)
  Tick = Int(Tick / 60)
  TickToMinutes = Tick
End Function

Public Function TickToSeconds(ByVal Tick As Long) As Long
  TickToSeconds = Int(Tick / 1000)
End Function

Public Function TickDays() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 24)
  TickDays = Tick
End Function

Public Function TickHours() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  Tick = Int(Tick / 60)
  TickHours = Tick
End Function

Public Function TickMinutes() As Long
  Dim Tick As Long
  Tick = Int(GetTickCount / 1000)
  Tick = Int(Tick / 60)
  TickMinutes = Tick
End Function

Public Function TickSeconds() As Long
  TickSeconds = Int(GetTickCount / 1000)
End Function

Public Function Tick() As Long
  Tick = GetTickCount
End Function

Public Function TimeLong() As Double
  TimeLong = Int(CDbl(Now) * 100000)
End Function

Public Function TimeToLong(TimeStamp As String) As Double
  TimeToLong = Int(CDbl(CDate(TimeStamp)) * 100000)
End Function

Public Function LongToTime(LongStamp As Double) As Date
  LongToTime = CDate(LongStamp / 100000)
End Function

Public Function TimeStamp() As String
  TimeStamp = Format(Now, "\[dd\/mm\/yy:hh:mm:ss\] ")
End Function

Public Function TimeToStamp(LongStamp As Double, Optional Fmt As String) As String
  Dim lDays As Long
  Dim lHours As Long
  Select Case UCase(Trim(Fmt))
    Case "TIME"
      lDays = Val(Format(LongToTime(LongStamp), "d"))
      lHours = Val(Format(LongToTime(LongStamp), "h")) + lDays * 24
      TimeToStamp = lHours & Format(LongToTime(LongStamp), ":mm:ss")
    Case Else
      TimeToStamp = Format(LongToTime(LongStamp), "\[dd\/mm\/yy:hh:mm:ss\] ")
  End Select
End Function

Public Function TimeLeftAfter(LongStamp As Double) As String
  Dim lH As Long
  Dim lM As Long
  Dim lS As Long
  Dim lSV As Long
  lSV = DateDiff("s", LongToTime(LongStamp), Now)
  lS = lSV Mod 60
  lM = lSV / 60 Mod 60
  lH = lSV / 60 / 60
  TimeLeftAfter = lH & ":" & IIf(lM < 10, "0", "") & lM & ":" & IIf(lS < 10, "0", "") & lS
End Function

Public Function YearsLeftAfter(LongStamp As Double) As Long
  YearsLeftAfter = Val(DateDiff("yyyy", LongToTime(LongStamp), Now))
End Function

