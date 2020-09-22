Attribute VB_Name = "basChannel"
Option Explicit

Public Const MaxChannels = 255      ' Maximal Channels            '

Public Channels As New clsChannels  ' Channels                    '



Public Function ChanMsg(ByVal From As String, ByVal Chan As String, ByVal Message As String) As Long
  ' Return Values:                    '
  ' Return Value from ChanSendToAll() '
  
  ChanMsg = ChanSendToAll(From, Chan, "PRIVMSG #" & Chan & ":" & Message)

End Function


Public Function ChanIsValid(ByVal Chan As String) As Boolean
  Dim bInValid  As Boolean
  Dim iChars    As String
  Dim i         As Long
  
  Chan = Trim(Chan)
  iChars = "`#*\|;:"",/"
  
  If Chan = "" Then bInValid = True
  If Not bInValid Then
    If Left(Chan, 1) = "#" Then Chan = Mid(Chan, 2)
    For i = 1 To Len(iChars)
      If InStr(Chan, Mid(iChars, i, 1)) Then
        bInValid = True
        Exit For
      End If
    Next i
  End If
  
  ChanIsValid = Not bInValid
End Function


Public Sub ChanUserLogout(ByVal UserID As String)
  Static bIn As Boolean
  Dim cChan As clsChannel
  If bIn Then Exit Sub
  
  UserID = Trim(UserID)
  bIn = True
  For Each cChan In Channels
    If cChan.Users.Exist(UserID) Then
      cChan.SendToUsers ":" & UserID & " PART #" & cChan.Caption
      cChan.Users.Remove UserID
      If cChan.Users.Count < 1 Then Channels.Remove cChan.Caption
    End If
  Next cChan
  bIn = False
End Sub


Public Function ChanSendToAll(ByVal From As String, ByVal Chan As String, ByVal CmdL As String) As Long
  ' Return Values:                  '
  ' 0  - Massage was sent           '
  ' 1  - Sender is not given        '
  ' 2  - Command Line is not given  '
  ' 3  - Channel is not given       '
  ' 4  - Channelname is invalid     '
  ' 5  - Channel does not exist     '
  
  Chan = Trim(Chan)
  From = Trim(From)
  CmdL = RTrim(CmdL)
  
  If Left(Chan, 1) = "#" Then Chan = Trim(Mid(Chan, 2))
  If Not ChanIsValid(Chan) Then ChanSendToAll = 4
  If Chan = "" Then ChanSendToAll = 3
  If Trim(CmdL) = "" Then ChanSendToAll = 2
  If From = "" Then ChanSendToAll = 1
  If ChanSendToAll Then Exit Function
    
  If Channels.Exist(Chan) Then
    Channels(Chan).SendToUsers ":" & From & " " & CmdL
  Else
    ChanSendToAll = 5
    Exit Function
  End If
End Function

