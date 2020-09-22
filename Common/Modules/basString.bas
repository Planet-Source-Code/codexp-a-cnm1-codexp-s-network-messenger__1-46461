Attribute VB_Name = "basString"
Option Explicit

Public Function GetToken(Expression As String, Delimiter As String, Index, Optional AllowEmpty As Boolean, Optional Limit As Long = -1) As String
  Dim Expr As String
  Dim SA() As String
  
  If Index < 1 Or Len(Expression) = 0 Then Exit Function
  
  Expr = Expression
  If Not AllowEmpty Then Expr = RemoveDoubleDelimiter(Expr, Delimiter)
  
  If Index = 1 And InStr(Expr, Delimiter) = 0 Then
    GetToken = Expr
    Exit Function
  End If
  
  SA = Split(Expr, Delimiter, Limit)
  
  If UBound(SA) < 0 Or UBound(SA) < Index - 1 Then Exit Function
  
  GetToken = SA(Index - 1)
  
End Function

Public Function RemoveDoubleDelimiter(Expression As String, Delimiter As String) As String
  Dim Expr As String
  Expr = Expression
  While Expr <> Replace(Expr, Delimiter & Delimiter, Delimiter)
    Expr = Replace(Expr, Delimiter & Delimiter, Delimiter)
  Wend
  RemoveDoubleDelimiter = TrimEx(Expr, Delimiter)
End Function

Public Function DelimiterCount(Expression As String, Delimiter As String, Optional AllowEmptyTokens As Boolean) As Long
  Dim lStart As Long
  Dim lCount As Long
  Dim Expr As String
  lStart = 1
  Expr = Expression
  If Not AllowEmptyTokens Then
    Expr = RemoveDoubleDelimiter(Expr, Delimiter)
  End If
  While InStr(lStart, Expr, Delimiter)
    lStart = InStr(lStart, Expr, Delimiter) + Len(Delimiter)
    lCount = lCount + 1
  Wend
  DelimiterCount = lCount
End Function

Public Function TrimEx(Expression As String, Optional What As String = " ") As String
  Dim Expr As String
  Expr = Expression
  TrimEx = RTrimEx(LTrimEx(Expr, What), What)
End Function

Public Function LTrimEx(Expression As String, Optional What As String = " ") As String
  Dim Expr As String
  Expr = Expression
  If Len(What) Then
    While What = Left(Expr, Len(What))
      Expr = Right(Expr, Len(Expr) - Len(What))
    Wend
  End If
  LTrimEx = Expr
End Function

Public Function RTrimEx(Expression As String, Optional What As String = " ") As String
  Dim Expr As String
  Expr = Expression
  If Len(What) Then
    While What = Right(Expr, Len(What))
      Expr = Left(Expr, Len(Expr) - Len(What))
    Wend
  End If
  RTrimEx = Expr
End Function

Public Function AddBackslash(ByVal Path As String) As String
  Path = Trim(Path)
  If Right(Path, 1) <> "\" Then Path = Path & "\"
  AddBackslash = Path
End Function

Public Function PrependSlash(ByVal Expression As String) As String
  If Left(Trim(Expression), 1) <> "/" Then
    PrependSlash = "/" & Expression
  Else
    PrependSlash = Expression
  End If
End Function

Public Function GetLastToken(Expression As String, Delimiter As String, Optional AllowEmpty As Boolean = True, Optional Limit As Long = -1) As String
  Dim Expr As String
  Dim SA() As String
  
  If Len(Expression) = 0 Then Exit Function
  
  Expr = Expression
  If Not AllowEmpty Then Expr = RemoveDoubleDelimiter(Expr, Delimiter)
  
  SA = Split(Expr, Delimiter, Limit)
  
  If UBound(SA) < 0 Then Exit Function
  
  GetLastToken = SA(UBound(SA))
  
End Function

Public Function DateIsValid(ByVal strDate As String) As Boolean
  On Error Resume Next
  strDate = CDate(strDate)
  If Err = 0 Then DateIsValid = True
End Function

Public Function BDateIsValid(ByVal strDate As String, Optional ByVal minAge As Long = 1, Optional ByVal maxAge As Long = 100) As Boolean
  On Error Resume Next
  strDate = CDate(strDate)
  If Err = 0 Then
    If YearsLeftAfter(TimeToLong(strDate)) <= maxAge Then
      If YearsLeftAfter(TimeToLong(strDate)) >= minAge Then
          BDateIsValid = True
      End If
    End If
  End If
End Function

