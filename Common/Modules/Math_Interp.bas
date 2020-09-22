Attribute VB_Name = "Math_Interp"
' vb@rchiv - Das große Visual-Basic Archiv
' Tools & Components - Entwicklerkomponenten für VB-32 Bit
'
' Copyright ©2002-2003 vb@rchiv
' Workshop-Autor: Dietmar G. Bayer (vb@rchiv)
'
' Der Programmcode darf für eigene Zwecke verwendet werden.
' Es ist nicht erlaubt Inhalte des Projektes ohne unserer
' Zustimmung zum Download anzubieten.
'
' Die Beispielskripte sind Computerprogramme, die gemäß
' des §2 Abs. 1 Nr. 69 aff. UrhG den urheberrechtlichen
' Schutz geniessen und dürfen nicht für eigene ausgegeben
' werden.
'
' Dieter Otter
' Software-Entwicklung & Vertrieb
' info@vbarchiv.de
' http://www.vbarchiv.de
' http://www.visualbasic-archiv.de
'
' info@tools4vb.de
' http://www.tools4vb.de
'======================================================
Option Explicit
Public Function Mathe_Interpreter(ByVal Formel As String) As String
'****************************************************************************************
'*Dies ist die Hauptfunktion.
'*Von hier wird die Mathematik gesteuert

'****************************************************************************************
  Formel = Norm(Formel)
  
  If InStr(1, Formel, "[") > 0 Then Formel = CSTA(Formel)         'Berechne und entferne cos[],sin[],tan[]und atn[]

  If InStr(1, Formel, "(") > 0 Then Formel = Klammer(Formel)      'Berechne den Klammerinhalt und entferne die Klammer

  If InStr(1, Formel, "^") > 0 Then Formel = MDHPM(Formel, "^")   'Berechne und entferne ^
  If InStr(1, Formel, "*") > 0 Then Formel = MDHPM(Formel, "*")   'Berechne und entferne *
  If InStr(1, Formel, "/") > 0 Then Formel = MDHPM(Formel, "/")   'Berechne und entferne /
  If Formel = "unmöglich" Then                                    'Bei Division Durch 0
      Err.Raise 11                                                'Fehler erzeugen
      Exit Function
  End If
  If InStr(1, Formel, "+") > 0 Then Formel = MDHPM(Formel, "+")   'Berechne und entferne +
  If InStr(2, Formel, "-") > 0 Then Formel = MDHPM(Formel, "-")   'Berechne und entferne -
  Formel = Norm(Formel)
  Mathe_Interpreter = Formel                                      'Und fertig!
End Function




' Doppeloperatoren (--, +-, -+, ++) werden entfernt
Function Entferne_PlusMinus(ByVal Formel As String) As String
  Dim Platz As Integer
  Dim NewFormel As String
  
  Entferne_PlusMinus = Formel
  
  If InStr(Formel, "++") Then Formel = Replace(Formel, "++", "+")
  If InStr(Formel, "+-") Then Formel = Replace(Formel, "+-", "-")
  If InStr(Formel, "-+") Then Formel = Replace(Formel, "-+", "-")
  If InStr(Formel, "--") Then Formel = Replace(Formel, "--", "+")
  
  Entferne_PlusMinus = Formel
  
End Function


' Formel normieren
Function Norm(Text As String) As String
  
  Text = LCase(Text) ' alles klein
  Text = Replace(Text, ",", ".") 'alle Kommas zu Punkten machen
  Text = Replace(Text, " ", "") 'Keine Spaces
  Norm = Entferne_PlusMinus(Text)
End Function
Function RP(Text As String) As String
  RP = Replace(Text, ",", ".") 'alle Kommas zu Punkten machen
End Function

'CSTA = Cos Sin Tan Atn

Public Function CSTA(ByVal Formel As String) As String
Dim Platz As Integer
Dim Y As Integer
Dim X As Integer
Dim VorneStr As String
Dim MittelStr As String
Dim NeuFormel As String
Dim XType As String
Formel = Norm(Formel)
CSTA = Formel
Platz = InStr(1, Formel, "]")
If Platz > 0 Then
    Y = 0
    X = Platz
    Do 'suche das innere der Klammer
        Y = Y + 1
        X = X - 1
        If X = 0 Then Exit Do '
        If Mid(Formel, X, 1) <> "[" Then
            MittelStr = Mid(Formel, X, Y)
        End If
    Loop Until Mid(Formel, X, 1) = "["
    VorneStr = "" '
    If X > 0 Then VorneStr = Mid(Formel, 1, X - 4)
    If X > 0 Then XType = Mid(Formel, X - 3, 3)
    Select Case XType
    Case "cos"
        NeuFormel = VorneStr & CStr(Cos(Format(MittelStr, "0.0000"))) & Mid(Formel, Platz + 1)
    Case "sin"
        NeuFormel = VorneStr & CStr(Sin(Format(MittelStr, "0.0000"))) & Mid(Formel, Platz + 1)
    Case "tan"
        NeuFormel = VorneStr & CStr(Tan(Format(MittelStr, "0.0000"))) & Mid(Formel, Platz + 1)
    Case "atn"
        NeuFormel = VorneStr & CStr(Atn(Format(MittelStr, "0.0000"))) & Mid(Formel, Platz + 1)
    End Select
    If InStr(1, NeuFormel, "]") > 0 Then
        NeuFormel = CSTA(NeuFormel) 'und ab in die Rekursion
    End If
    CSTA = RP(Format(Replace(NeuFormel, ".", ","), "0.0000"))
End If
End Function
Function Klammer(ByVal Formel As String) As String
  Dim Platz As Integer
  Dim Y As Integer
  Dim X As Integer
  Dim VorneStr As String
  Dim MittelStr As String  ' der ist neu, da das zu berechnende Teil zwischen 2 Klammern steht
  Dim NeuFormel As String
  
  Formel = Norm(Formel) 'Die Doppelten wieder weg
  
  Klammer = Formel ' aus sicherheit, wenn es keine klammer gibt
  Platz = InStr(1, Formel, ")")
  If Platz > 0 Then
    Y = 0
    X = Platz
    Do 'das Innere der Klammer suchen
        Y = Y + 1
        X = X - 1
        If X = 0 Then Exit Do
        If Mid(Formel, X, 1) <> "(" Then
            MittelStr = Mid(Formel, X, Y)
        End If
    Loop Until Mid(Formel, X, 1) = "("
    VorneStr = ""
    If X > 0 Then VorneStr = Mid(Formel, 1, X - 1)
    ' die Hauptfunktion rekursiv aufrufen damit der innere Teil der Funktion interpretiert wird
    NeuFormel = VorneStr & Mathe_Interpreter(MittelStr) & Mid(Formel, Platz + 1)
    
    If InStr(1, Formel, ")") > 0 Then
        NeuFormel = Klammer(NeuFormel) 'Wieder in die Rekursion für  mehr Klammern
    End If
    Klammer = RP(Format(Replace(NeuFormel, ".", ","), "0,0000"))
  End If
End Function


Function MDHPM(ByVal Formel As String, Operator As String) As String
  ' Stelle an der das Zeichen gefunden wurde : hier ein * oder /
  Dim Platz As Integer
  ' Variable zu Vorwärts/Rückwärtslaufen im String
  Dim X As Integer
  ' Zähler
  Dim Y As Integer
  ' Welche Zahl steht vor dem Zeichen ?
  Dim vorne As String
  ' Welche Zahl steht nach dem Zeichen ?
  Dim hinten As String
  ' Neue Formel nach einer Teil-Lösung
  Dim NewFormel As String
  ' Speichern des alten Strings
  Dim AltStr As String
  ' Die jeweils anderen Operatoren
  Dim AlleOperatoren As String
  Dim gefunden As Boolean
  Dim i As Integer
  Dim NotFund As Boolean
  Dim AltStr_Hinten As String
  
  AlleOperatoren = "+-*/^"
  
  ' Erst mal sicherstellen, dass keine Doppel-Operatoren da sind.
  Formel = Norm(Formel)
  If InStr(Formel, "e") Then Exit Function  ' wegen 2.0000056e-3
  
  ' Für den Fall, dass keine Operation angefordert ist
  MDHPM = Formel

  If FindOperator(Mid(Formel, 2), AlleOperatoren) = 0 Then
    MDHPM = Formel
    Exit Function
  End If
  
  ' wo seht der Operator
  Platz = InStr(1, Formel, Operator)
  If Platz > 0 Then
    Y = 0
    X = Platz
    ' wandere zurück, um die Zahl vor dem Operator zu finden
    Do
      Y = Y + 1
      X = X - 1
      If X = 0 Then
        ' für den Fall, dass wir am Anfang gelandet sind
        vorne = Mid(Formel, X + 1, Y - 1)
        Exit Do
      End If
      
      If FindOperator(Mid(Formel, X, 1), AlleOperatoren) > 0 Then
        If Y = 1 Then
          If InStr(Mid(Formel, X, 1), "+") Then
            ' nötig wg. -22*+12
            vorne = Mid(Formel, X, Y)
            Exit Do
          End If
          If InStr(Mid(Formel, X, 1), "-") Then
            ' nötig wg. -22*+12
            vorne = Mid(Formel, X, Y)
            Exit Do
          End If
        End If
        vorne = Mid(Formel, X + 1, Y)
        Exit Do
      End If
    Loop
    AltStr = Left(Formel, X)
    
    ' Nun das Ganze vorwärts:
    Y = 0
    X = Platz + 1
    Do
      Y = Y + 1
      ' X = X + 1
      
      If Y > Len(Formel) Then
        hinten = Mid(Formel, X)
        AltStr_Hinten = ""
        Exit Do
      End If
              
      If FindOperator(Mid(Formel, X, 1), AlleOperatoren) > 0 Then
        ' (InStr(Mid(Formel, X, 1), AlleOperatoren) > 0 Then
        hinten = Mid(Formel, X, Y)
        AltStr_Hinten = Mid(Formel, X + Y)
        Exit Do
      End If
    Loop

    ' Das wiederum so lange machen, bis wir die Zahl hinter
    ' dem Operator gefunden haben, oder das Ende erreicht ist.
    ' So wird jetzt unsere neue Formel berechnet ausschauen:

    If AltStr = "-" Then
      AltStr = ""
      vorne = "-" + vorne
    End If
    If AltStr = "+" Then
      AltStr = ""
      vorne = "+" + vorne
    End If
    
    Select Case Operator
      Case "*"
        NewFormel = AltStr & CStr((Val(vorne) * Val(hinten))) & AltStr_Hinten
        If InStr(1, NewFormel, "*") > 0 Then
          ' Die Rekursion aufrufen, da noch weitere Multiplikation zu machen sind
          NewFormel = MDHPM(NewFormel, Operator)
        End If
    
      Case "/"
        NewFormel = AltStr & CStr((Val(vorne) / Val(hinten))) & AltStr_Hinten
        If Err.Number = 11 Then
          ' Division durch 0 ist nicht möglich
          MDHPM = "unmöglich"
          Exit Function
        End If

        If InStr(1, NewFormel, "/") > 0 Then
          ' Die Rekursion aufrufen, da noch weitere Divisionen zu machen sind
          NewFormel = MDHPM(NewFormel, Operator)
        End If
        
      Case "^"
        NewFormel = AltStr & CStr((Val(vorne) ^ Val(hinten))) & AltStr_Hinten
        If InStr(1, NewFormel, "^") > 0 Then
          ' Die Rekursion aufrufen, da noch weitere Potenzierungen zu machen sind
          NewFormel = MDHPM(NewFormel, Operator)
        End If
        
      Case "+"
        NewFormel = AltStr & CStr((Val(vorne) + Val(hinten))) & AltStr_Hinten
        If InStr(1, NewFormel, "+") > 0 Then
          ' Die Rekursion aufrufen, da noch weitere Additionen zu machen sind
          NewFormel = MDHPM(NewFormel, Operator)
        End If
        
      Case "-"
        NewFormel = AltStr & CStr((Val(vorne) - Val(hinten))) & AltStr_Hinten
        If InStr(1, NewFormel, "-") > 0 Then
          ' Die Rekursion aufrufen, da noch weitere Subtraktionen zu machen sind
          NewFormel = MDHPM(NewFormel, Operator)
        End If
      
    End Select
    ' und fertig
    MDHPM = CStr(NewFormel)
    MDHPM = RP(Format(Replace(NewFormel, ".", ","), "0.0000"))
  End If
End Function

Function FindOperator(Text As String, Operator As String) As Long
  Dim i As Long
  
  FindOperator = 0
  For i = 1 To Len(Operator)
    If InStr(Text, Mid(Operator, i, 1)) > 0 Then
      FindOperator = InStr(Text, Mid(Operator, i, 1))
      Exit Function
    End If
  Next

End Function

