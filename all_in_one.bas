

Global N As Integer
Global ans() As String
Global tmp As String
Global emptyStr As String
Global field() As Variant
Global startLoc As String
Global L() As Variant
Global L_comp() As Variant
Global score As Integer
Global baseNum As Integer
Global scoreLoc As String


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long


Sub main()


baseNum = 2
N = 4
W = 200
emptyStr = " "
startLoc = "AF117"
scoreLoc = "AD118"

ReDim field(4, 4) As Variant
ReDim L(4) As Variant
ReDim L_comp(4) As Variant


Call initField
Call showField


If checkAlive = False Then
   MsgBox ("game over")
   Exit Sub
End If


If ActiveSheet.Range(scoreLoc).Value = "" Then
    Call numPop
End If



Do
    If checkAlive = False Then
        MsgBox ("game over")
        Exit Do
    End If
    
    If GetAsyncKeyState(vbKeyReturn) <> 0 Then
        Exit Do
    End If
    
    If GetAsyncKeyState(vbKeyRight) <> 0 Then
           If moveRight = True Then
                Call showField
                Call numPop
           End If
    ElseIf GetAsyncKeyState(vbKeyLeft) <> 0 Then
           If moveLeft = True Then
                Call showField
                Call numPop
           End If
    ElseIf GetAsyncKeyState(vbKeyUp) <> 0 Then
           If moveUp = True Then
                Call showField
                Call numPop
           End If
    ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
           If moveDown = True Then
                Call showField
                Call numPop
           End If
    End If
    
    Sleep W
'    DoEvents
Loop


End Sub

Function initField()

'For i = 0 To N - 1
'    For j = 0 To N - 1
'        field(i, j) = emptyStr
'    Next j
'Next i

score = ActiveSheet.Range(scoreLoc).Value

x = ActiveSheet.Range(startLoc).Row
y = ActiveSheet.Range(startLoc).Column

For i = 0 To N - 1
    For j = 0 To N - 1
        If ActiveSheet.Cells(x + i, y + j).Value = "" Then
            field(i, j) = emptyStr
        Else
            field(i, j) = ActiveSheet.Cells(x + i, y + j).Value
        End If
    Next j
Next i
    


End Function

Function numPop()

Dim choice() As String
choice = openInd
Dim chozen As String
chozen = choice(Int(Rnd * UBound(choice)))
Dim pair() As String
Dim x, y As Integer
pair = Split(chozen, ",")
x = Int(pair(0))
y = Int(pair(1))
    If Rnd <= 0.9 Then
        field(x, y) = baseNum
    Else
        field(x, y) = baseNum * 2
    End If

Sleep 500
Call showField

End Function


Function L_clear()

For i = 0 To N - 1
L(i) = emptyStr
L_comp(i) = emptyStr
Next i

End Function

Function fill_L(i)

For x = i To N - 1

L(x) = emptyStr

Next x

End Function



Function Lcheck()

change_in_L = False

For i = 0 To N - 1
    If L(i) <> L_comp(i) Then
        change_in_L = True
        Exit For
    End If
Next i

Lcheck = change_in_L

End Function


Function moveRight()
    anyMove = False
    For x = 0 To N - 1
        i = 0
        Call L_clear
        For y = 0 To N - 1
            If field(x, N - 1 - y) <> emptyStr Then
                L(i) = field(x, N - 1 - y)
                i = i + 1
            End If
            L_comp(y) = field(x, N - 1 - y)
        Next y
        
        Call fusion(0)
        Call fill_L(i)
        If Lcheck = True Then
            anyMove = True
            For j = 0 To N - 1
                field(x, j) = L(N - 1 - j)
            Next j
        End If
    Next x
    moveRight = anyMove
End Function

Function moveLeft()
    anyMove = False
    For x = 0 To N - 1
        i = 0
        Call L_clear
        For y = 0 To N - 1
            If field(x, y) <> emptyStr Then
                L(i) = field(x, y)
                i = i + 1
            End If
            L_comp(y) = field(x, y)
        Next y
        Call fusion(0)
        Call fill_L(i)
        If Lcheck = True Then
            anyMove = True
            For j = 0 To N - 1
                field(x, j) = L(j)
            Next j
        End If
    Next x
    moveLeft = anyMove
End Function

Function moveUp()
    anyMove = False
    For y = 0 To N - 1
        i = 0
        Call L_clear
        For x = 0 To N - 1
            If field(x, y) <> emptyStr Then
                L(i) = field(x, y)
                i = i + 1
            End If
            L_comp(x) = field(x, y)
        Next x
        Call fusion(0)
        Call fill_L(i)
        If Lcheck = True Then
            anyMove = True
            For j = 0 To N - 1
                field(j, y) = L(j)
            Next j
        End If
    Next y
    moveUp = anyMove
End Function

Function moveDown()
    anyMove = False
    For y = 0 To N - 1
        i = 0
        Call L_clear
        For x = 0 To N - 1
            If field(N - 1 - x, y) <> emptyStr Then
                L(i) = field(N - 1 - x, y)
                i = i + 1
            End If
            L_comp(x) = field(N - 1 - x, y)
        Next x
        Call fusion(0)
        Call fill_L(i)
        If Lcheck = True Then
            anyMove = True
            For j = 0 To N - 1
                field(j, y) = L(N - 1 - j)
            Next j
        End If
    Next y
    moveDown = anyMove
End Function

Function fusion(x)
    For i = x To N - 2
        If L(i) <> emptyStr And L(i) = L(i + 1) Then
            L(i) = Int(L(i)) * 2
            score = score + L(i)
            ActiveSheet.Range(scoreLoc).Value = score
            Call squeezeL(i + 1)
            Call fusion(i + 1)
        End If
    Next i
End Function

Function squeezeL(i)
    For x = i To N - 2
        L(x) = L(x + 1)
    Next x
    L(N - 1) = emptyStr
End Function




Function showField()

ActiveSheet.Range(startLoc).Resize(N, N).Value = field

End Function


Function openInd()
Dim tmp As String
For i = 0 To N - 1
    For j = 0 To N - 1
        If field(i, j) = emptyStr Then
            tmp = tmp & i & "," & j & "/"
        End If
    Next j
Next i
openInd = Split(tmp, "/")

End Function



Function checkAlive()
    isAlive = True
    Dim choice() As String
    choice = openInd
    If UBound(choice) <= 0 Then
   
        For i = 0 To N - 1
            For j = 0 To N - 2
                If field(i, j) = field(i, j + 1) Then
                     checkAlive = isAlive
                     Exit Function
                End If
            Next j
        Next i
        
        For j = 0 To N - 1
            For i = 0 To N - 2
                If field(i, j) = field(i + 1, j) Then
                     checkAlive = isAlive
                     Exit Function
                End If
            Next i
        Next j
        
        isAlive = False
        
    End If
    
    checkAlive = isAlive
    
End Function








