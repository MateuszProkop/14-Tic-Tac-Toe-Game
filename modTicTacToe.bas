Attribute VB_Name = "modTicTacToe"
Function ticTacToe(letter) As Boolean

'range("c2:e4") is named as gameBoard
Dim gameRng As Range
Set gameRng = Sheet7.Range("gameBoard")

If gameRng(1) = letter And gameRng(2) = letter And gameRng(3) = letter Then
    ticTacToe = True
ElseIf gameRng(4) = letter And gameRng(5) = letter And gameRng(6) = letter Then
    ticTacToe = True
ElseIf gameRng(7) = letter And gameRng(8) = letter And gameRng(9) = letter Then
    ticTacToe = True
ElseIf gameRng(1) = letter And gameRng(4) = letter And gameRng(7) = letter Then
    ticTacToe = True
ElseIf gameRng(2) = letter And gameRng(5) = letter And gameRng(8) = letter Then
    ticTacToe = True
ElseIf gameRng(3) = letter And gameRng(6) = letter And gameRng(9) = letter Then
    ticTacToe = True
ElseIf gameRng(1) = letter And gameRng(5) = letter And gameRng(9) = letter Then
    ticTacToe = True
ElseIf gameRng(3) = letter And gameRng(5) = letter And gameRng(7) = letter Then
    ticTacToe = True
End If

End Function

Sub enemyMoveTicTacToe(compLetter)

'range("c2:e4") is named as gameBoard
Dim gameRng As Range
Set gameRng = Sheet7.Range("gameBoard")

If Application.WorksheetFunction.CountA(gameRng) = 9 Then
    MsgBox "Draw"
    Exit Sub
End If

startAgain:
randomNumber = Application.WorksheetFunction.RandBetween(1, 9)

If gameRng(randomNumber) = Empty Then
    Application.EnableEvents = False
    gameRng(randomNumber) = compLetter
    Application.EnableEvents = True
    
    If ticTacToe(compLetter) Then
        MsgBox "You lose"
        Exit Sub
    End If
Else
    GoTo startAgain
End If

End Sub
