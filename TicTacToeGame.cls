VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

If Target.Cells.Count > 1 Then Exit Sub
'game board range = range("c2:e4")

If Target.Value = "X" Then
    If ticTacToe("X") Then
        MsgBox "You win"
    Else
        enemyMoveTicTacToe ("O")
    End If
ElseIf Target.Value = "O" Then
    If ticTacToe("O") Then
        MsgBox "You win"
    Else
        enemyMoveTicTacToe ("X")
    End If
End If

End Sub


