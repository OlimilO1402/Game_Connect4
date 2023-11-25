Attribute VB_Name = "MApp"
Option Explicit
Public DefaultWidth  As Long
Public DefaultHeight As Long
Public DefaultPlayer1 As Player
Public DefaultPlayer2 As Player

Sub Main()
    DefaultWidth = 7
    DefaultHeight = 6
    Set DefaultPlayer1 = MNew.Player("Player1", ColorConstants.vbYellow)
    Set DefaultPlayer2 = MNew.Player("Player2", ColorConstants.vbRed)
    FMain.Show
End Sub

Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function

