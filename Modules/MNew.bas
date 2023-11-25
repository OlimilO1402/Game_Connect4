Attribute VB_Name = "MNew"
Option Explicit

Public Function Player(ByVal aName As String, ByVal aColor As Long) As Player
    Set Player = New Player: Player.New_ aName, aColor
End Function

Public Function Connect4(ByVal aWidth As Byte, ByVal aHeight As Byte) As Connect4
    Set Connect4 = New Connect4: Connect4.New_ aWidth, aHeight
End Function
