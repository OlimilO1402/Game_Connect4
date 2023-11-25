VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name  As String
Public Color As Long

Friend Sub New_(ByVal aName As String, ByVal aColor As Long)
    Name = aName: Color = aColor
End Sub

Friend Sub NewC(other As Player)
    Name = other.Name
    Color = other.Color
End Sub

Friend Function Clone() As Player
    Set Clone = New Player: Clone.NewC Me
End Function

Public Sub Swap(other As Player)
    Dim tmp As Player: Set tmp = other.Clone
    other.NewC Me: Me.NewC tmp
End Sub

