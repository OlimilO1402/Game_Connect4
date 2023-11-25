VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Connect4Chip
    None = 0
    Player1 = 1
    Player2 = 2
    Highlighted = &H10
End Enum

Private m_Field()    As Connect4Chip
Private m_IsGameOver As Boolean
Private m_Width      As Byte
Private m_Height     As Byte
Private m_CountDown  As Long
Private m_NextPlayer As Connect4Chip

Friend Sub New_(ByVal aWidth As Byte, ByVal aHeight As Byte)
    m_Width = aWidth:      m_Height = aHeight
    ReDim m_Field(1 To m_Width, 1 To m_Height)
    m_CountDown = CLng(m_Width) * CLng(m_Height)
    m_NextPlayer = Connect4Chip.Player1 'always starts with Player1
    m_IsGameOver = False
End Sub

Friend Sub NewC(other As Connect4)
    With other
        m_Width = .Width: m_Height = .Height
        ReDim m_Field(1 To m_Width, 1 To m_Height)
        .FieldCopy m_Field
        m_CountDown = .CountDown
        m_NextPlayer = .NextPlayer
        m_IsGameOver = .IsGameOver
    End With
End Sub

Friend Function Clone() As Connect4
    Set Clone = New Connect4: Clone.NewC Me
End Function

Public Property Get IsGameOver() As Boolean
    IsGameOver = m_IsGameOver
End Property

Public Property Get CountDown() As Long
    CountDown = m_CountDown
End Property

Public Property Get Width() As Byte
    Width = m_Width
End Property

Public Property Get Height() As Byte
    Height = m_Height
End Property

Public Property Get Field(ByVal iX As Byte, ByVal iY As Byte) As Connect4Chip
    If (iX < 0) Or (m_Width < iX) Then Exit Property
    If (iY < 0) Or (m_Height < iY) Then Exit Property
    Field = m_Field(iX, iY)
End Property

Public Sub FieldCopy(ByRef field_out() As Connect4Chip)
    field_out = m_Field
End Sub

Public Property Get NextPlayer() As Connect4Chip
    NextPlayer = m_NextPlayer
End Property

Public Function NextPlayerDropChip(ByVal ColumnIndex As Byte) As Boolean
    Select Case m_NextPlayer
    Case Player1: NextPlayerDropChip = PlayerDropChip(ColumnIndex, Player1)
    Case Player2: NextPlayerDropChip = PlayerDropChip(ColumnIndex, Player2)
    End Select
    TogglePlayer
End Function

Private Sub TogglePlayer()
    m_NextPlayer = IIf(m_NextPlayer = Player2, Player1, Player2)
    If m_CountDown > 0 Then
        m_CountDown = m_CountDown - 1
        m_IsGameOver = (m_CountDown = 0)
    End If
End Sub

Private Function PlayerDropChip(ByVal ColumnIndex As Byte, ByVal aChip As Connect4Chip) As Boolean
    If m_IsGameOver Then Exit Function
    'return true if successful
    If Not CheckColumnIndex(ColumnIndex) Then
        MsgBox "Out of range, select column between 0 and " & m_Width
        Exit Function
    End If
    If IsColumnFull(ColumnIndex) Then
        MsgBox "Please select another column, the column with nr " & ColumnIndex & " is full!"
        Exit Function
    End If
    DropChip ColumnIndex, aChip
    PlayerDropChip = True
End Function

Private Function CheckColumnIndex(ByVal ColumnIndex As Byte) As Boolean
    CheckColumnIndex = (1 <= ColumnIndex) And (ColumnIndex <= m_Width)
End Function

Private Function IsColumnFull(ByVal ColumnIndex As Byte) As Boolean
    IsColumnFull = m_Field(ColumnIndex, m_Height) <> None
End Function

Private Sub DropChip(ByVal ColumnIndex As Byte, ByVal aChip As Connect4Chip)
    Dim i As Long
    For i = m_Height To 1 Step -1
        If m_Field(ColumnIndex, i) <> None Then
            m_Field(ColumnIndex, i + 1) = aChip
            Exit For
        End If
        If i = 1 Then
            m_Field(ColumnIndex, 1) = aChip
            Exit For
        End If
    Next
End Sub

Public Function Check4() As Connect4Chip
    Dim aChip As Connect4Chip
    aChip = CheckH
    If aChip = None Then
        aChip = CheckV
        If aChip = None Then
            aChip = CheckD
            If aChip = None Then
                Check4 = None
                Exit Function
            End If
        End If
    End If
    Check4 = aChip
End Function

Private Function CheckH() As Connect4Chip
    Dim aChip As Connect4Chip
    Dim i As Long, j As Long
    For j = 1 To m_Height
        For i = 1 To m_Width - 3
            If m_Field(i, j) <> None Then
                aChip = m_Field(i, j)
                If m_Field(i + 1, j) = aChip Then
                    If m_Field(i + 2, j) = aChip Then
                        If m_Field(i + 3, j) = aChip Then
                            CheckH = aChip
                            m_Field(i + 0, j) = aChip Or Highlighted
                            m_Field(i + 1, j) = aChip Or Highlighted
                            m_Field(i + 2, j) = aChip Or Highlighted
                            m_Field(i + 3, j) = aChip Or Highlighted
                            m_IsGameOver = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    Next
End Function

Private Function CheckV() As Connect4Chip
    Dim aChip As Connect4Chip
    Dim i As Long, j As Long
    For i = 1 To m_Width
        For j = 1 To m_Height - 3
            If m_Field(i, j) <> None Then
                aChip = m_Field(i, j)
                If m_Field(i, j + 1) = aChip Then
                    If m_Field(i, j + 2) = aChip Then
                        If m_Field(i, j + 3) = aChip Then
                            CheckV = aChip
                            m_Field(i, j + 0) = aChip Or Highlighted
                            m_Field(i, j + 1) = aChip Or Highlighted
                            m_Field(i, j + 2) = aChip Or Highlighted
                            m_Field(i, j + 3) = aChip Or Highlighted
                            m_IsGameOver = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    Next
End Function

Private Function CheckD() As Connect4Chip
    Dim aChip As Connect4Chip
    Dim i As Long, j As Long
    For i = 1 To m_Width - 3
        For j = 1 To m_Height - 3
            If m_Field(i, j) <> None Then
                aChip = m_Field(i, j)
                If m_Field(i + 1, j + 1) = aChip Then
                    If m_Field(i + 2, j + 2) = aChip Then
                        If m_Field(i + 3, j + 3) = aChip Then
                            CheckD = aChip
                            m_Field(i + 0, j + 0) = aChip Or Highlighted
                            m_Field(i + 1, j + 1) = aChip Or Highlighted
                            m_Field(i + 2, j + 2) = aChip Or Highlighted
                            m_Field(i + 3, j + 3) = aChip Or Highlighted
                            m_IsGameOver = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    Next
    For i = 4 To m_Width
        For j = 1 To m_Height - 3
            If m_Field(i, j) <> None Then
                aChip = m_Field(i, j)
                If m_Field(i - 1, j + 1) = aChip Then
                    If m_Field(i - 2, j + 2) = aChip Then
                        If m_Field(i - 3, j + 3) = aChip Then
                            CheckD = aChip
                            m_Field(i - 0, j + 0) = aChip Or Highlighted
                            m_Field(i - 1, j + 1) = aChip Or Highlighted
                            m_Field(i - 2, j + 2) = aChip Or Highlighted
                            m_Field(i - 3, j + 3) = aChip Or Highlighted
                            m_IsGameOver = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    Next
End Function
