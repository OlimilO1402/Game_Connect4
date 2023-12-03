VERSION 5.00
Begin VB.Form FNewGame 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "New Game"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "Players 1 && 2"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   3375
      Begin VB.CommandButton BtnSwap 
         Caption         =   "Swap"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TbNamePlayer2 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox TbNamePlayer1 
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Colors:"
         Height          =   225
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Names:"
         Height          =   225
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.ComboBox CbHeight 
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox CbWidth 
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   225
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   525
   End
End
Attribute VB_Name = "FNewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Width   As Long
Private m_Height  As Long
Private m_Player1 As Player
Private m_Player2 As Player
Private m_Result  As VbMsgBoxResult

Private Sub Form_Load()
    fillCB CbWidth
    fillCB CbHeight
End Sub

Private Sub fillCB(aCB As ComboBox)
    Dim i As Byte: For i = 4 To 30: aCB.AddItem i: Next
End Sub

Public Function ShowDialog(aOwner As Form, aGame_out As Connect4, aPl1_out As Player, aPl2_out As Player) As VbMsgBoxResult
    'UpdateData
    If aGame_out Is Nothing Then
        m_Width = MApp.DefaultWidth
        m_Height = MApp.DefaultHeight
    Else
        m_Width = aGame_out.Width
        m_Height = aGame_out.Height
    End If
    If aPl1_out Is Nothing Then
        Set m_Player1 = MApp.DefaultPlayer1.Clone
    Else
        Set m_Player1 = aPl1_out.Clone
    End If
    If aPl2_out Is Nothing Then
        Set m_Player2 = MApp.DefaultPlayer2.Clone
    Else
        Set m_Player2 = aPl2_out.Clone
    End If
    
    UpdateView
    
    Me.Move aOwner.Left + (aOwner.Width - Me.Width) / 2, aOwner.Top + (aOwner.Height - Me.Height) / 2
    
    Me.Show vbModal, aOwner
    
    ShowDialog = m_Result
    If m_Result = vbOK Then 'UpdateData
        If aGame_out Is Nothing Then
            Set aGame_out = MNew.Connect4(m_Width, m_Height)
        Else
            aGame_out.New_ m_Width, m_Height
        End If
        If aPl1_out Is Nothing Then
            Set aPl1_out = MNew.Player(m_Player1.Name, m_Player1.Color)
        Else
            aPl1_out.Swap m_Player1
        End If
        If aPl2_out Is Nothing Then
            Set aPl2_out = MNew.Player(TbNamePlayer2.Text, Picture2.BackColor)
        Else
            aPl2_out.Swap m_Player2
        End If
    End If
End Function

Private Sub UpdateView()
    CbWidth.ListIndex = m_Width - 4
    CbHeight.ListIndex = m_Height - 4
    TbNamePlayer1.Text = m_Player1.Name
    Picture1.BackColor = m_Player1.Color
    TbNamePlayer2.Text = m_Player2.Name
    Picture2.BackColor = m_Player2.Color
End Sub

Private Sub CbWidth_Click()
    m_Width = CLng(CbWidth.Text)
End Sub

Private Sub CbHeight_Click()
    m_Height = CLng(CbHeight.Text)
End Sub

Private Sub BtnSwap_Click()
    m_Player1.Swap m_Player2
    UpdateView
End Sub

Private Sub TbNamePlayer1_LostFocus()
    m_Player1.Name = TbNamePlayer1.Text
End Sub

Private Sub TbNamePlayer2_LostFocus()
    m_Player2.Name = TbNamePlayer2.Text
End Sub

Private Sub Picture1_Click()
    GetColor Picture1
    m_Player1.Color = Picture1.BackColor
End Sub

Private Sub Picture2_Click()
    GetColor Picture2
    m_Player2.Color = Picture2.BackColor
End Sub

Private Sub GetColor(aPB As PictureBox)
    Dim ColorDlg As New ColorDialog
    ColorDlg.Color = aPB.BackColor
    If ColorDlg.ShowDialog = vbCancel Then Exit Sub
    aPB.BackColor = ColorDlg.Color
End Sub

Private Sub BtnOK_Click()
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel
    Unload Me
End Sub

