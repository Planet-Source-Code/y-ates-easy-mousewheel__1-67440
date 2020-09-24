VERSION 5.00
Begin VB.Form EasyMouseWheel 
   Caption         =   " Easy MouseWheel :)"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbDeactivate 
      Height          =   300
      Left            =   1155
      TabIndex        =   4
      Text            =   "tbDeactivate"
      Top             =   5595
      Width           =   1080
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3915
      Index           =   0
      Left            =   1575
      ScaleHeight     =   3885
      ScaleWidth      =   3120
      TabIndex        =   1
      Top             =   450
      Width           =   3150
      Begin VB.PictureBox PB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E9B287&
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   1
         Left            =   2850
         ScaleHeight     =   975
         ScaleWidth      =   210
         TabIndex        =   2
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use MouseWheel in here"
         Height          =   240
         Left            =   480
         TabIndex        =   3
         Top             =   1605
         Width           =   1845
      End
   End
   Begin VB.ComboBox cmbMouseWheel 
      Height          =   315
      ItemData        =   "EasyMouseWheel.frx":0000
      Left            =   60
      List            =   "EasyMouseWheel.frx":0002
      TabIndex        =   0
      Top             =   5595
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   210
      Left            =   75
      TabIndex        =   5
      Top             =   5340
      Width           =   6060
   End
End
Attribute VB_Name = "EasyMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''EasyMouseWheel'''''''''''''''''''''''
'Original idea belongs to Warren Goff  "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=57682&lngWId=1"
'I only improved it
'This project doesn't use Mouse Event api's. Only pure VB. So it's very light.
'Author: Hasan Yunus Ates
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private AutoClick As Boolean
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Sub cmbMouseWheel_Click()
    
    Select Case cmbMouseWheel.Text 'Up or Down
        Case "Up"
            PB(1).Top = PB(1).Top - 90
        Case "Down"
            PB(1).Top = PB(1).Top + 90
    End Select
    
    AutoClick = True
        Me.cmbMouseWheel.ListIndex = 1 'Make Idle
    AutoClick = False
    '<EXTRA: PB(1) Don't be invisible >
        If PB(1).Top < 0 Then PB(1).Top = 0
        If PB(1).Top > (PB(0).Height - PB(1).Height) Then PB(1).Top = (PB(0).Height - PB(1).Height)
    '</EXTRA: PB(1) Don't be invisible >>
End Sub

Private Sub Form_Load()
    cmbMouseWheel.AddItem "Up" 'Add references
    cmbMouseWheel.AddItem "Idle"
    cmbMouseWheel.AddItem "Down"
    
    cmbMouseWheel.ListIndex = 1 'select Idle
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GetActiveWindow = Me.hWnd And Me.ActiveControl <> tbDeactivate Then tbDeactivate.SetFocus
    '<Info>
        If GetActiveWindow <> Me.hWnd Then
            Label2.Caption = "Info : Form is inactive. So you can't Scroll."
        Else
            Label2.Caption = "Info : You can't Scroll."
        End If
    '</Info>
End Sub

Private Sub PB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GetActiveWindow = Me.hWnd And Me.ActiveControl <> cmbMouseWheel Then cmbMouseWheel.SetFocus
    '<Info>
        If GetActiveWindow <> Me.hWnd Then
            Label2.Caption = "Info : Form is inactive. So you can't Scroll."
        Else
            Label2.Caption = "Info : You can Scroll."
        End If
    '</Info>
End Sub
