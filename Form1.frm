VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Low-level Streaming Audio Sample Project"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "stop"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "play"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
   End
   Begin VB.CommandButton Command1 
      Caption         =   "file"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fMovingSlider As Boolean

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Label1.Caption = CommonDialog1.filename
OpenFile CommonDialog1.filename
Slider1.Value = 0
End Sub

Private Sub Command2_Click()
Play
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
StopPlay
End Sub


Private Sub Form_Load()
Slider1.Min = 0
Slider1.Max = 100
Initialize Me.hwnd
fMovingSlider = False
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
fMovingSlider = True
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
FileSeek (Slider1.Value / 100) * Length
fMovingSlider = False
End Sub

Private Sub Timer1_Timer()
If (fMovingSlider) Then
    Exit Sub
End If
If (Playing() = False) Then
    Timer1.Enabled = False
End If
Slider1.Value = (Position() / Length()) * 100
End Sub
