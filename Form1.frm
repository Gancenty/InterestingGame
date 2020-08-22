VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respect"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6675
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   720
      Top             =   3600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "我不要我拒绝"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "我同意哈哈哈"
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "爱你宝贝"
      Height          =   2175
      Left            =   390
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   90
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   3960
         Picture         =   "Form1.frx":048A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   480
         Picture         =   "Form1.frx":1D50CC
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1500
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MsgBox "感谢相遇！", vbOKOnly, "来蛮来蛮"
End
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Randomize
Command2.Top = Int((Form1.Height - Command2.Height) * Rnd)
Command2.Left = Int((Form1.Width - Command2.Width) * Rnd)
End Sub


Private Sub Form_Unload(Cancel As Integer)
MsgBox "不行你不能退出！！", vbOKOnly, "呜呜呜你就答应嘛"
Cancel = True
End Sub


Private Sub Timer1_Timer()
Label1.Caption = "---------------->"
End Sub
