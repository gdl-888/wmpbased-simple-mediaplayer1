VERSION 5.00
Begin VB.Form urlf 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "경로로 열기"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4830
   Icon            =   "urlf.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "urlf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Form1.pl.URL = Text1.Text
    Form1.ListBox1.AddItem (Text1.Text)
    Form1.Timer1.Enabled = True
    Form1.pb.Visible = True
    If Form1.ddrm.Checked = True Then
        Form1.WindowState = 1
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
