VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "옵션"
   ClientHeight    =   4065
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5340
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame3 
      Caption         =   "화면표시"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   5055
      Begin VB.CheckBox Check4 
         Caption         =   "재생 콘트롤러"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   600
         Value           =   1  '확인
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
         Caption         =   "조작 도구"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Value           =   1  '확인
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "슬라이더 대신 진행율 표시기 사용"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "실험실"
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   5055
      Begin VB.CheckBox Check1 
         Caption         =   "모든 기능 표시"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "색상"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton Option4 
         Caption         =   "보라"
         Height          =   180
         Left            =   3840
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "노랑"
         Height          =   180
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "파랑"
         Height          =   180
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "검정"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "예제 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "예제 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  '없음
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "예제 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&A/적용"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "취소"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   3615
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3615
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   " * 바꾼 설정은 굵게 표시됩니다."
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Check1.FontBold = True
End Sub

Private Sub Check2_Click()
    Check2.FontBold = True
End Sub

Private Sub Check3_Click()
    Check3.FontBold = True
End Sub

Private Sub Check4_Click()
Check4.FontBold = True

End Sub

Private Sub cmdApply_Click()
    If Check1.Value = Checked Then
        Form1.cfu.Visible = True
        Form1.sbar3.Visible = True
        Form1.hlpm.Visible = True
        Form1.al.Checked = True
    Else
        Form1.al.Checked = False
        Form1.cfu.Visible = False
        Form1.sbar3.Visible = False
        Form1.hlpm.Visible = False
    End If
    
    If Option1.Value = True Then
    Form1.blf.Visible = False
    Form1.blk.Checked = True
    Form1.blc.Checked = False
    Form1.yel.Checked = False
    Form1.ylf.Visible = False
    Form1.Label3.ForeColor = &HFFFFFF
    Form1.sec.ForeColor = &HFFFFFF
    Form1.Label1.BackColor = &H808080
    Form1.prf.Visible = False
    
    ElseIf Option2.Value = True Then
    Form1.blf.Visible = True
    Form1.blk.Checked = False
    Form1.blc.Checked = True
    Form1.yel.Checked = False
    Form1.sec.ForeColor = &H0&
    Form1.ylf.Visible = False
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HC87200
    Form1.pur.Checked = False
    
    ElseIf Option3.Value = True Then
    Form1.pur.Checked = False
    Form1.blf.Visible = False
    Form1.blk.Checked = False
    Form1.blc.Checked = False
    Form1.yel.Checked = True
    Form1.sec.ForeColor = &H0&
    Form1.ylf.Visible = True
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HC0C0&
    
    ElseIf Option4.Value = True Then
    Form1.prf.Visible = True
    Form1.ylf.Visible = False
    Form1.pur.Checked = True
    Form1.blf.Visible = False
    Form1.blk.Checked = False
    Form1.blc.Checked = False
    Form1.yel.Checked = False
    Form1.sec.ForeColor = &H0&
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HFF6699
    
    End If
    
    If Check2.Value = 1 Then
        Form1.Pbar.Visible = True
        Form1.nosl.Checked = True
    Else
        Form1.Pbar.Visible = False
        Form1.nosl.Checked = False
    End If
    
    If Check3.Value = 0 Then
    If Form1.Width = 14160 Then
        Form1.Width = 12195
        Form1.pap.Checked = False
    End If
    Else
        Form1.Width = 14160
        Form1.pap.Checked = True
    End If
    
    If Check4.Value = 0 Then
    If Form1.Height = 8535 Then
        Form1.Height = 7725
        Form1.pcm.Checked = False
    End If
    Else
        Form1.Height = 8535
        Form1.pcm.Checked = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Check1.Value = Checked Then
        Form1.cfu.Visible = True
        Form1.sbar3.Visible = True
        Form1.hlpm.Visible = True
        Form1.al.Checked = True
    Else
        Form1.al.Checked = False
        Form1.cfu.Visible = False
        Form1.sbar3.Visible = False
        Form1.hlpm.Visible = False
    End If
    
    If Option1.Value = True Then
    Form1.blf.Visible = False
    Form1.blk.Checked = True
    Form1.blc.Checked = False
    Form1.yel.Checked = False
    Form1.ylf.Visible = False
    Form1.Label3.ForeColor = &HFFFFFF
    Form1.sec.ForeColor = &HFFFFFF
    Form1.Label1.BackColor = &H808080
    Form1.prf.Visible = False
    
    ElseIf Option2.Value = True Then
    Form1.blf.Visible = True
    Form1.blk.Checked = False
    Form1.blc.Checked = True
    Form1.yel.Checked = False
    Form1.sec.ForeColor = &H0&
    Form1.ylf.Visible = False
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HC87200
    Form1.pur.Checked = False
    
    ElseIf Option3.Value = True Then
    Form1.pur.Checked = False
    Form1.blf.Visible = False
    Form1.blk.Checked = False
    Form1.blc.Checked = False
    Form1.yel.Checked = True
    Form1.sec.ForeColor = &H0&
    Form1.ylf.Visible = True
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HC0C0&
    
    ElseIf Option4.Value = True Then
    Form1.prf.Visible = True
    Form1.ylf.Visible = False
    Form1.pur.Checked = True
    Form1.blf.Visible = False
    Form1.blk.Checked = False
    Form1.blc.Checked = False
    Form1.yel.Checked = False
    Form1.sec.ForeColor = &H0&
    Form1.Label3.ForeColor = &H0&
    Form1.Label1.BackColor = &HFF6699
    
    End If
    
    If Check2.Value = 1 Then
        Form1.Pbar.Visible = True
        Form1.nosl.Checked = True
    Else
        Form1.Pbar.Visible = False
        Form1.nosl.Checked = False
    End If
    
    If Check3.Value = 0 Then
    If Form1.Width = 14160 Then
        Form1.Width = 12195
        Form1.pap.Checked = False
    End If
    Else
        Form1.Width = 14160
        Form1.pap.Checked = True
    End If
    
    If Check4.Value = 0 Then
    If Form1.Height = 8535 Then
        Form1.Height = 7725
        Form1.pcm.Checked = False
    End If
    Else
        Form1.Height = 8535
        Form1.pcm.Checked = True
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    '폼을 가운데에 놓습니다.
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    If Form1.al.Checked = True Then
        Check1.Value = Checked
    End If
    If Form1.blk.Checked = True Then
    Option1.Value = True
    ElseIf Form1.blc.Checked = True Then
    Option2.Value = True
    ElseIf Form1.yel.Checked = True Then
    Option3.Value = True
    Else
    Option4.Value = True
    End If
    If Form1.nosl.Checked = True Then
    Check2.Value = Checked
    Else
    Check2.Value = 0
    End If
    If Form1.pap.Checked = True Then
    Check3.Value = 1
    Else
    Check3.Value = 0
    End If
    If Form1.pcm.Checked = True Then
    Check4.Value = 1
    Else
    Check4.Value = 0
    End If
End Sub


Private Sub Option1_Click()
    Option1.FontBold = True
    
End Sub

Private Sub Option2_Click()
Option2.FontBold = True
End Sub

Private Sub Option3_Click()
    Option3.FontBold = True
    
End Sub

Private Sub Option4_Click()
    Option4.FontBold = True
    
End Sub
