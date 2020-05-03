VERSION 5.00
Begin VB.Form abt 
   BorderStyle     =   1  '단일 고정
   Caption         =   "버전 정보"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "abt.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "(C)저작권자 Accessable,                                     (주)마이크로소프트"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "개발준비: 4348년 9월 11일"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "개발시작: 5월 16일"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "개발완료: 4352년 5월 15일"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "심플 미디어 재생기 버전 2.2.0"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "abt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub
