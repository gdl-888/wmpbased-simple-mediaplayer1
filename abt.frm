VERSION 5.00
Begin VB.Form abt 
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ����"
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
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "(C)���۱��� Accessable,                                     (��)����ũ�μ���Ʈ"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "�����غ�: 4348�� 9�� 11��"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "���߽���: 5�� 16��"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "���߿Ϸ�: 4352�� 5�� 15��"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "���� �̵�� ����� ���� 2.2.0"
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
