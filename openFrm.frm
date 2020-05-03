VERSION 5.00
Begin VB.Form openFrm 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "열기"
   ClientHeight    =   3660
   ClientLeft      =   9825
   ClientTop       =   7125
   ClientWidth     =   6540
   Icon            =   "openFrm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   255
      Left            =   6000
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   12
      Top             =   240
      Width           =   5175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "사용(&U)"
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   3240
      Value           =   1  '확인
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "openFrm.frx":000C
      Left            =   120
      List            =   "openFrm.frx":0013
      Style           =   2  '드롭다운 목록
      TabIndex        =   9
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "취소"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "확인"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   1980
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2070
      Left            =   120
      Pattern         =   "*.wpl;*.mp3;*.mp4;*.mpg;*.mpe;*.mpeg;*.mp2;*.mp1;*.wma;*.wmv;*.mov;*.avi;*.wav;*.mid;*.midi;*.rmi;*.wtv;*.dvrms;*.dvr-ms;*.cda"
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "경로:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "파일 형식:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "파일 목록:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "디렉토리 구조:"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "드라이브:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "openFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
    Form1.pl.URL = File1.Path + "\" + File1.FileName
    Form1.ListBox1.AddItem (File1.FileName)
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

Private Sub Command3_Click()
    On Error Resume Next
    Dir1.Path = Text1.Text
    File1.Path = Text1.Text
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Text1.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    Text1.Text = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
    Combo1.ListIndex = 0
    Dir1.Path = "C:\WINDOWS\MEDIA"
    File1.Path = "C:\WINDOWS\MEDIA"
End Sub
