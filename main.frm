VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "심플 미디어 재생기"
   ClientHeight    =   7860
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "main.frx":030A
   ScaleHeight     =   7860
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "main.frx":168A18
      Left            =   360
      List            =   "main.frx":168A1A
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin MCI.MMControl MMC 
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   8040
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin ComctlLib.ProgressBar Pbar 
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ListBox ListBox1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "main.frx":168A1C
      Left            =   360
      List            =   "main.frx":168A1E
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "음량"
      Top             =   6480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      _Version        =   327682
      BorderStyle     =   1
      TickFrequency   =   10
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   7
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   6480
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "main.frx":168A20
      Left            =   240
      List            =   "main.frx":168A27
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      ToolTipText     =   "모드 선택"
      Top             =   960
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11520
      Top             =   240
   End
   Begin ComctlLib.Slider vol 
      Height          =   1815
      Left            =   13200
      TabIndex        =   0
      ToolTipText     =   "음량"
      Top             =   4320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3201
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      Max             =   100
      SelStart        =   50
      TickStyle       =   1
      TickFrequency   =   25
      Value           =   50
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      Caption         =   "재생 기록"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label sec 
      BackStyle       =   0  '투명
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Image pb 
      Height          =   870
      Left            =   6560
      Picture         =   "main.frx":168A31
      ToolTipText     =   "일시 중지"
      Top             =   7020
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Command4 
      Height          =   855
      Left            =   8040
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Command2 
      Height          =   855
      Left            =   6240
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Command1 
      Height          =   855
      Left            =   4320
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image CommandButton2 
      Height          =   495
      Left            =   12600
      Top             =   600
      Width           =   495
   End
   Begin VB.Image CommandButton1 
      Height          =   495
      Left            =   13080
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   6600
      ToolTipText     =   "재생"
      Top             =   7080
      Width           =   975
   End
   Begin VB.Image Image12 
      Height          =   375
      Left            =   12120
      ToolTipText     =   "음량 줄이기"
      Top             =   3000
      Width           =   735
   End
   Begin VB.Image Image11 
      Height          =   375
      Left            =   12120
      ToolTipText     =   "음량 키우기"
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image Image9 
      Height          =   735
      Left            =   7560
      ToolTipText     =   "목록의 다음 곡"
      Top             =   7080
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   735
      Left            =   12120
      ToolTipText     =   "전체 화면"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  '투명
      Caption         =   "열었던 파일"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   120
      ToolTipText     =   "최근 목록 비우기"
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   735
      Left            =   6120
      ToolTipText     =   "목록의 이전 곡"
      Top             =   7080
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   735
      Left            =   8280
      ToolTipText     =   "녹음 (미구현)"
      Top             =   7080
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   12120
      ToolTipText     =   "풀그림 조작 도구 숨기기"
      Top             =   2040
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   5160
      ToolTipText     =   "중지"
      Top             =   7200
      Width           =   735
   End
   Begin WMPLibCtl.WindowsMediaPlayer pl 
      Height          =   5040
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   10020
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17674
      _cy             =   8890
   End
   Begin VB.Image blf 
      Height          =   7905
      Left            =   0
      Picture         =   "main.frx":16B5F3
      Top             =   0
      Visible         =   0   'False
      Width           =   14040
   End
   Begin VB.Image ylf 
      Height          =   7920
      Left            =   0
      Picture         =   "main.frx":2D4ABD
      Top             =   0
      Visible         =   0   'False
      Width           =   14070
   End
   Begin VB.Image prf 
      Height          =   7935
      Left            =   0
      Picture         =   "main.frx":43FAFF
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
   End
   Begin VB.Menu filem 
      Caption         =   "&F/파일"
      Begin VB.Menu ope 
         Caption         =   "&O/불러오기"
      End
      Begin VB.Menu urlope 
         Caption         =   "&R/경로 열기"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu dosprmpt 
         Caption         =   "&D/나들이"
      End
      Begin VB.Menu exit 
         Caption         =   "&X/비상문"
      End
   End
   Begin VB.Menu vie 
      Caption         =   "&V/보기"
      Begin VB.Menu pap 
         Caption         =   "&N/풀그림 조작 도구"
         Checked         =   -1  'True
      End
      Begin VB.Menu pcm 
         Caption         =   "&C/음악 조작 도구"
         Checked         =   -1  'True
      End
      Begin VB.Menu col 
         Caption         =   "&L/색 구성표"
         Begin VB.Menu blk 
            Caption         =   "&G/검정"
            Checked         =   -1  'True
         End
         Begin VB.Menu blc 
            Caption         =   "&B/파랑"
         End
         Begin VB.Menu yel 
            Caption         =   "&Y/노랑"
         End
         Begin VB.Menu pur 
            Caption         =   "&R/보라"
         End
         Begin VB.Menu redc 
            Caption         =   "&D/빨강"
            Visible         =   0   'False
         End
         Begin VB.Menu grcol 
            Caption         =   "&S/초록"
            Visible         =   0   'False
         End
      End
   End
   Begin VB.Menu playm 
      Caption         =   "&P/재생"
      Begin VB.Menu pla 
         Caption         =   "&Y/재생"
      End
      Begin VB.Menu puas 
         Caption         =   "&U/일시 중지"
      End
      Begin VB.Menu sto 
         Caption         =   "&S/중지"
      End
   End
   Begin VB.Menu too 
      Caption         =   "&T/도구"
      Begin VB.Menu clc 
         Caption         =   "&C/최근 재생 목록 비우기"
      End
      Begin VB.Menu cfu 
         Caption         =   "&U/판 올림"
         Visible         =   0   'False
      End
      Begin VB.Menu cfuqa 
         Caption         =   "&F/업데이트가 있는지 확하기"
         Visible         =   0   'False
      End
      Begin VB.Menu sbar 
         Caption         =   "-"
      End
      Begin VB.Menu nosl 
         Caption         =   "&P/슬라이더 사용 안 함"
      End
      Begin VB.Menu ddrm 
         Caption         =   "&N/단독 실행"
      End
      Begin VB.Menu SBAR2 
         Caption         =   "-"
      End
      Begin VB.Menu exp 
         Caption         =   "&E/실험실"
         Begin VB.Menu al 
            Caption         =   "&A/모든 기능 표시"
         End
      End
      Begin VB.Menu sbar4 
         Caption         =   "-"
      End
      Begin VB.Menu optm 
         Caption         =   "&O/설정..."
      End
   End
   Begin VB.Menu helpm 
      Caption         =   "&H/도움말"
      Begin VB.Menu KickYourNeck 
         Caption         =   "&C/목차"
         Visible         =   0   'False
      End
      Begin VB.Menu qwnm 
         Caption         =   "&L/색인"
         Visible         =   0   'False
      End
      Begin VB.Menu hlpm 
         Caption         =   "&L/도움말 (미구현)"
         Visible         =   0   'False
      End
      Begin VB.Menu sbar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu abou 
         Caption         =   "&A/판 정보"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abou_Click()
    abt.Show
End Sub

Private Sub al_Click()
    If al.Checked = False Then
        al.Checked = True
        cfu.Visible = True
        sbar3.Visible = True
        hlpm.Visible = True
    Else
        al.Checked = False
        cfu.Visible = False
        sbar3.Visible = False
        hlpm.Visible = False
    End If
End Sub

Private Sub blc_Click()
    blf.Visible = True
    blk.Checked = False
    blc.Checked = True
    yel.Checked = False
    sec.ForeColor = &H0&
    ylf.Visible = False
    Label3.ForeColor = &H0&
    Label1.BackColor = &HC87200
    pur.Checked = False
End Sub

Private Sub blk_Click()
    blf.Visible = False
    blk.Checked = True
    blc.Checked = False
    yel.Checked = False
    ylf.Visible = False
    Label3.ForeColor = &HFFFFFF
    sec.ForeColor = &HFFFFFF
    Label1.BackColor = &H808080
    prf.Visible = False
End Sub

Private Sub cfu_Click()
    MsgBox "업데이트를 확인할 수 없습니다.", vbCritical, "업데이트 확인"
End Sub

Private Sub clc_Click()
    ListBox1.Clear
End Sub

Private Sub Command1_Click()
    openFrm.Show
End Sub


Private Sub Command2_Click()
abt.Show
End Sub

Private Sub Command3_Click()
    
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub Command5_Click()
    pl.Controls.currentPosition = pl.Controls.currentPosition - 10
End Sub

Private Sub Command6_Click()
    pl.Controls.currentPosition = pl.Controls.currentPosition + 10
End Sub

Private Sub ddrm_Click()
    If ddrm.Checked = True Then
    ddrm.Checked = False
    Else
    ddrm.Checked = True
    End If
End Sub

Private Sub dosprmpt_Click()
    Me.WindowState = 1
End Sub

Private Sub hlpm_Click()
    MsgBox "도움말 기능이 아직 개발되지 않았습니다.", 16
End Sub

Private Sub Image11_Click()
    vol.Value = vol.Value + 10
End Sub

Private Sub Image12_Click()
    vol.Value = vol.Value - 10
End Sub

Private Sub Image8_Click()
On Error Resume Next
    pl.fullScreen = True
End Sub

Private Sub Image9_Click()
    On Error Resume Next
    pl.Controls.Next
End Sub

Private Sub Label4_Click()

End Sub

Private Sub nosl_Click()
    If nosl.Checked = False Then
        Pbar.Visible = True
        nosl.Checked = True
    Else
        Pbar.Visible = False
        nosl.Checked = False
    End If
End Sub

Private Sub optm_Click()
    frmOptions.Show
End Sub

Private Sub pap_Click()
'12195 14160
    If Me.Width = 14160 Then
        Me.Width = 12195
        pap.Checked = False
    Else
        Me.Width = 14160
        pap.Checked = True
    End If
End Sub

Private Sub pb_Click()
    On Error Resume Next
    pl.Controls.pause
    Timer1.Enabled = False
    pb.Visible = False
End Sub

Private Sub pcm_Click()
'8535 7725
    If Me.Height = 8535 Then
        Me.Height = 7725
        pcm.Checked = False
    Else
        Me.Height = 8535
        pcm.Checked = True
    End If
End Sub

Private Sub pl_MediaChange(ByVal Item As Object)
On Error Resume Next
    Slider1.Max = pl.currentMedia.duration
    Pbar.Max = pl.currentMedia.duration
    Me.Caption = pl.currentMedia.Name + " - 심플 미디어 재생기"
    Label3.Caption = pl.currentMedia.durationString
    List1.AddItem (pl.currentMedia.Name)
End Sub

Private Sub pur_Click()
    prf.Visible = True
    ylf.Visible = False
    pur.Checked = True
    blf.Visible = False
    blk.Checked = False
    blc.Checked = False
    yel.Checked = False
    sec.ForeColor = &H0&
    Label3.ForeColor = &H0&
    Label1.BackColor = &HFF6699
End Sub

Private Sub Slider1_Scroll()
On Error Resume Next
    pl.Controls.currentPosition = Slider1.Value
    Pbar.Value = Slider1.Value
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Slider1.Value = pl.Controls.currentPosition
    'Label2.Caption = Int(Int(Slider1.Value) / 60) & ":" & Int(Int(Slider1.Value) - Int(Slider1.Value) / 60)
    sec.Caption = pl.Controls.currentPositionString
    Pbar.Value = Slider1.Value
End Sub

Private Sub CommandButton1_Click()
End
End Sub

Private Sub CommandButton2_Click()
Me.WindowState = 1
End Sub



Private Sub exit_Click()
    End
End Sub

Private Sub Form_Load()
    pl.settings.setMode "loop", True
    pl.settings.mute = False
    Combo1.ListIndex = 0
    '14160 8535
    Me.Width = 14160
    Me.Height = 8535
    Pbar.Value = 0
    'MsgBox Screen.Width & " " & Screen.Height
End Sub



Private Sub Image1_Click()
    On Error Resume Next
    pl.Controls.play
    Timer1.Enabled = True
    pb.Visible = True
End Sub

Private Sub Image3_Click()
    On Error Resume Next
    pl.Controls.Stop
    Timer1.Enabled = False
    Slider1.Value = 0
    pb.Visible = False
End Sub

Private Sub Image4_Click()
    '12195 14160
    If Me.Width = 14160 Then
        Me.Width = 12195
        pap.Checked = False
    Else
        Me.Width = 14160
        pap.Checked = True
    End If
End Sub

Private Sub Image5_Click()
MsgBox "개발중인 기능입니다.", 16, "미디어 재생기"
End Sub

Private Sub Image6_Click()
    On Error Resume Next
    pl.Controls.Previous
End Sub

Private Sub Image7_Click()
    ListBox1.Clear
End Sub

Private Sub ope_Click()
    openFrm.Show
End Sub

Private Sub pla_Click()
On Error Resume Next
    pl.Controls.play
    Timer1.Enabled = True
    pb.Visible = True
End Sub

Private Sub puas_Click()
On Error Resume Next
pl.Controls.pause
Timer1.Enabled = False
pb.Visible = False
End Sub

Private Sub sto_Click()
On Error Resume Next
pl.Controls.Stop
Timer1.Enabled = False
pb.Visible = False
End Sub

Private Sub urlope_Click()
    urlf.Show
End Sub

Private Sub vol_Change()
    pl.settings.volume = vol.Value
End Sub

Private Sub yel_Click()
    pur.Checked = False
    blf.Visible = False
    blk.Checked = False
    blc.Checked = False
    yel.Checked = True
    sec.ForeColor = &H0&
    ylf.Visible = True
    Label3.ForeColor = &H0&
    Label1.BackColor = &HC0C0&
End Sub
