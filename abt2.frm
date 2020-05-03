VERSION 5.00
Begin VB.Form abt 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "MyApp ����"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "abt2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  '�����
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  '����
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "abt2.frx":000C
      ScaleHeight     =   337.12
      ScaleMode       =   0  '�����
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Ȯ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4004
      TabIndex        =   0
      Top             =   2640
      Width           =   1587
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&S/�ý��� ����..."
      Height          =   345
      Left            =   4004
      TabIndex        =   2
      Top             =   3075
      Width           =   1587
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '���� �ܻ�
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "���� ���α׷� ����"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "����"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "(C)���۱��� Accessable,                                     (��)����ũ�μ���Ʈ"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   2625
      Width           =   3510
   End
End
Attribute VB_Name = "abt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ������Ʈ�� ���� �ɼ�...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ������Ʈ�� Ű ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode null ���� ���ڿ�
Const REG_DWORD = 4                      ' 32��Ʈ ����

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "��" & " ����"
    lblVersion.Caption = App.Major & "°�� " & App.Minor & "°������ " & App.Revision & "° �⿩"
    lblTitle.Caption = "���� �̵�� �����"
    lblDescription.Caption = "�����ϰ� ����Ʈ�� �÷��̾�" & vbNewLine & vbNewLine & "���߿Ϸ�: 4352�� 5�� 15��, ���߽���: 5�� 16��" & vbNewLine & "�����غ�: 4348�� 9�� 11��"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' �ý��� ���� ���α׷��� ��ο� �̸��� ������Ʈ������ ���� �ɴϴ�...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    '  �ý��� ���� ���α׷��� ��θ� ������Ʈ�������� ���� �ɴϴ�...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' �˷��� 32��Ʈ ���� ������ ���� ���θ� Ȯ���մϴ�.
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���� - ������ ã�� �� �����ϴ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ������Ʈ�� �׸��� ã�� �� �����ϴ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "������ �ý��� ������ ����� �� �����ϴ�.", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ���� ī����
    Dim rc As Long                                          ' ��ȯ �ڵ�
    Dim hKey As Long                                        ' ���� �ִ� ������Ʈ�� Ű ó��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ������Ʈ�� Ű�� ������ ����
    Dim tmpVal As String                                    ' ������Ʈ�� Ű ���� �ӽ÷� ����
    Dim KeyValSize As Long                                  ' ������Ʈ�� Ű ������ ũ��
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ������Ʈ�� Ű�� ���ϴ�.
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������ ó���մϴ�...
    
    tmpVal = String$(1024, 0)                             ' ������ ũ�⸦ �Ҵ��մϴ�.
    KeyValSize = 1024                                       ' ���� ũ�⸦ ǥ���մϴ�.
    
    '------------------------------------------------------------
    ' ������Ʈ�� Ű ���� �о�ɴϴ�...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Ű ���� �������� �ۼ��մϴ�.
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' ������ ó���մϴ�.
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95�� Null ���� ���ڿ��� �߰��մϴ�...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null�� ã�ҽ��ϴ�. ���ڿ����� �����մϴ�.
    Else                                                    ' WinNT�� Null ���� ���ڿ� �߰����� �ʽ��ϴ�...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null�� ã�� ���߽��ϴ�. ���ڿ������� �����մϴ�.
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������ ������ �˻��մϴ�.
    Case REG_SZ                                             ' ���ڿ� ������Ʈ�� Ű ������ ����
        KeyVal = tmpVal                                     ' ���ڿ� ���� �����մϴ�.
    Case REG_DWORD                                          ' ���� �ܾ� ������Ʈ�� Ű ������ ����
        For i = Len(tmpVal) To 1 Step -1                    ' ���� ��Ʈ�� ��ȯ�մϴ�.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' �� ���ڸ� ���ں��� �ۼ��մϴ�.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ���� �ܾ ���ڿ��� ��ȯ�մϴ�.
    End Select
    
    GetKeyValue = True                                      ' ������ ��ȯ�մϴ�.
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű�� �ݽ��ϴ�.
    Exit Function                                           ' �����մϴ�.
    
GetKeyError:      ' ������ �߻��ϸ� ����ϴ�...
    KeyVal = ""                                             ' ��ȯ���� �� ���ڿ��� �����մϴ�.
    GetKeyValue = False                                     ' ���и� ��ȯ�մϴ�.
    rc = RegCloseKey(hKey)                                  ' ������Ʈ�� Ű�� �ݽ��ϴ�.
End Function

