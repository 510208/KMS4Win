VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "����ڪ����ε{��"
   ClientHeight    =   3750
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2588.317
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  '����
      Caption         =   "�ˬdKMS���A��"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '����
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   120
      Picture         =   "frmAbout.frx":10CA
      ScaleHeight     =   630
      ScaleWidth      =   1050
      TabIndex        =   7
      Top             =   120
      Width           =   1080
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�\�@��YouTube"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '�S���ؽu
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":1682
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '�S���ؽu
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmAbout.frx":169D
      Top             =   840
      Width           =   5415
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '����
      Cancel          =   -1  'True
      Caption         =   "�T�w"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Appearance      =   0  '����
      Caption         =   "�t�θ�T(&S)..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '����u
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "KMS4Win"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Top             =   120
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
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "v1.0 (Alpha)"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���U���X�w���ʿﶵ...
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
                     
' ���U���X ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' �H Unicode nul ���������r��
Const REG_DWORD = 4                      ' 32-�줸�ƭ�

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

Private Sub Command1_Click()
Shell "cmd.exe /c start " & "https://www.kms.pub/check.html"
End Sub

Private Sub Command3_Click()
Shell "cmd.exe /c start " & "https://www.youtube.com/channel/UC6orwHdQNVzwHsA6M7HYD9g"
End Sub

Private Sub Form_Load()
    Me.Caption = "���� " & App.Title
    lblVersion.Caption = "���� " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|\�W��...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ���ձq���U�Ϩ��o�t�θ�T�{�����|...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' �ˬd�w���� 32 �줸�ɮת����O�_�s�b
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���~ - �䤣���ɮ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���~ - �䤣����U����...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "�ثe�L�k���Ѩt�θ�T", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' �j��p�ƾ�
    Dim rc As Long                                          ' �Ǧ^�N�X
    Dim hKey As Long                                        ' �}�Ҫ����U���X������N�X
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ���U���X����ƫ��A
    Dim tmpVal As String                                    ' ���U���X�Ȫ��Ȧs�Ŷ�
    Dim KeyValSize As Long                                  ' ���U���X�ܼƪ��j�p
    '------------------------------------------------------------
    ' �}�� KeyRoot {HKEY_LOCAL_MACHINE...} ���U�����U���X (RegKey)
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' �}�ҵ��U���X
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~...
    
    tmpVal = String$(1024, 0)                               ' �t�m�ܼƪŶ�
    KeyValSize = 1024                                       ' �Х��ܼƤj�p
    
    '------------------------------------------------------------
    ' �^�����U���X��...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���o/�إ߾��X��
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �B�z���~
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 �|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' ��� Null�A�q�r�ꤤ���X
    Else                                                    ' WinNT ���|�[�J�H Null ���������r��...
        tmpVal = Left(tmpVal, KeyValSize)                   ' �䤣�� Null�A���X�r��
    End If
    '------------------------------------------------------------
    ' �M�w���X�Ȫ��ഫ���A...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' �j�M��ƫ��A...
    Case REG_SZ                                             ' String ���U���X��ƫ��A
        KeyVal = tmpVal                                     ' �ƻs�r���
    Case REG_DWORD                                          ' Double Word ���U���X��ƫ��A
        For i = Len(tmpVal) To 1 Step -1                    ' �ഫ�C�@�Ӧ줸
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' �v�r�إ߭�
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' �N Double Word �ഫ�� String
    End Select
    
    GetKeyValue = True                                      ' �Ǧ^���\���T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
    Exit Function                                           ' ���}
    
GetKeyError:      ' ���~�o�ͫ�M��...
    KeyVal = ""                                             ' �]�w�Ǧ^�Ȭ��Ŧr��
    GetKeyValue = False                                     ' �Ǧ^���Ѫ��T��
    rc = RegCloseKey(hKey)                                  ' �������U���X
End Function

