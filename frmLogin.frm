VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�ҥ�"
   ClientHeight    =   2400
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1417.999
   ScaleMode       =   0  '�ϥΪ̦ۭq
   ScaleWidth      =   5126.644
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      Height          =   345
      IMEMode         =   3  '�Ȥ�
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   2205
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      Caption         =   "�T�w(&S)"
      Default         =   -1  'True
      Height          =   390
      Left            =   2880
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "����(&C)"
      Height          =   390
      Left            =   4200
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   3
      Top             =   1860
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '����
      Height          =   345
      IMEMode         =   3  '�Ȥ�
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label4 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "htps://www.youtube.com/channel/(�o�̬OID)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label Label3 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�K�X�����Z��YT��ID�A�Ĥ@��u��5�ӭ^/�Ʀr�A�ĤG��n��J19�ӭ^/�Ʀr"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "-"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "KMS4Win-�ҥ�"
      BeginProperty Font 
         Name            =   "�L�n������"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   765
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      BackStyle       =   0  '�z��
      Caption         =   "�K�X(&P):"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   660
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '�]�w�����ܼƬ� false �ӥN��
    '���Ѫ��n�J
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '�ˬd�K�X�����T��
    If txtPassword.Text + Text1.Text = "UC6orwHdQNVzwHsA6M7HYD9g" Then
        '�N�ǻ����\�T���ܩI�s���Ƶ{����
        '�Ƶ{�����{���X�m�󦹳B
        '�]�w�����ܼƬO��²�檺�覡
        LoginSucceeded = True
        Form1.loginmode.Caption = "�w�ҥ�"
        Form1.loginmode.BackColor = &HC0FFC0
        Form1.Command2.BackColor = &HC0FFC0
        Form1.loginmode.ForeColor = &H8000&
        Form1.Command1.Enabled = True
        Me.Hide
    Else
        MsgBox "�K�X���~�A�Э��s��J!", , "�n�J"
        txtPassword.SetFocus
    End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 19 Then
    Text1.SetFocus
Else
    If Len(Text1.Text) = 0 Then
        txtPassword.SetFocus
    End If
End If
End Sub

Private Sub txtPassword_Change()
If Len(txtPassword.Text) = 5 Then
    Text1.SetFocus
End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdOK_Click
End If
End Sub
