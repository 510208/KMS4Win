VERSION 5.00
Begin VB.Form Dialog 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "��ܻݱҥΤ��n��"
   ClientHeight    =   2820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4200
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Dialog.frx":0000
      Top             =   960
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "Microsoft Office"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "Microsoft Windows(7,8,101,11)"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  '����
      BackColor       =   &H80000005&
      Caption         =   "�����"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  '����
      Caption         =   "����"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  '����
      Caption         =   "�T�w"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   2175
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
If Option1(1).Value = True Then
    Shell "cmd.exe /c start " & "KMS.bat"
Else
    If Option1(2).Value = True Then
        Shell "cmd.exe /c start " & "KMSOffice.bat"
    End If
End If
End Sub

Private Sub Timer1_Timer()
If Option1(0).Value = True Then
    Text1.Text = "�п�ܲ��~"
    OKButton.Caption = "�����"
    OKButton.Enabled = False
Else
    If Option1(1).Value = True Then
        Text1.Text = "�Х��T�wKMS���A�����`"
        OKButton.Caption = "�T�{�ҥ�Windows(&W)"
        OKButton.Enabled = True
    Else
        Text1.Text = "�бN���n��Ҧ���Ʋ��JOffice�w�˸�Ƨ�(��m�b���i�H�ۦ�d��)" & vbNewLine & "�T�w�Ҧ���Ʋ��J��Ƨ����A���O�S����Ƨ��]��I"
        OKButton.Caption = "�T�{�ҥ�Office(&O)"
        OKButton.Enabled = True
    End If
End If
End Sub
