VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "�ҥγq��"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CheckBox Check2 
      Appearance      =   0  '����
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ڤw�q�\���Z��YT"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4200
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  '����
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ڤw�\Ū�äF�ѥH�W����"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      BackColor       =   &H00C0FFC0&
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form2.frx":10CA
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '����
      Caption         =   "�ڪ��D�F�A�}�l(&S)"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dialog.Show
End Sub

Private Sub Form_Activate()
Command1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If (Check1.Value) = 1 And (Check2.Value) = 1 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
End Sub
