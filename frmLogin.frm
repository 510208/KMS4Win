VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "啟用"
   ClientHeight    =   2400
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1417.999
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   5126.644
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   600
      Width           =   2205
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      Caption         =   "確定(&S)"
      Default         =   -1  'True
      Height          =   390
      Left            =   2880
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1860
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   390
      Left            =   4200
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1860
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  '平面
      Height          =   345
      IMEMode         =   3  '暫止
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label4 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "htps://www.youtube.com/channel/(這裡是ID)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label Label3 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "密碼為企鵝哥YT的ID，第一格只有5個英/數字，第二格要輸入19個英/數字"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "KMS4Win-啟用"
      BeginProperty Font 
         Name            =   "微軟正黑體"
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
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      Caption         =   "密碼(&P):"
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
    '設定全域變數為 false 來代表
    '失敗的登入
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '檢查密碼的正確性
    If txtPassword.Text + Text1.Text = "UC6orwHdQNVzwHsA6M7HYD9g" Then
        '將傳遞成功訊息至呼叫本副程式之
        '副程式的程式碼置於此處
        '設定全域變數是最簡單的方式
        LoginSucceeded = True
        Form1.loginmode.Caption = "已啟用"
        Form1.loginmode.BackColor = &HC0FFC0
        Form1.Command2.BackColor = &HC0FFC0
        Form1.loginmode.ForeColor = &H8000&
        Form1.Command1.Enabled = True
        Me.Hide
    Else
        MsgBox "密碼錯誤，請重新輸入!", , "登入"
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
