VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   1  '單線固定
   Caption         =   "KMS4Win"
   ClientHeight    =   3960
   ClientLeft      =   12060
   ClientTop       =   6090
   ClientWidth     =   3960
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3960
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command2 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      Caption         =   "啟用軟體"
      Height          =   375
      Left            =   2040
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Text            =   "Form1.frx":10CA
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      Caption         =   "一鍵啟用Windows10、11(&B)"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":11D4
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "離開(&E)"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      Caption         =   "關於(&A)"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  '平面
      BackColor       =   &H00E0E0E0&
      Caption         =   "訂閱作者YouTube(&Y)"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "啟用狀態："
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label loginmode 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  '單線固定
      Caption         =   "未啟用"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.Menu right 
      Caption         =   "rightbuton"
      Visible         =   0   'False
      Begin VB.Menu about 
         Caption         =   "關於軟體(&A)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu run 
         Caption         =   "執行程式(&B)"
         Enabled         =   0   'False
      End
      Begin VB.Menu enabledKMS 
         Caption         =   "啟用軟體(&U)"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Show
End Sub


Private Sub Command2_Click()
frmLogin.Show
End Sub

Private Sub Command3_Click()
Shell "cmd.exe /c start " & "https://www.youtube.com/channel/UC6orwHdQNVzwHsA6M7HYD9g"
End Sub

Private Sub Command4_Click()
frmAbout.Show

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
Command1.ToolTipText = "請先啟用軟體"
MsgBox "請先啟用軟體", 48, "啟用"
Shell "cmd.exe /c start " & "https://www.youtube.com/channel/UC6orwHdQNVzwHsA6M7HYD9g"
Shell "cmd.exe /c start " & "https://510208web.lionfree.net/"
frmLogin.Show
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu right
End If
End Sub
