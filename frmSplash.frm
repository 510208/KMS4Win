VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   BorderStyle     =   3  '雙線固定對話方塊
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Appearance      =   0  '平面
         Height          =   1065
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "版權　企鵝哥免費軟體-開源說明資訊"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   4
         Top             =   3060
         Width           =   3015
      End
      Begin VB.Label lblCompany 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "編寫　企鵝哥寫軟體"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   3390
         Width           =   3015
      End
      Begin VB.Label lblWarning 
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6330
         TabIndex        =   5
         Top             =   2700
         Width           =   525
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "作業平台"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5520
         TabIndex        =   6
         Top             =   2340
         Width           =   1335
      End
      Begin VB.Label lblProductName 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "KMS4Win"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   27.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   2385
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '靠右對齊
         Appearance      =   0  '平面
         BackColor       =   &H80000005&
         Caption         =   "尚未註冊"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "企鵝哥510208 榮譽發行"
         BeginProperty Font 
            Name            =   "微軟正黑體 Light"
            Size            =   12
            Charset         =   136
            Weight          =   290
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2355
         TabIndex        =   7
         Top             =   720
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblWarning.Caption = frmAbout.Text2.Text
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
Form1.Show
Timer1.Enabled = False
Timer1.Interval = 0
End Sub
