VERSION 5.00
Begin VB.Form 关于 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "关于"
   ClientHeight    =   3480
   ClientLeft      =   10005
   ClientTop       =   6090
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4575
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "许可协议"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "GPL v3.0 or later"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label GoBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Go >"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label SiteLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "官网"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "这是Wiess Lab中的一个玩具"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "云夏神社"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "制作"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "编译"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "P-代码"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label veri 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "无法获取!"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "版本"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label updateBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "作者"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label 作者 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Riband"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label GotoSiteBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   ">"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "关于"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Version = "V-" + CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    veri.Caption = Version
    MsgBox "本软件仅供学习与交流,禁止任何形式的商业用途和非法用途,对于非法使用本软件的损失,作者不承担任何责任,希望使用者遵纪守法"
End Sub

Private Sub GoBtn_Click()
    nul = Shell("cmd /c start https://riband.github.io/ToyCmdBuilder/", vbHide)
End Sub


Private Sub Label9_Click()
    nul = Shell("cmd /c start https://spdx.org/licenses/GPL-3.0-or-later.html#licenseText", vbHide)
End Sub
