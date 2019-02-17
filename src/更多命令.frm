VERSION 5.00
Begin VB.Form 更多命令 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "命令"
   ClientHeight    =   2940
   ClientLeft      =   10395
   ClientTop       =   5475
   ClientWidth     =   3000
   Icon            =   "更多命令.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3000
   Begin VB.Label exit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label 我的电脑 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "我的电脑"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      ToolTipText     =   "其实应该叫资源管理器"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label 写字板 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "写字板"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label cmd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "弹窗(cmd) "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label 录音机 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "录音机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label 画图 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "画图"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label 记事本 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "记事本"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label 计算器 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "计算器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label 屏幕键盘 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "屏幕键盘"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label 检查Windows版本 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "检查系统版本"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "更多命令"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
AddCB ("exit")
End Sub

Private Sub 断网_Click()
MsgBox "只对宽带连接/ADSL有效!"
AddCB ("rasphone -h 宽带连接")
AddCB ("rasphone -h ADSL")
End Sub

Private Sub 计算器_Click()
AddCB ("calc")
End Sub

Private Sub cmd_Click()
AddCB ("cmd")
End Sub

Private Sub 我的电脑_Click()
AddCB ("explorer")
End Sub

Private Sub 画图_Click()
AddCB ("mspaint")
End Sub

Private Sub 记事本_Click()
a = MsgBox("是否要打开一个文件？", vbYesNo, "是否要打开一个文件？")
If a = vbYes Then
AddCB ("notepad " + (InputBox("文件名")))
Else
AddCB ("notepad")
End If
End Sub

Private Sub 录音机_Click()
AddCB ("sndrec32")
End Sub

Private Sub 检查Windows版本_Click()
AddCB ("winver")
End Sub

Public Function AddCB(cmd As String)
CMDMAKER.CodeBox.Text = CMDMAKER.CodeBox.Text + "start " + cmd + vbCrLf
End Function

Private Sub 屏幕键盘_Click()
AddCB ("osk")
End Sub

Private Sub 写字板_Click()
a = MsgBox("是否要打开一个文件？", vbYesNo, "是否要打开一个文件？")
If a = vbYes Then
AddCB ("write " + (InputBox("文件名")))
Else
AddCB ("write")
End If
End Sub

Private Sub 修改开机密码_Click()
MsgBox "这非常危险,请自行承担后果"
pwdnew = InputBox("密码改为?")
If pwdnew <> "" Then
CMDMAKER.CodeBox.Text = CMDMAKER.CodeBox.Text + "net user %username% " + pwdnew + vbCrLf
End If
End Sub
