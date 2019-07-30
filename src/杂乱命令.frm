VERSION 5.00
Begin VB.Form 杂乱命令 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChkdskBtn"
   ClientHeight    =   1530
   ClientLeft      =   10590
   ClientTop       =   6885
   ClientWidth     =   2490
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2490
   Begin VB.Label DateSetBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "修改日期"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "设置默认控制台前景和背景颜色"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label TimeSetBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "修改时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
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
      ToolTipText     =   "设置默认控制台前景和背景颜色"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label ChkdskBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "chkdsk磁盘检查工具"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "设置默认控制台前景和背景颜色"
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Tskill 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "结束进程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "设置默认控制台前景和背景颜色"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "杂乱命令"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ChkdskBtn_Click()
CMDMAKER.AddToText "这可以检查并修复磁盘中的一些错误" + vbCrLf + "盘符(例如: C:),如果要修复,请在盘符前/后加入一个 /f (例如: C: /f)", "chkdsk", "chkdsk ", ""
End Sub

Private Sub DateSetBtn_Click()
CMDMAKER.AddToText "把日期改为(年月日 如 2016-04-01)", "date", "date ", ""
End Sub

Private Sub PingAttackBtn_Click()
MsgBox "仅供学习和研究,切勿用于非法用途,造成的一切后果与损失与作者无关!" + vbCrLf + "注意:单个人攻击效果可能不明显,同时攻击的人越多越好"
CMDMAKER.AddToText "请输入ip地址或网址" + vbCrLf + "不带http:// 例如 www.example.com 或 111.222.111.222", "ping -l 65500 -n 65535", "ping -l 65500 -n 65535 ", ""
End Sub

Private Sub TimeSetBtn_Click()
CMDMAKER.AddToText "把时间改为(例如 8:50:00)", "time", "time ", ""
End Sub

Private Sub Tskill_Click()
CMDMAKER.AddToText "要结束的进程名(不带.exe)", "tskill", "tskill ", ""
End Sub
