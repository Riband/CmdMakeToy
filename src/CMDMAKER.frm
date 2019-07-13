VERSION 5.00
Begin VB.Form CMDMAKER 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ToyCmdBuilder"
   ClientHeight    =   8145
   ClientLeft      =   7155
   ClientTop       =   3285
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CMDMAKER.frx":0000
   LinkTopic       =   "CMDMAKER"
   ScaleHeight     =   8145
   ScaleWidth      =   10050
   Begin VB.Frame ScriptFrm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "VBS/JS调用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   65
      Top             =   6720
      Width           =   4335
      Begin VB.Label MshtaMsgBoxBtn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "消息框"
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
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame ClassicVBSFrm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "经典VBS调用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   61
      Top             =   6720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label VBSCDBtn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "弹出光驱"
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
         Left            =   1560
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label VbsMsgbox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "消息框"
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
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Opinion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "选项"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   10080
      TabIndex        =   52
      Top             =   120
      Width           =   2775
      Begin VB.CheckBox UsersEditBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "手动编辑"
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
         TabIndex        =   60
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CheckBox ForCheckAble 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "循环检查"
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
         TabIndex        =   57
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox AutoAddPause 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "在文件最后添加pasue"
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
         TabIndex        =   55
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox AddByCmdMaker 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "添加ByCMT注释"
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
         TabIndex        =   54
         Top             =   840
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox VBSSupport 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "经典VBS调用支持"
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
         TabIndex        =   53
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label ClearForWithoutEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         Caption         =   "计数清零"
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
         TabIndex        =   58
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "窗口控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   10080
      TabIndex        =   48
      Top             =   3000
      Width           =   2775
      Begin VB.Label exit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         Caption         =   "退出而不清理临时文件"
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
         TabIndex        =   51
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label clean 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "清理本次临时文件"
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
         TabIndex        =   50
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label 清除所有缓存 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "清除所有临时文件"
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
         TabIndex        =   49
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame 窗口控制 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "窗口控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   10080
      TabIndex        =   45
      Top             =   6120
      Width           =   2775
      Begin VB.OptionButton WindowNormal 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         Caption         =   "正常窗口"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton WindowMini 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         Caption         =   "最小化"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame AutoSaveFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "自动保存"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   10080
      TabIndex        =   40
      Top             =   4800
      Width           =   2775
      Begin VB.TextBox AutoSaveTimeBox 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   240
         TabIndex        =   42
         Text            =   "20"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Timer AutoSave 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2280
         Top             =   240
      End
      Begin VB.CheckBox AutoSaveBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "    自动保存"
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
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label AutoSave确定 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "确定"
         Enabled         =   0   'False
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
         Left            =   1920
         TabIndex        =   44
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "秒"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1560
         TabIndex        =   43
         Top             =   720
         Width           =   360
      End
   End
   Begin VB.Frame Other 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "其他"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   4335
      Begin VB.Label 显示杂乱命令 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "更多"
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
         Left            =   3120
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
      Begin VB.Label title 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "设置标题"
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
         Left            =   840
         TabIndex        =   39
         ToolTipText     =   "设置 CMD.EXE 会话的窗口标题"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label COLOR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "窗口颜色"
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
         Left            =   1920
         TabIndex        =   34
         ToolTipText     =   "设置默认控制台前景和背景颜色"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label wait 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "等待"
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
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame files 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "文件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   4335
      Begin VB.Label creadfile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "创建文件"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label DEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "删除"
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
         Left            =   1560
         TabIndex        =   14
         ToolTipText     =   "删除至少一个文件"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label type 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "读取并显示"
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
         Left            =   2760
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label randfile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "创建文件(随机文件名)"
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
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label copy 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "复制"
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
         Left            =   2760
         TabIndex        =   11
         ToolTipText     =   "将至少一个文件复制到另一个位置"
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame while 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "顺序控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   4335
      Begin VB.Label ForLoop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "无限循环"
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
         Left            =   1680
         TabIndex        =   56
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label goto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "跳转"
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
         Left            =   2880
         TabIndex        =   38
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label 标记 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "标记"
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
         Left            =   1680
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label loopall 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "重复所有"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label ForEnd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "循环结尾"
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
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label for 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "有限循环"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "功能"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   4335
      Begin VB.Label start 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "启动"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label shutdown 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "关机"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label More 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "更多"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox CodeBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      Height          =   6975
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Frame io 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
      Begin VB.Label CLS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "清除屏幕"
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
         TabIndex        =   24
         ToolTipText     =   "清除屏幕"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label input_ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "输入"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label pause 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "停住窗口"
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
         Left            =   2760
         TabIndex        =   22
         ToolTipText     =   "暂停批处理文件的处理并显示消息"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label echo_noenter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "显示(不换行)"
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
         Left            =   2760
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label echobl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "显示变量"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label echo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "显示"
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
         TabIndex        =   19
         ToolTipText     =   "显示消息，或将命令回显打开或关闭"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox FilePathBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   795
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label ReloadBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "重载"
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
      Left            =   6480
      TabIndex        =   64
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label CreateCmdFileBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "新建"
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
      TabIndex        =   26
      Top             =   540
      Width           =   615
   End
   Begin VB.Label updateBTN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "检查更新"
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
      Left            =   3240
      TabIndex        =   35
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label 状态栏 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "程序已启动"
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
      Left            =   240
      TabIndex        =   33
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Label Bigger 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9120
      TabIndex        =   32
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Ver 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "版本错误"
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
      Left            =   7920
      TabIndex        =   31
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label RunCmdFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "执行"
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
      Left            =   6480
      TabIndex        =   30
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label WriteToFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "写入文件"
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
      Left            =   4800
      TabIndex        =   29
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "打开文件所在位置"
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
      Left            =   7440
      TabIndex        =   28
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label WriteAndRun 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "写入并执行"
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
      Left            =   4800
      TabIndex        =   27
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label pat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "文件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "关于"
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
      Left            =   9120
      TabIndex        =   36
      Top             =   7680
      Width           =   855
   End
End
Attribute VB_Name = "CMDMAKER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FileName As String
Dim WindowFlag As Boolean
Dim Version As String
Dim UsingTmpFile As Boolean
Dim VBSSupportCode As String
Dim WindowControl As Integer
Dim ForWithoutEnd As Integer
Dim RunAble As Boolean



Private Sub AutoSave_Timer()
    Open FileName For Output As #1

    If AddByCmdMakerValue = 1 Then
        Print #1, vbCrLf + "::By ToyCmdBuilder of Wiess Lab"
    End If

    Print #1, CodeBox.Text

    If AutoAddPause.Value = 1 Then
        Print #1, vbCrLf + "pause"
        End If
    Close #1
    状态栏.Caption = Str(Time) + "自动保存一次"

End Sub

Private Sub AutoSaveBtn_Click()
    AutoSave.Enabled = AutoSaveBtn.Value
    AutoSaveTimeBox.Enabled = AutoSaveBtn.Value
    AutoSave确定.Enabled = AutoSaveBtn.Value
End Sub


Private Sub Bigger_Click()
    BeBigger (3000)
End Sub

Private Sub cleaan_Click()
    i = MsgBox("删除所有tmp*.cmd文件？", vbYesNo, "继续?")
    If i <> "" Then
        Kill App.Path + "\tmp*.cmd"
    End If
End Sub


Private Sub clean_Click()
    状态栏.Caption = "开始清理"
    If UsingTmpFile = True And Dir(FileName) <> "" Then
        Kill FileName
        状态栏.Caption = "清理结束"
    Else
        状态栏.Caption = "没有发现临时文件"
    End If
End Sub



Private Sub ClearForWithoutEnd_Click()
    If MsgBox("注意,这应该在循环检查引起了混乱并仍要使用循环检查时才清零,继续?", vbOKCancel) = 1 Then
        ForWithoutEnd = 0
        ForEnd.BackColor = &H808000
        状态栏.Caption = "循环计数清零"
    End If
End Sub

Private Sub CLS_Click()
    CodeBox.Text = CodeBox.Text + "cls" + vbCrLf
End Sub

Private Sub cmd_Click()
    CodeBox.Text = CodeBox.Text + "start cmd" + vbCrLf
End Sub

Private Sub COLOR_Click()
    AddToText "设置颜色,两个16进制数,第一个代表背景色,第二个代表前景色 如:9F" & vbCrLf & "0=黑" & vbCrLf & "1=蓝" & vbCrLf & "2=绿" & vbCrLf & "3=浅绿" & vbCrLf & "4=红" & vbCrLf & "5=紫" & vbCrLf & "6=黄" & vbCrLf & "7=白" & vbCrLf & "8=灰" & vbCrLf & "9=淡蓝" & vbCrLf & "A=淡绿" & vbCrLf & "B=淡浅绿" & vbCrLf & "C=淡红" & vbCrLf & "D=淡紫" & vbCrLf & "E=淡黄" & vbCrLf & "F=亮白", "color", "color ", ""
End Sub

Private Sub CreateCmdFileBtn_Click()

Dim i As String
    i = InputBox("文件名?(不带.cmd)", "生成")
    If i <> "" Then
        FileName = "\" + i + ".cmd"
        Shell "cmd /c echo pause > " + Chr(34) + App.Path + FileName + Chr(34), vbHide
        FilePathBox.Text = App.Path + "\cmertmp\" + i + ".cmd"
        UsingTmpFile = False
    End If
End Sub


Private Sub copy_Click()
    fp = InputBox("旧路径和文件名?(例如 C:\1.txt)")
    If fp <> "" Then
        np = InputBox("新路径?(例如 C:\)")
    
        If np <> "" Then
           CodeBox.Text = CodeBox.Text + "copy " + fp + " " + np
        End If
    End If
End Sub

Private Sub creadfile_Click()
    inputtext = InputBox("文件内容？")
    If inputtext <> "" Then
        AddToText "文件名？", "echo  >", "echo " + inputtext + " > ", ""
    End If
End Sub




Private Sub DEL_Click()
    AddToText "删除什么？", "del", "del ", ""
End Sub

Private Sub echo_Click()
    Dim i As String
    AddToText "显示什么?", "echo", "echo ", ""
End Sub

Private Sub echo_noenter_Click()
    AddToText "显示什么?", "echo", "set /p  = ", " < nul"
End Sub

Private Sub echobl_Click()
    AddToText "显示什么变量?", "echo", "echo %", "%"
End Sub


Private Sub exit_Click()
    End
End Sub

'Private Sub exitclean_Click()
'If Dir(App.Path + tmpfile) <> "" Then
'存在
'If UsingTmpFile = True Then
'Kill FileName
'Else
'状态栏.Caption = "没有使用临时文件"
'End If
'End

'End Sub

Private Sub exp_Click()
    Shell "explorer " + Chr(34) + App.Path + "\cmertmp\", vbNormalFocus
End Sub


Private Sub for_Click()
    MsgBox ("从现在开始，直到你单击 有限循环-终点 按钮 之间的所有代码都将重复执行你将输入的次数" + vbCrLf + "注意:每个循环开始必须对应一个循环结束")
    Dim i As String
    i = InputBox("循环几次?(最大65535)", "for /l")
    If i <> "" Then
        CodeBox.Text = CodeBox.Text + "for /l %%i in (1,1," + i + ") do (" + vbCrLf
        If ForCheckAble.Value = 1 Then
            ForWithoutEnd = ForWithoutEnd + 1
            ForEnd.BackColor = &HFF&
            状态栏.Caption = "循环开始,请在循环结束时按下循环-结尾"
        End If
    End If
End Sub

Private Sub ForEnd_Click()
    If ForCheckAble.Value = 1 Then
        If ForWithoutEnd <= 0 Then
            i = MsgBox("循环开始和结尾必须配对使用,据检测,没有足够多的循环开始,所以无需结尾,仍要加入结尾?", vbOKCancel, "循环检查")
                If i = 1 Then
                    CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                    ForWithoutEnd = ForWithoutEnd - 1
                End If
        End If
        
        If ForWithoutEnd = 1 Then
                ForEnd.BackColor = &H808000
                ForWithoutEnd = ForWithoutEnd - 1
                CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                状态栏.Caption = "循环结尾"
        End If
        
        If ForWithoutEnd > 1 Then
                ForWithoutEnd = ForWithoutEnd - 1
                CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                状态栏.Caption = "循环结尾"
        End If
        
    End If
End Sub

Private Sub ForLoop_Click()
   MsgBox ("从现在开始，直到你单击 循环-结束 按钮 之间的所有代码都将无限重复执行(理论上)" + vbCrLf + "注意:每个循环开始必须对应一个循环结束")
   CodeBox.Text = CodeBox.Text + "for /l %%i in (1,0,1) do (" + vbCrLf
   If ForCheckAble.Value = 1 Then
        ForEnd.BackColor = &HFF&
        ForWithoutEnd = ForWithoutEnd + 1
        状态栏.Caption = "循环开始"
   End If
End Sub

Private Sub Form_Load()
'设置版本
    Version = "V-" + CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Ver.Caption = Version
    WindowFlag = True

'创建TMP路径
    Shell "cmd /c md " + App.Path + "\cmertmp\", vbHide
'创建TMP文件
Do
    FileName = App.Path + "\cmertmp\tmp" + CStr(Minute(Time)) + CStr(Second(Time)) + ".cmd"
Loop While Dir(FileName) <> "" '存在

     Open FileName For Output As #3
     Print #3, "@echo off"
     Close #3
    'Shell "cmd /c echo pause > " + FileName, vbHide
    FilePathBox.Text = FileName
    UsingTmpFile = True
'VBS支持
    VBSSupportCode = "goto end" + vbCrLf
    VBSSupportCode = VBSSupportCode + ":vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "set v=%1" + vbCrLf
    VBSSupportCode = VBSSupportCode + "echo %v:~1,-1% > %~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "%~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "del %~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + ":end" + vbCrLf
    '不最小化
    WindowControl = 1
    'For循环结束管理
    ForWithoutEnd = 0
    '解决写入执行蹩脚的bug
    RunAble = True
    '单击经典VBS Temp
    VBSSupport.Value = 1
    VBSSupport_Click
End Sub



Private Sub input__Click()
    AddToText "把什么作为输入变量的名字呢?", "set /p", "set /p ", "="
End Sub


Private Sub Label2_Click()
    关于.Show
End Sub

Private Sub Label3_Click()
    a = Shell("cmd /k help", vbNormalFocus)
End Sub

Private Sub MshtaMsgBoxBtn_Click()
    AddToText "显示的内容", "VBS:MsgBox", "mshta vbscript:msgbox(" & Chr(34), Chr(34) + ",64," & Chr(34) & "提示" & Chr(34) & ")(window.close)"
End Sub

Private Sub ReloadBtn_Click()
ReloadBtn.Enabled = False 'UI
状态栏.Caption = "重载中..."

Dim ReloadFileInput As String

Open FileName For Input As #1
    Input #1, ReloadFileInput
    CodeBox.Text = ReloadFileInput
Close #1

ReloadBtn.Enabled = True 'UI
    状态栏.Caption = "重载完成"
End Sub

Private Sub UsersEditBtn_Click()
    CodeBox.Enabled = UsersEditBtn.Value
End Sub

Private Sub VBSCDBtn_Click()
    CodeBox.Text = CodeBox.Text + "call :vbs " & Chr(34) & "createobject(" & Chr(34) & "wmplayer.ocx" & Chr(34) & ").cdromcollection.item(0).eject " & Chr(34) & vbCrLf
End Sub

Private Sub 显示杂乱命令_Click()
杂乱命令.Show
End Sub

Private Sub loopall_Click()
    a = MsgBox("这可以防止关闭窗口结束无限循环(一定程度上),之前的所有内容将进行无限循环，继续？", vbYesNo, "是否继续？")
    If a <> 7 Then
        CodeBox.Text = CodeBox.Text + "start %~f0" + vbCrLf
    End If
End Sub

Private Sub More_Click()
    更多命令.Show
End Sub

Private Sub pause_Click()
    CodeBox.Text = CodeBox.Text + "pause" + vbCrLf
End Sub

Private Sub randfile_Click()
    inputtext = InputBox("文件内容？")
    CodeBox.Text = CodeBox.Text + "echo " + inputtext + " > %random%%%i.txt" + vbCrLf
End Sub

Private Sub RunCmdFile_Click()
    If RunAble = True Then
        状态栏.Caption = "执行中"
        Shell FileName, vbNormalFocus
        状态栏.Caption = "执行完毕"
        Else
        状态栏.Caption = "拒绝执行,请重新写入"
    End If
End Sub

Private Sub shutdown_Click()
    AddToText "倒计时几秒？", "shutdown", "shutdown -s -t ", ""
End Sub

'Me.BorderStyle = 2

Private Sub start_Click()
    AddToText "启动什么", "start", "start ", ""
End Sub

Private Sub title_Click()
    AddToText "标题设置为:", "title", "title ", ""
End Sub


Private Sub type_Click()
    AddToText "文件名？", "type", "type ", ""
End Sub

Private Sub updateBTN_Click()
    Update.Show
End Sub

Private Sub VbsMsgbox_Click()
AddToText "显示的内容", "VBS:MsgBox", "call :vbs " & Chr(34) & "MsgBox " & Chr(34), Chr(34) & Chr(34)
End Sub

Private Sub VBSSupport_Click()
    ClassicVBSFrm.Enabled = VBSSupport.Value
    ClassicVBSFrm.Visible = VBSSupport.Value
    Dim tmpVAble As Boolean '解决 ScriptFrm.Visible = Not VBSSupport.Value 无效
    tmpVAble = VBSSupport.Value
    ScriptFrm.Enabled = Not tmpVAble
    ScriptFrm.Visible = Not tmpVAble
End Sub

Private Sub Ver_Click()
    关于.Show
End Sub

Private Sub wait_Click()
    i = InputBox("等待几秒？", "ping -n * 127.0.0.1 > nul")
If i <> "" Then
        i = CStr(Int(Val(i)) + 1)
        If i > 0 Then
            CodeBox.Text = CodeBox.Text + "ping -n " + i + " 127.0.0.1 > nul" + vbCrLf
        End If
End If
End Sub

Private Sub WindowMini_Click()
WindowControl = 2
End Sub

Private Sub WindowNormal_Click()
    WindowControl = 1
End Sub

Private Sub WriteToFile_Click()
'UI Act
WriteAndRun.Enabled = False
WriteToFile.Enabled = False

    If ForWithoutEnd > 0 And ForCheckAble.Value <> 1 Then
        ForCheckRequest = MsgBox("有" + CStr(ForWithoutEnd) + "个循环没有结尾,仍要写入?", vbOKCancel, "循环检查")
        If ForCheckRequest <> 1 Then
            状态栏.Caption = "取消了执行"
            RunAble = False
            GoTo DoNotWrite
        End If
        
    End If
    Open FileName For Output As #1
    If WindowControl <> 1 Then
        Print #1, Chr(64) & Chr(105) & Chr(102) & Chr(32) & Chr(34) & Chr(37) & Chr(49) & Chr(34) & Chr(32) & Chr(110) & Chr(101) & Chr(113) & Chr(32) & Chr(34) & Chr(117) & Chr(115) & Chr(101) & Chr(100) & Chr(34) & Chr(32) & Chr(40) & Chr(115) & Chr(116) & Chr(97) & Chr(114) & Chr(116) & Chr(32) & Chr(47) & Chr(109) & Chr(105) & Chr(110) & Chr(32) & Chr(37) & Chr(126) & Chr(102) & Chr(48) & Chr(32) & Chr(117) & Chr(115) & Chr(101) & Chr(100) & Chr(32) & Chr(38) & Chr(32) & Chr(101) & Chr(120) & Chr(105) & Chr(116) & Chr(41)
    End If
    If AddByCmdMaker.Value = 1 Then
        Print #1, "::By ToyCmdBuilder"
    End If
    
    Print #1, CodeBox.Text

    If AutoAddPause.Value = 1 Then
        Print #1, vbCrLf + "pause"
    End If
    
    If VBSSupport.Value = 1 Then
        Print #1, vbCrLf + VBSSupportCode
    End If
    Close #1
    RunAble = True
    状态栏.Caption = "写入完毕"
DoNotWrite:

'UI Act
WriteToFile.Enabled = True
WriteAndRun.Enabled = True
End Sub

Private Sub WriteAndRun_Click()
    WriteToFile_Click
    RunCmdFile_Click
End Sub

Public Function AddToText(Hint As String, title As String, BeforeInput As String, AfterInput As String)
    Dim i As String
    i = InputBox(Hint, title)
    If i <> "" Then
        CodeBox.Text = CodeBox.Text + BeforeInput + i + AfterInput + vbCrLf
    End If
End Function

Public Function BeBigger(wight As Integer)
    Dim i As Integer
    once = wight / 10
    If WindowFlag = True Then
        For i = 0 To 10
            CMDMAKER.Width = CMDMAKER.Width + once
        Next i
        Bigger.Left = Bigger.Left + wight
        Bigger.Caption = "←"
        WindowFlag = False
    Else
        For i = 0 To 10
            CMDMAKER.Width = CMDMAKER.Width - once
        Next i
        Bigger.Left = Bigger.Left - wight
        Bigger.Caption = "→"
        WindowFlag = True
    End If
End Function



Private Sub 标记_Click()
    AddToText "输入标记名:" + vbCrLf + "但应当注意,当使用跳转造成循环时,某些安全软件会报毒", "标记 :", ":", ""
End Sub

Private Sub goto_Click()
    AddToText "跳转到哪个标记? " + vbCrLf + "但应当注意,当使用跳转造成循环时,某些安全软件会报毒", "跳转 goto", "goto ", ""
End Sub

Private Sub 清除所有缓存_Click()
    状态栏.Caption = "开始清理"
    If Dir(App.Path + "/tmp*.cmd") <> "" Then  '存在
        a = MsgBox("这会删除所有以tmp开头,以.cmd结尾的文件", vbYesNo, "是否继续？")
        If a = vbYes Then
            Kill App.Path + "/tmp*.cmd"
            状态栏.Caption = "清理结束"
        End If
    Else
        状态栏.Caption = "没有发现临时文件"
    End If

End Sub


Private Sub Form_Unload(Index As Integer)
'If  Then
'存在
    If UsingTmpFile = True And CodeBox.Text = "@echo off" + vbCrLf And Dir(FileName) <> "" Then
       Kill FileName
    End If
    End
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button And vbLeftButton Then
'Me.Move Me.Left - mX + X, Me.Top - mY + Y
'End If
'End Sub

'Private Sub 使用RichTextBox_Click()
'i = MsgBox("警告:这是一项实验性功能,用来支持超过65KB的文档,您一般不会用到它,除非您知道您在做什么", vbYesNo, "警告")
'End Sub
Private Sub AutoSave确定_Click()
    AutoSaveTime = Val(AutoSaveTimeBox)
    If AutoSaveTime > 0 And AutoSaveTime < 61 Then
        AutoSave.Interval = AutoSaveTime * 1000
    Else
        MsgBox "输入非法:" + CStr(AutoSaveTime)
        AutoSaveTimeBox = "20"
    End If
End Sub


