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
      Name            =   "΢���ź�"
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
      Caption         =   "VBS/JS����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��Ϣ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "����VBS����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��Ϣ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "ѡ��"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   10080
      TabIndex        =   52
      Top             =   120
      Width           =   2775
      Begin VB.CheckBox UsersEditBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "�ֶ��༭"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "ѭ�����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "���ļ�������pasue"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "���ByCMTע��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "����VBS����֧��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "���ڿ���"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�˳�����������ʱ�ļ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��������ʱ�ļ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ������л��� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "���������ʱ�ļ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
   Begin VB.Frame ���ڿ��� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "���ڿ���"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��������"
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
         Caption         =   "��С��"
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
      Caption         =   "�Զ�����"
      BeginProperty Font 
         Name            =   "����"
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
            Name            =   "΢���ź�"
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
         Caption         =   "    �Զ�����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label AutoSaveȷ�� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "ȷ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Begin VB.Label ��ʾ�������� 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "���ñ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "���� CMD.EXE �Ự�Ĵ��ڱ���"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label COLOR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "������ɫ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "����Ĭ�Ͽ���̨ǰ���ͱ�����ɫ"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label wait 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "�ȴ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "�ļ�"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�����ļ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "ɾ������һ���ļ�"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label type 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "��ȡ����ʾ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�����ļ�(����ļ���)"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "������һ���ļ����Ƶ���һ��λ��"
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame while 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "˳�����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "����ѭ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��ת"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.Label ��� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�ظ�����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "ѭ����β"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "����ѭ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "�ػ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "��ʾ"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�����Ļ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "�����Ļ"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label input_ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "ͣס����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "��ͣ�������ļ��Ĵ�����ʾ��Ϣ"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label echo_noenter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "��ʾ(������)"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��ʾ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Caption         =   "��ʾ"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         ToolTipText     =   "��ʾ��Ϣ����������Դ򿪻�ر�"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox FilePathBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�½�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
   Begin VB.Label ״̬�� 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�汾����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "ִ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "д���ļ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���ļ�����λ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "д�벢ִ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�ļ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
    ״̬��.Caption = Str(Time) + "�Զ�����һ��"

End Sub

Private Sub AutoSaveBtn_Click()
    AutoSave.Enabled = AutoSaveBtn.Value
    AutoSaveTimeBox.Enabled = AutoSaveBtn.Value
    AutoSaveȷ��.Enabled = AutoSaveBtn.Value
End Sub


Private Sub Bigger_Click()
    BeBigger (3000)
End Sub

Private Sub cleaan_Click()
    i = MsgBox("ɾ������tmp*.cmd�ļ���", vbYesNo, "����?")
    If i <> "" Then
        Kill App.Path + "\tmp*.cmd"
    End If
End Sub


Private Sub clean_Click()
    ״̬��.Caption = "��ʼ����"
    If UsingTmpFile = True And Dir(FileName) <> "" Then
        Kill FileName
        ״̬��.Caption = "�������"
    Else
        ״̬��.Caption = "û�з�����ʱ�ļ�"
    End If
End Sub



Private Sub ClearForWithoutEnd_Click()
    If MsgBox("ע��,��Ӧ����ѭ����������˻��Ҳ���Ҫʹ��ѭ�����ʱ������,����?", vbOKCancel) = 1 Then
        ForWithoutEnd = 0
        ForEnd.BackColor = &H808000
        ״̬��.Caption = "ѭ����������"
    End If
End Sub

Private Sub CLS_Click()
    CodeBox.Text = CodeBox.Text + "cls" + vbCrLf
End Sub

Private Sub cmd_Click()
    CodeBox.Text = CodeBox.Text + "start cmd" + vbCrLf
End Sub

Private Sub COLOR_Click()
    AddToText "������ɫ,����16������,��һ��������ɫ,�ڶ�������ǰ��ɫ ��:9F" & vbCrLf & "0=��" & vbCrLf & "1=��" & vbCrLf & "2=��" & vbCrLf & "3=ǳ��" & vbCrLf & "4=��" & vbCrLf & "5=��" & vbCrLf & "6=��" & vbCrLf & "7=��" & vbCrLf & "8=��" & vbCrLf & "9=����" & vbCrLf & "A=����" & vbCrLf & "B=��ǳ��" & vbCrLf & "C=����" & vbCrLf & "D=����" & vbCrLf & "E=����" & vbCrLf & "F=����", "color", "color ", ""
End Sub

Private Sub CreateCmdFileBtn_Click()

Dim i As String
    i = InputBox("�ļ���?(����.cmd)", "����")
    If i <> "" Then
        FileName = "\" + i + ".cmd"
        Shell "cmd /c echo pause > " + Chr(34) + App.Path + FileName + Chr(34), vbHide
        FilePathBox.Text = App.Path + "\cmertmp\" + i + ".cmd"
        UsingTmpFile = False
    End If
End Sub


Private Sub copy_Click()
    fp = InputBox("��·�����ļ���?(���� C:\1.txt)")
    If fp <> "" Then
        np = InputBox("��·��?(���� C:\)")
    
        If np <> "" Then
           CodeBox.Text = CodeBox.Text + "copy " + fp + " " + np
        End If
    End If
End Sub

Private Sub creadfile_Click()
    inputtext = InputBox("�ļ����ݣ�")
    If inputtext <> "" Then
        AddToText "�ļ�����", "echo  >", "echo " + inputtext + " > ", ""
    End If
End Sub




Private Sub DEL_Click()
    AddToText "ɾ��ʲô��", "del", "del ", ""
End Sub

Private Sub echo_Click()
    Dim i As String
    AddToText "��ʾʲô?", "echo", "echo ", ""
End Sub

Private Sub echo_noenter_Click()
    AddToText "��ʾʲô?", "echo", "set /p  = ", " < nul"
End Sub

Private Sub echobl_Click()
    AddToText "��ʾʲô����?", "echo", "echo %", "%"
End Sub


Private Sub exit_Click()
    End
End Sub

'Private Sub exitclean_Click()
'If Dir(App.Path + tmpfile) <> "" Then
'����
'If UsingTmpFile = True Then
'Kill FileName
'Else
'״̬��.Caption = "û��ʹ����ʱ�ļ�"
'End If
'End

'End Sub

Private Sub exp_Click()
    Shell "explorer " + Chr(34) + App.Path + "\cmertmp\", vbNormalFocus
End Sub


Private Sub for_Click()
    MsgBox ("�����ڿ�ʼ��ֱ���㵥�� ����ѭ��-�յ� ��ť ֮������д��붼���ظ�ִ���㽫����Ĵ���" + vbCrLf + "ע��:ÿ��ѭ����ʼ�����Ӧһ��ѭ������")
    Dim i As String
    i = InputBox("ѭ������?(���65535)", "for /l")
    If i <> "" Then
        CodeBox.Text = CodeBox.Text + "for /l %%i in (1,1," + i + ") do (" + vbCrLf
        If ForCheckAble.Value = 1 Then
            ForWithoutEnd = ForWithoutEnd + 1
            ForEnd.BackColor = &HFF&
            ״̬��.Caption = "ѭ����ʼ,����ѭ������ʱ����ѭ��-��β"
        End If
    End If
End Sub

Private Sub ForEnd_Click()
    If ForCheckAble.Value = 1 Then
        If ForWithoutEnd <= 0 Then
            i = MsgBox("ѭ����ʼ�ͽ�β�������ʹ��,�ݼ��,û���㹻���ѭ����ʼ,���������β,��Ҫ�����β?", vbOKCancel, "ѭ�����")
                If i = 1 Then
                    CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                    ForWithoutEnd = ForWithoutEnd - 1
                End If
        End If
        
        If ForWithoutEnd = 1 Then
                ForEnd.BackColor = &H808000
                ForWithoutEnd = ForWithoutEnd - 1
                CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                ״̬��.Caption = "ѭ����β"
        End If
        
        If ForWithoutEnd > 1 Then
                ForWithoutEnd = ForWithoutEnd - 1
                CodeBox.Text = CodeBox.Text + ")" + vbCrLf
                ״̬��.Caption = "ѭ����β"
        End If
        
    End If
End Sub

Private Sub ForLoop_Click()
   MsgBox ("�����ڿ�ʼ��ֱ���㵥�� ѭ��-���� ��ť ֮������д��붼�������ظ�ִ��(������)" + vbCrLf + "ע��:ÿ��ѭ����ʼ�����Ӧһ��ѭ������")
   CodeBox.Text = CodeBox.Text + "for /l %%i in (1,0,1) do (" + vbCrLf
   If ForCheckAble.Value = 1 Then
        ForEnd.BackColor = &HFF&
        ForWithoutEnd = ForWithoutEnd + 1
        ״̬��.Caption = "ѭ����ʼ"
   End If
End Sub

Private Sub Form_Load()
'���ð汾
    Version = "V-" + CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
    Ver.Caption = Version
    WindowFlag = True

'����TMP·��
    Shell "cmd /c md " + App.Path + "\cmertmp\", vbHide
'����TMP�ļ�
Do
    FileName = App.Path + "\cmertmp\tmp" + CStr(Minute(Time)) + CStr(Second(Time)) + ".cmd"
Loop While Dir(FileName) <> "" '����

     Open FileName For Output As #3
     Print #3, "@echo off"
     Close #3
    'Shell "cmd /c echo pause > " + FileName, vbHide
    FilePathBox.Text = FileName
    UsingTmpFile = True
'VBS֧��
    VBSSupportCode = "goto end" + vbCrLf
    VBSSupportCode = VBSSupportCode + ":vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "set v=%1" + vbCrLf
    VBSSupportCode = VBSSupportCode + "echo %v:~1,-1% > %~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "%~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + "del %~dp0tmp12.vbs" + vbCrLf
    VBSSupportCode = VBSSupportCode + ":end" + vbCrLf
    '����С��
    WindowControl = 1
    'Forѭ����������
    ForWithoutEnd = 0
    '���д��ִ�����ŵ�bug
    RunAble = True
    '��������VBS Temp
    VBSSupport.Value = 1
    VBSSupport_Click
End Sub



Private Sub input__Click()
    AddToText "��ʲô��Ϊ���������������?", "set /p", "set /p ", "="
End Sub


Private Sub Label2_Click()
    ����.Show
End Sub

Private Sub Label3_Click()
    a = Shell("cmd /k help", vbNormalFocus)
End Sub

Private Sub MshtaMsgBoxBtn_Click()
    AddToText "��ʾ������", "VBS:MsgBox", "mshta vbscript:msgbox(" & Chr(34), Chr(34) + ",64," & Chr(34) & "��ʾ" & Chr(34) & ")(window.close)"
End Sub

Private Sub ReloadBtn_Click()
ReloadBtn.Enabled = False 'UI
״̬��.Caption = "������..."

Dim ReloadFileInput As String

Open FileName For Input As #1
    Input #1, ReloadFileInput
    CodeBox.Text = ReloadFileInput
Close #1

ReloadBtn.Enabled = True 'UI
    ״̬��.Caption = "�������"
End Sub

Private Sub UsersEditBtn_Click()
    CodeBox.Enabled = UsersEditBtn.Value
End Sub

Private Sub VBSCDBtn_Click()
    CodeBox.Text = CodeBox.Text + "call :vbs " & Chr(34) & "createobject(" & Chr(34) & "wmplayer.ocx" & Chr(34) & ").cdromcollection.item(0).eject " & Chr(34) & vbCrLf
End Sub

Private Sub ��ʾ��������_Click()
��������.Show
End Sub

Private Sub loopall_Click()
    a = MsgBox("����Է�ֹ�رմ��ڽ�������ѭ��(һ���̶���),֮ǰ���������ݽ���������ѭ����������", vbYesNo, "�Ƿ������")
    If a <> 7 Then
        CodeBox.Text = CodeBox.Text + "start %~f0" + vbCrLf
    End If
End Sub

Private Sub More_Click()
    ��������.Show
End Sub

Private Sub pause_Click()
    CodeBox.Text = CodeBox.Text + "pause" + vbCrLf
End Sub

Private Sub randfile_Click()
    inputtext = InputBox("�ļ����ݣ�")
    CodeBox.Text = CodeBox.Text + "echo " + inputtext + " > %random%%%i.txt" + vbCrLf
End Sub

Private Sub RunCmdFile_Click()
    If RunAble = True Then
        ״̬��.Caption = "ִ����"
        Shell FileName, vbNormalFocus
        ״̬��.Caption = "ִ�����"
        Else
        ״̬��.Caption = "�ܾ�ִ��,������д��"
    End If
End Sub

Private Sub shutdown_Click()
    AddToText "����ʱ���룿", "shutdown", "shutdown -s -t ", ""
End Sub

'Me.BorderStyle = 2

Private Sub start_Click()
    AddToText "����ʲô", "start", "start ", ""
End Sub

Private Sub title_Click()
    AddToText "��������Ϊ:", "title", "title ", ""
End Sub


Private Sub type_Click()
    AddToText "�ļ�����", "type", "type ", ""
End Sub

Private Sub updateBTN_Click()
    Update.Show
End Sub

Private Sub VbsMsgbox_Click()
AddToText "��ʾ������", "VBS:MsgBox", "call :vbs " & Chr(34) & "MsgBox " & Chr(34), Chr(34) & Chr(34)
End Sub

Private Sub VBSSupport_Click()
    ClassicVBSFrm.Enabled = VBSSupport.Value
    ClassicVBSFrm.Visible = VBSSupport.Value
    Dim tmpVAble As Boolean '��� ScriptFrm.Visible = Not VBSSupport.Value ��Ч
    tmpVAble = VBSSupport.Value
    ScriptFrm.Enabled = Not tmpVAble
    ScriptFrm.Visible = Not tmpVAble
End Sub

Private Sub Ver_Click()
    ����.Show
End Sub

Private Sub wait_Click()
    i = InputBox("�ȴ����룿", "ping -n * 127.0.0.1 > nul")
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
        ForCheckRequest = MsgBox("��" + CStr(ForWithoutEnd) + "��ѭ��û�н�β,��Ҫд��?", vbOKCancel, "ѭ�����")
        If ForCheckRequest <> 1 Then
            ״̬��.Caption = "ȡ����ִ��"
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
    ״̬��.Caption = "д�����"
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
        Bigger.Caption = "��"
        WindowFlag = False
    Else
        For i = 0 To 10
            CMDMAKER.Width = CMDMAKER.Width - once
        Next i
        Bigger.Left = Bigger.Left - wight
        Bigger.Caption = "��"
        WindowFlag = True
    End If
End Function



Private Sub ���_Click()
    AddToText "��������:" + vbCrLf + "��Ӧ��ע��,��ʹ����ת���ѭ��ʱ,ĳЩ��ȫ����ᱨ��", "��� :", ":", ""
End Sub

Private Sub goto_Click()
    AddToText "��ת���ĸ����? " + vbCrLf + "��Ӧ��ע��,��ʹ����ת���ѭ��ʱ,ĳЩ��ȫ����ᱨ��", "��ת goto", "goto ", ""
End Sub

Private Sub ������л���_Click()
    ״̬��.Caption = "��ʼ����"
    If Dir(App.Path + "/tmp*.cmd") <> "" Then  '����
        a = MsgBox("���ɾ��������tmp��ͷ,��.cmd��β���ļ�", vbYesNo, "�Ƿ������")
        If a = vbYes Then
            Kill App.Path + "/tmp*.cmd"
            ״̬��.Caption = "�������"
        End If
    Else
        ״̬��.Caption = "û�з�����ʱ�ļ�"
    End If

End Sub


Private Sub Form_Unload(Index As Integer)
'If  Then
'����
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

'Private Sub ʹ��RichTextBox_Click()
'i = MsgBox("����:����һ��ʵ���Թ���,����֧�ֳ���65KB���ĵ�,��һ�㲻���õ���,������֪��������ʲô", vbYesNo, "����")
'End Sub
Private Sub AutoSaveȷ��_Click()
    AutoSaveTime = Val(AutoSaveTimeBox)
    If AutoSaveTime > 0 And AutoSaveTime < 61 Then
        AutoSave.Interval = AutoSaveTime * 1000
    Else
        MsgBox "����Ƿ�:" + CStr(AutoSaveTime)
        AutoSaveTimeBox = "20"
    End If
End Sub


