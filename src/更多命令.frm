VERSION 5.00
Begin VB.Form �������� 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   2940
   ClientLeft      =   10395
   ClientTop       =   5475
   ClientWidth     =   3000
   Icon            =   "��������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   3000
   Begin VB.Label exit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "�˳�"
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
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label �ҵĵ��� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "�ҵĵ���"
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
      TabIndex        =   8
      ToolTipText     =   "��ʵӦ�ý���Դ������"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label д�ְ� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "д�ְ�"
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
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label cmd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "����(cmd) "
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
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ¼���� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "¼����"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label ��ͼ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "��ͼ"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label ���±� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "���±�"
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
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label ������ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "������"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label ��Ļ���� 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "��Ļ����"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ���Windows�汾 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      Caption         =   "���ϵͳ�汾"
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
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
AddCB ("exit")
End Sub

Private Sub ����_Click()
MsgBox "ֻ�Կ������/ADSL��Ч!"
AddCB ("rasphone -h �������")
AddCB ("rasphone -h ADSL")
End Sub

Private Sub ������_Click()
AddCB ("calc")
End Sub

Private Sub cmd_Click()
AddCB ("cmd")
End Sub

Private Sub �ҵĵ���_Click()
AddCB ("explorer")
End Sub

Private Sub ��ͼ_Click()
AddCB ("mspaint")
End Sub

Private Sub ���±�_Click()
a = MsgBox("�Ƿ�Ҫ��һ���ļ���", vbYesNo, "�Ƿ�Ҫ��һ���ļ���")
If a = vbYes Then
AddCB ("notepad " + (InputBox("�ļ���")))
Else
AddCB ("notepad")
End If
End Sub

Private Sub ¼����_Click()
AddCB ("sndrec32")
End Sub

Private Sub ���Windows�汾_Click()
AddCB ("winver")
End Sub

Public Function AddCB(cmd As String)
CMDMAKER.CodeBox.Text = CMDMAKER.CodeBox.Text + "start " + cmd + vbCrLf
End Function

Private Sub ��Ļ����_Click()
AddCB ("osk")
End Sub

Private Sub д�ְ�_Click()
a = MsgBox("�Ƿ�Ҫ��һ���ļ���", vbYesNo, "�Ƿ�Ҫ��һ���ļ���")
If a = vbYes Then
AddCB ("write " + (InputBox("�ļ���")))
Else
AddCB ("write")
End If
End Sub

Private Sub �޸Ŀ�������_Click()
MsgBox "��ǳ�Σ��,�����ге����"
pwdnew = InputBox("�����Ϊ?")
If pwdnew <> "" Then
CMDMAKER.CodeBox.Text = CMDMAKER.CodeBox.Text + "net user %username% " + pwdnew + vbCrLf
End If
End Sub
