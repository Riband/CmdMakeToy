VERSION 5.00
Begin VB.Form �������� 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChkdskBtn"
   ClientHeight    =   1530
   ClientLeft      =   10590
   ClientTop       =   6885
   ClientWidth     =   2490
   BeginProperty Font 
      Name            =   "΢���ź�"
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
      Caption         =   "�޸�����"
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
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "����Ĭ�Ͽ���̨ǰ���ͱ�����ɫ"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label TimeSetBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "�޸�ʱ��"
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
      TabIndex        =   2
      ToolTipText     =   "����Ĭ�Ͽ���̨ǰ���ͱ�����ɫ"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label ChkdskBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "chkdsk���̼�鹤��"
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
      TabIndex        =   1
      ToolTipText     =   "����Ĭ�Ͽ���̨ǰ���ͱ�����ɫ"
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Tskill 
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "����Ĭ�Ͽ���̨ǰ���ͱ�����ɫ"
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ChkdskBtn_Click()
CMDMAKER.AddToText "����Լ�鲢�޸������е�һЩ����" + vbCrLf + "�̷�(����: C:),���Ҫ�޸�,�����̷�ǰ/�����һ�� /f (����: C: /f)", "chkdsk", "chkdsk ", ""
End Sub

Private Sub DateSetBtn_Click()
CMDMAKER.AddToText "�����ڸ�Ϊ(������ �� 2016-04-01)", "date", "date ", ""
End Sub

Private Sub PingAttackBtn_Click()
MsgBox "����ѧϰ���о�,�������ڷǷ���;,��ɵ�һ�к������ʧ�������޹�!" + vbCrLf + "ע��:�����˹���Ч�����ܲ�����,ͬʱ��������Խ��Խ��"
CMDMAKER.AddToText "������ip��ַ����ַ" + vbCrLf + "����http:// ���� www.example.com �� 111.222.111.222", "ping -l 65500 -n 65535", "ping -l 65500 -n 65535 ", ""
End Sub

Private Sub TimeSetBtn_Click()
CMDMAKER.AddToText "��ʱ���Ϊ(���� 8:50:00)", "time", "time ", ""
End Sub

Private Sub Tskill_Click()
CMDMAKER.AddToText "Ҫ�����Ľ�����(����.exe)", "tskill", "tskill ", ""
End Sub
