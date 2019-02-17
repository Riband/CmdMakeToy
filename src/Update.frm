VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Update 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "¼ì²é¸üÐÂ"
   ClientHeight    =   2505
   ClientLeft      =   11685
   ClientTop       =   6540
   ClientWidth     =   4185
   Icon            =   "Update.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4185
   Begin VB.TextBox urlbox 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Update.frx":000C
      Top             =   4080
      Width           =   3735
   End
   Begin VB.TextBox UpdateUrlBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      ExtentX         =   7223
      ExtentY         =   4260
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label FlashUrlBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label WebSiteBtn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      Caption         =   "¹ÙÍø"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label UpdateCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "¸üÐÂµØÖ·"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ver As String

Private Sub FlashUrlBtn_Click()
WebBrowser1.Navigate2 UpdateUrlBox.Text & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision)
End Sub

Sub Form_Resize()
WebBrowser1.Width = Update.Width - 50
'WebBrowser1.Height = Update.Height - 100
End Sub

Private Sub Form_Load()
Ver = CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision) 'like "3.23"
UpdateUrlBox.Text = "https://riband.bitbucket.io/inner/update/update.html#ToyCmdBuilder="
WebBrowser1.Navigate2 UpdateUrlBox.Text & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision)
urlbox.Text = "url:" + UpdateUrlBox.Text & CStr(App.Major) & "." & CStr(App.Minor) & CStr(App.Revision)
End Sub


Private Sub WebSiteBtn_Click()
Shell "start www.example.com"
End Sub
