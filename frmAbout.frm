VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtDisclaimer 
      BackColor       =   &H8000000F&
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   25
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6735
   End
   Begin VB.Image imgTitle 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      Caption         =   "http://www.geocities.com/lio889"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":1994
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1440
      Width           =   3270
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "lio_889@ziplip.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      MouseIcon       =   "frmAbout.frx":1C9E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Label lblHomePageCap 
      AutoSize        =   -1  'True
      Caption         =   "Home Page:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label lblEmailCap 
      AutoSize        =   -1  'True
      Caption         =   "E-mail:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "Copyright (C) August/2000 Khaery Rida."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "crackPassword"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2625
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
frmAbout.Hide

End Sub

Private Sub Form_Load()
txtDisclaimer.Text = "This program is 100% FREEWARE." & Chr$(13) & Chr$(10)
txtDisclaimer.Text = txtDisclaimer.Text & "You may copy and/or distribute this program (or any portion of it) in any way you may find it useful." & Chr$(13) & Chr$(10)
txtDisclaimer.Text = txtDisclaimer.Text & "I can NOT be responsibility for any damage and/or loss of data of any kind caused by this program. USE IT AT YOUR OWN RISK!" & Chr$(13) & Chr$(10)
txtDisclaimer.Text = txtDisclaimer.Text & "I accept no responsibility if your boss/tutor/parents/friends/cat catch you using it!" & Chr$(13) & Chr$(10)
txtDisclaimer.Text = txtDisclaimer.Text & "By using this program, you accept the conditions above." + Chr$(13) + Chr$(10)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
frmAbout.Hide

End Sub

Private Sub lblEmail_Click()
RetVal = ShellExecute(Me.hwnd, vbNullString, "mailto:lio_889@ziplip.com?subject=" & MainTitle, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL)

End Sub

Private Sub lblHomePage_Click()
RetVal = ShellExecute(Me.hwnd, vbNullString, "http://www.geocities.com/lio889", vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL)

End Sub

