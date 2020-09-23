VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "crackPassword"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTextBox 
      Caption         =   "Notify when targeting a TextBox"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   4560
      ScaleHeight     =   1875
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblYCap 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblXCap 
         AutoSize        =   -1  'True
         Caption         =   "X:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblTrack 
         AutoSize        =   -1  'True
         Caption         =   "Targeting:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Target with this"
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
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   930
      End
      Begin VB.Image imgTarget 
         Height          =   480
         Left            =   360
         Picture         =   "frmMain.frx":0CCA
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.CheckBox chkOnTop 
      Caption         =   "Always keep this window on top"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label lblRelease 
      AutoSize        =   -1  'True
      Caption         =   "Release!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image imgNull 
      Height          =   15
      Left            =   840
      Picture         =   "frmMain.frx":1994
      Top             =   2880
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Image imgCross 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":1DDA
      Top             =   2880
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Output"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":2AA4
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkOnTop_Click()

If chkOnTop.Value = 1 Then
               SetWindowPos frmMain.hwnd, HWND_TOPMOST, frmMain.Left / 15, _
                            frmMain.Top / 15, frmMain.Width / 15, _
                            frmMain.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else
               SetWindowPos frmMain.hwnd, HWND_NOTOPMOST, frmMain.Left / 15, _
                            frmMain.Top / 15, frmMain.Width / 15, _
                            frmMain.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If

End Sub

Private Sub cmdAbout_Click()
frmAbout.Show 1

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub Form_Load()
Load frmAbout
chkTextBox.Value = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmAbout
    End

End Sub

Private Sub imgTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
' Visual effects
    Targeting = True
    imgTarget.Picture = imgNull.Picture
    lblTrack.Visible = True
    lblXCap.Visible = True
    lblYCap.Visible = True
    lblX.Visible = True
    lblY.Visible = True
    
    Me.MousePointer = 99
    Me.MouseIcon = imgCross.Picture
    txtOutput.Text = ""
        
End Sub

Private Sub imgTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim sName As String, sClassName As String * 255, TempHwnd As Long
    
    If Targeting = False Then Exit Sub
    Call GetCursorPos(CursorPosition)

    ' Display mouse's location
    lblX.Caption = CursorPosition.X
    lblY.Caption = CursorPosition.Y

    ' Check whether the cursor is pointing to a TextBox or not
    TempHwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y)
    Call GetClassName(TempHwnd, sClassName, 255)
    sName = Trim(Left(sClassName, InStr(sClassName, vbNullChar) - 1))
    
    If chkTextBox.Value = 1 Then
    
    If sName = "Edit" Or InStr(sName, "TextBox") > 0 Then
        lblRelease.Visible = True
    Else
        lblRelease.Visible = False
    End If
   End If
   

End Sub

Private Sub imgTarget_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim TargetLen As Long, TempString As String, hwnd As Long
' Visual effects
    Targeting = False
    lblTrack.Visible = False
    lblXCap.Visible = False
    lblYCap.Visible = False
    lblX.Visible = False
    lblY.Visible = False
    lblRelease.Visible = False
    
    imgTarget.Picture = imgCross.Picture
    Me.MousePointer = 0

'
Call GetCursorPos(CursorPosition)
    hwnd = WindowFromPoint(CursorPosition.X, CursorPosition.Y) ' Get target window's handle
    hwnd = GetTopLevelParent(hwnd)          ' Get target window's parent's handle
    TargetLen& = SendMessage(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
    TempString$ = String(TargetLen&, 0&)
    Call sendmessagebystring(hwnd&, WM_GETTEXT, TargetLen& + 1, TempString$)
    txtOutput.Text = TempString$


End Sub
