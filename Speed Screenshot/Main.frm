VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stokes Speed Screenshot"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   6255
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaveFiles 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5280
      Top             =   0
   End
   Begin VB.CommandButton cmdFullPreview 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Full Preview"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdHide 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "screenshot"
      Top             =   5640
      Width           =   1095
   End
   Begin Speed_Screenshot.TrayControl TrayControl1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Speed Screenshot (Double Click to Save Screenshot)"
   End
   Begin VB.CheckBox chkWholeScreen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Capture Whole Screen"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4920
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtSSNum 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "0"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton cmdGetScreen 
      BackColor       =   &H00C0C000&
      Caption         =   "Get Screen!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox picSS 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   5715
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Picture Box"
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label lblSaveFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Files:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   6240
      Width           =   975
   End
   Begin VB.Line Line6 
      X1              =   240
      X2              =   6000
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Image imgSS 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   4335
      Left            =   240
      MouseIcon       =   "Main.frx":2E7A
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Image Box"
      Top             =   240
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Line Line5 
      X1              =   2760
      X2              =   2760
      Y1              =   4800
      Y2              =   6120
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   4800
      Y2              =   6600
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   240
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   4800
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6000
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblSSPrompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Menu mnuIconMenu 
      Caption         =   "Icon Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIconMenuGetScreen 
         Caption         =   "Get Screen"
      End
      Begin VB.Menu mnuIconMenuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconMenuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuIconMenuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuIconMenuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconMenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGetScreen 
         Caption         =   "&Get Screen"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "&Hide Me"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewPicBox 
         Caption         =   "&Picture Box (Default and faster)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewImgBox 
         Caption         =   "I&mage Box (Shows Everything but slower)"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsCapture 
         Caption         =   "&Capture Whole Screen"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsReset 
         Caption         =   "&Reset"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2

Private Sub Form_Load()
    Me.Caption = "Stokes Speed Screenshot (v" & App.Major & "." & App.Minor & App.Revision & ")"
    mnuIconMenuShow.Enabled = False
    TrayControl1.Enabled = True
    txtSaveFiles.Text = "" & App.Path
End Sub

Private Sub CopyToClipboard(ByVal form_only As Boolean)

  Dim alt_scan_code As Long

    If form_only Then
        alt_scan_code = MapVirtualKey(VK_MENU, 0)
        keybd_event VK_MENU, alt_scan_code, 0, 0
        DoEvents
    End If
    keybd_event VK_SNAPSHOT, 0, 0, 0
    DoEvents
    If form_only Then
        keybd_event VK_MENU, alt_scan_code, KEYEVENTF_KEYUP, 0
        DoEvents
    End If

End Sub

Private Sub cmdGetScreen_Click()

    If picSS.Visible = True Then
       picSS.Enabled = True
       CopyToClipboard (chkWholeScreen.Value = vbUnchecked)
       picSS.Picture = Clipboard.GetData
       imgSS.Picture = Clipboard.GetData
       SavePicture picSS.Picture, txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"
       txtSSNum.Text = txtSSNum.Text + 1
       cmdFullPreview.Enabled = True
    Else
       imgSS.Visible = True
       imgSS.Enabled = True
       CopyToClipboard (chkWholeScreen.Value = vbUnchecked)
       imgSS.Picture = Clipboard.GetData
       picSS.Picture = Clipboard.GetData
       SavePicture imgSS.Picture, txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"
       txtSSNum.Text = txtSSNum.Text + 1
       cmdFullPreview.Enabled = True
    End If
    
End Sub

Private Sub chkWholeScreen_Click()
    If chkWholeScreen.Value = vbChecked Then
       mnuOptionsCapture.Checked = True
    Else
       mnuOptionsCapture.Checked = False
    End If
End Sub

Private Sub cmdFullPreview_Click()

    If frmMain.imgSS.Picture = 0 Then
       MsgBox "No screenshot loaded!", vbCritical, "Oops!"
       Exit Sub
    Else
        frmShowFullScreen.imgFS.Picture = frmMain.imgSS.Picture
        frmShowFullScreen.Show
    End If
    
End Sub

Private Sub cmdHide_Click()
    mnuIconMenuHide.Enabled = False
    mnuIconMenuShow.Enabled = True
    frmMain.Hide
End Sub

Private Sub cmdReset_Click()
    chkWholeScreen.Value = vbChecked
    txtSSNum.Text = "0"
    txtFileName.Text = "screenshot"
    txtSaveFiles.Text = "" & App.Path
End Sub

Private Sub imgSS_DblClick()
    Call cmdFullPreview_Click
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFileGetScreen_Click()
    Call cmdGetScreen_Click
End Sub

Private Sub mnuFileHide_Click()
    Call cmdHide_Click
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuIconMenuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuIconMenuGetScreen_Click()
    Call cmdGetScreen_Click
End Sub

Private Sub mnuIconMenuHide_Click()
    mnuIconMenuHide.Enabled = False
    mnuIconMenuShow.Enabled = True
    frmMain.Hide
End Sub

Private Sub mnuIconMenuShow_Click()
    mnuIconMenuHide.Enabled = True
    mnuIconMenuShow.Enabled = False
    frmMain.Show
End Sub

Private Sub mnuOptionsCapture_Click()
    
    mnuOptionsCapture.Checked = Not mnuOptionsCapture.Checked
    
    If mnuOptionsCapture.Checked = True Then
       chkWholeScreen.Value = vbChecked
    Else
       chkWholeScreen.Value = vbUnchecked
    End If
    
End Sub

Private Sub mnuOptionsReset_Click()
    Call cmdReset_Click
End Sub

Private Sub mnuViewImgBox_Click()
    
    ' Check and uncheck
    mnuViewImgBox.Checked = Not mnuViewImgBox.Checked
    
    ' Show img box
    If mnuViewImgBox.Checked = True Then
       picSS.Visible = False
       imgSS.Visible = True
       mnuViewPicBox.Checked = False
    Else
       ' Keep it checked
       mnuViewImgBox.Checked = True
    End If
    
End Sub

Private Sub mnuViewPicBox_Click()
    
    ' Check and uncheck
    mnuViewPicBox.Checked = Not mnuViewPicBox.Checked
    
    ' Show pic boxes
    If mnuViewPicBox.Checked = True Then
       picSS.Visible = True
       imgSS.Visible = False
       mnuViewImgBox.Checked = False
    Else
       ' Keep it checked
       mnuViewPicBox.Checked = True
    End If
    
End Sub

Private Sub picSS_DblClick()
    Call cmdFullPreview_Click
End Sub

Private Sub Timer1_Timer()
    ' If user presses Ctrl+G then take a screenshot
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyG) Then
       Call cmdGetScreen_Click
    End If
End Sub

Private Sub TrayControl1_DblClick()
    Call cmdGetScreen_Click
End Sub

Private Sub TrayControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuIconMenu
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TrayControl1.Enabled = False
End Sub
