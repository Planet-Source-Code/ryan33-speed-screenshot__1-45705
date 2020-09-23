VERSION 5.00
Begin VB.UserControl TrayControl 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Tray Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   -15
      TabIndex        =   0
      Top             =   15
      Width           =   540
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TrayControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API Types
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'API Declares
Private Declare Function ShellNotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'API Constants
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_USER = &H400
Private Const WM_ICONNOTIFY = WM_USER + 100
Private Const ID_TASKBARICON = 100
Private Const WM_MOUSEMOVE = &H200

'Module level variables
Dim lHwnd As Long

'Default Property Values:
Const m_def_ToolTipText = ""

'Property Variables:
Dim m_ToolTipText As String

'Event Declarations:
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."


Private Sub UpdateIcon(nAction As Integer)
    
    Dim nid As NOTIFYICONDATA
        
    'Update tray icon data
    nid.cbSize = LenB(nid)
    nid.hwnd = lHwnd
    nid.uID = ID_TASKBARICON
    nid.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
    nid.uCallbackMessage = WM_MOUSEMOVE
    If Not nAction = NIM_DELETE Then
        nid.hIcon = UserControl.Extender.Parent.Icon
        nid.szTip = m_ToolTipText & Chr$(0)
    End If
    ShellNotifyIcon nAction, nid
    
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case X
        Case 7695 'Left MouseDown
            UserControl_MouseDown vbLeftButton, 0, 0, 0
        Case 7710 'Left MouseUp
            UserControl_MouseUp vbLeftButton, 0, 0, 0
        Case 7725 'Left DoubleClick
            UserControl_DblClick
        Case 7740 'Right MouseDown
            UserControl_MouseDown vbRightButton, 0, 0, 0
        Case 7755 'Right MouseUp
            UserControl_MouseDown vbRightButton, 0, 0, 0
    End Select

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

    If Ambient.UserMode Then
        lHwnd = UserControl.hwnd
        If UserControl.Enabled Then
            UpdateIcon NIM_ADD
        End If
    End If

End Sub

Private Sub UserControl_Resize()
    UserControl.Size 32 * Screen.TwipsPerPixelX, 32 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Terminate()
    If Not lHwnd = 0 And UserControl.Enabled Then
        UpdateIcon NIM_DELETE
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
    UpdateIcon NIM_MODIFY
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    If UserControl.Enabled Then
        UpdateIcon NIM_ADD
    Else
        UpdateIcon NIM_DELETE
    End If
End Property

