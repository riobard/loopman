VERSION 5.00
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "rmoc3260.dll"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loopman 1.6"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5220
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RealAudioObjectsCtl.RealAudio Real 
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      AUTOSTART       =   0   'False
      SHUFFLE         =   0   'False
      PREFETCH        =   0   'False
      NOLABELS        =   0   'False
      LOOP            =   0   'False
      NUMLOOP         =   0
      CENTER          =   0   'False
      MAINTAINASPECT  =   0   'False
      BACKGROUNDCOLOR =   "#000000"
   End
   Begin VB.Label LblTimeSS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label cmdRepeat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdReset 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdBegin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdPlayPause 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cmdOpen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Position 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   -30
      Top             =   240
      Width           =   30
   End
   Begin VB.Shape LoopArea 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   0
      Top             =   240
      Width           =   1545
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBegin_Click()
    DoBegin
End Sub

Private Sub cmdEnd_Click()
    DoEnd
End Sub

Private Sub cmdOpen_Click()
    DoOpen
End Sub

Private Sub cmdPlayPause_Click()
    DoPlayPause
End Sub

Private Sub cmdRepeat_Click()
    DoRepeat
End Sub

Private Sub cmdReset_Click()
    DoReset
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Real.GetPlayState = 3 Or Real.GetPlayState = 5 Then
        If x > lBegin / Real.GetLength * frmMain.ScaleWidth And x < lEnd / Real.GetLength * frmMain.ScaleWidth Then
            Real.SetPosition (x / frmMain.ScaleWidth * Real.GetLength)
        End If
    End If
End Sub
Private Sub Form_Load()
    AlwaysOnTop = SetWindowPos(frmMain.hwnd, -1, 0, 0, 0, 0, 3)
    bRepeat = False
    'Set Hotkeys
    SetHotkey KEYRESET, vbKeyEscape, "Add"
    SetHotkey KEYPLAYPAUSE, vbKeyF1, "Add"
    SetHotkey KEYBEGIN, vbKeyF2, "Add"
    SetHotkey KEYEND, vbKeyF3, "Add"
    SetHotkey KEYREPEAT, vbKeyF9, "Add"
    SetHotkey KEYBACKWARD, "Ctrl, 37", "Add"
    SetHotkey KEYFORWARD, "Ctrl, 39", "Add"
    SetHotkey KEYGOTOBEGIN, vbKeyF4, "Add"
    SetHotkey KEYBACKWARD5S, vbKeyF7, "Add"
    SetHotkey KEYFORWARD5S, vbKeyF8, "Add"
    ''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetHotkey KEYRESET, "", "Del"
SetHotkey KEYPLAYPAUSE, "", "Del"
SetHotkey KEYBEGIN, "", "Del"
SetHotkey KEYEND, "", "Del"
SetHotkey KEYREPEAT, "", "Del"
SetHotkey KEYBACKWARD, "", "Del"
SetHotkey KEYFORWARD, "", "Del"
SetHotkey KEYGOTOBEGIN, "", "Del"
SetHotkey KEYBACKWARD5S, "", "Del"
SetHotkey KEYFORWARD5S, "", "Del"
End Sub

Private Sub real_OnClipOpened(ByVal shortClipName As String, ByVal url As String)
    lBegin = 0
    lEnd = Real.GetLength
End Sub

Private Sub real_OnPositionChange(ByVal lPos As Long, ByVal lLen As Long)
    If (lPos < lBegin Or lPos + 500 >= lEnd) Then
        Real.SetPosition (lBegin)
    End If
    Position.Move lPos / lLen * frmMain.ScaleWidth
    LblTimeSS.Caption = Int(Real.GetPosition / 1000) Mod 60
End Sub

Private Sub real_OnStateChange(ByVal lOldState As Long, ByVal lNewState As Long)
    If lNewState = 3 Or lNewState = 5 Then
        cmdPlayPause.Caption = ";"
    Else
        cmdPlayPause.Caption = "4"
    End If
End Sub
