VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using Themes in VB Applications"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   4920
   End
   Begin VB.PictureBox picPreview 
      Height          =   3135
      Left            =   360
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   1
      Top             =   1800
      Width           =   5175
   End
   Begin VB.CheckBox chkEnableThemes 
      Caption         =   "Enable Themes"
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tooltip:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label lblTooltip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   510
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Canonical Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblCanonicalName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   465
   End
   Begin VB.Label lblThemeName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Theme:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblTheme 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ThemeTest created by The KPD-Team
'Copyright (c) 2001, The KPD-Team
'Visit our site at http://www.allapi.net/
'or email us at KPDTeam@allapi.net

Private Const SM_CYCAPTION = 4
Private Const SM_CYBORDER = 6
Private Const SM_CXBORDER = 5
Private Const SM_CXMENUSIZE = 54
Private Const SM_CYMENUSIZE = 55
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private S As SIZE, FullRect As RECT
Private Sub chkEnableThemes_Click()
    'Enabe or disable theming
    EnableTheming (chkEnableThemes.Value = vbChecked)
End Sub
Private Sub DrawThemedWindow(hdc As Long, DestRect As RECT)
    Dim hTheme As Long, R As RECT, CB As RECT, lTBarHeight As Long
    'Open the WINDOW theme data
    hTheme = OpenThemeData(Me.hWnd, "WINDOW")
    'Draw the Title Bar
    GetThemePartSize hTheme, hdc, WP_CAPTION, CS_ACTIVE, DestRect, TS_TRUE, S
    lTBarHeight = S.cy
    SetRect R, DestRect.Left, DestRect.Top, DestRect.Right, DestRect.Top + S.cy
    DrawThemeBackground hTheme, hdc, WP_CAPTION, CS_ACTIVE, R, ByVal 0&
    'Draw the Close button
    GetThemePartSize hTheme, hdc, WP_CLOSEBUTTON, CBS_NORMAL, DestRect, TS_TRUE, S
    SetRect CB, 0, 0, S.cx, S.cy
    SetRect R, DestRect.Right - CB.Right - 5, DestRect.Top + (lTBarHeight - S.cy) / 2, DestRect.Right - 5, DestRect.Top + CB.bottom + (lTBarHeight - S.cy) / 2
    DrawThemeBackground hTheme, hdc, WP_CLOSEBUTTON, CBS_NORMAL, R, ByVal 0&
    'Draw the Max button
    GetThemePartSize hTheme, hdc, WP_MAXBUTTON, MAXBS_NORMAL, DestRect, TS_TRUE, S
    SetRect CB, 0, 0, S.cx, S.cy
    OffsetRect R, -(CB.Right + 2), 0
    DrawThemeBackground hTheme, hdc, WP_MAXBUTTON, MAXBS_NORMAL, R, ByVal 0&
    'Draw the Min button
    GetThemePartSize hTheme, hdc, WP_MINBUTTON, MINBS_NORMAL, DestRect, TS_TRUE, S
    SetRect CB, 0, 0, S.cx, S.cy
    OffsetRect R, -(CB.Right + 2), 0
    DrawThemeBackground hTheme, hdc, WP_MINBUTTON, MINBS_NORMAL, R, ByVal 0&
    'Draw the left border
    GetThemePartSize hTheme, hdc, WP_FRAMELEFT, FS_ACTIVE, DestRect, TS_TRUE, S
    SetRect R, 0, DestRect.Top + GetSystemMetrics(SM_CYCAPTION), S.cx, DestRect.bottom
    DrawThemeBackground hTheme, hdc, WP_FRAMELEFT, FS_ACTIVE, R, ByVal 0&
    'Draw the right border
    GetThemePartSize hTheme, hdc, WP_FRAMERIGHT, FS_ACTIVE, DestRect, TS_TRUE, S
    SetRect R, DestRect.Right - S.cx, DestRect.Top + GetSystemMetrics(SM_CYCAPTION), DestRect.Right, DestRect.bottom
    DrawThemeBackground hTheme, hdc, WP_FRAMERIGHT, FS_ACTIVE, R, ByVal 0&
    'Draw the bottom border
    GetThemePartSize hTheme, hdc, WP_FRAMEBOTTOM, FS_ACTIVE, DestRect, TS_TRUE, S
    SetRect R, 0, DestRect.bottom - S.cy, DestRect.Right, DestRect.bottom
    DrawThemeBackground hTheme, hdc, WP_FRAMEBOTTOM, FS_ACTIVE, R, ByVal 0&
    'Clean up
    CloseThemeData hTheme
End Sub
Private Sub DrawButton(hdc As Long, DestRect As RECT, Caption As String)
    Dim hTheme As Long
    'Open the BUTTON theme data
    hTheme = OpenThemeData(Me.hWnd, "BUTTON")
    'Draw the button background
    DrawThemeBackground hTheme, hdc, BP_PUSHBUTTON, PBS_NORMAL, DestRect, ByVal 0&
    'Draw the caption
    DrawThemeText hTheme, hdc, BP_PUSHBUTTON, PBS_NORMAL, Caption, -1, DT_CENTER Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_SINGLELINE, 0, DestRect
    'Clean up
    CloseThemeData hTheme
End Sub
Private Sub DrawProgressBar(hdc As Long, DestRect As RECT, Value As Long)
    Dim hTheme As Long, R As RECT
    'Open the PROGRESS theme data
    hTheme = OpenThemeData(Me.hWnd, "PROGRESS")
    'Draw the progress bar background
    DrawThemeBackground hTheme, hdc, PP_BAR, 0, DestRect, ByVal 0&
    'Draw the bar
    SetRect R, DestRect.Left + 2, DestRect.Top + 2, DestRect.Left + 2 + (DestRect.Right - DestRect.Left - 4) / 100 * Value - 2, DestRect.bottom - 2
    DrawThemeBackground hTheme, hdc, PP_CHUNK, 0, R, ByVal 0&
    'Clean up
    CloseThemeData hTheme
End Sub
Private Sub Form_Load()
    'Check whether themes are supported
    If AreThemesSupported = False Then
        MsgBox "This project requires Theme support...", vbCritical
        Unload Me
        Exit Sub
    End If
    'Initialize FullRect
    SetRect FullRect, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight
    'Retrieve some info about the current theme
    lblTheme.Caption = GetCurrentTheme
    lblThemeName.Caption = GetThemeProperty(lblTheme.Caption, SZ_THDOCPROP_DISPLAYNAME)
    lblCanonicalName.Caption = GetThemeProperty(lblTheme.Caption, SZ_THDOCPROP_CANONICALNAME)
    lblAuthor.Caption = GetThemeProperty(lblTheme.Caption, SZ_THDOCPROP_AUTHOR)
    lblTooltip.Caption = GetThemeProperty(lblTheme.Caption, SZ_THDOCPROP_TOOLTIP)
    'Start the progress bar timer
    tmrProgress.Enabled = True
End Sub
Function GetCurrentTheme() As String
    Dim ZeroPos As Long
    'Create a buffer
    GetCurrentTheme = String(255, 0)
    'Get the name of the current theme
    GetCurrentThemeName GetCurrentTheme, Len(GetCurrentTheme), vbNullString, 0, vbNullString, 0
    'Strip off trailing Chr$(0)'s
    ZeroPos = InStr(1, GetCurrentTheme, Chr$(0))
    If ZeroPos > 0 Then
        GetCurrentTheme = Left$(GetCurrentTheme, ZeroPos - 1)
    End If
End Function
Function GetThemeProperty(sFile As String, sProperty As String) As String
    Dim ZeroPos As Long
    'Create a buffer
    GetThemeProperty = String(255, 0)
    'Retrieve the documentation
    GetThemeDocumentationProperty sFile, sProperty, GetThemeProperty, Len(GetThemeProperty)
    'Strip off trailing Chr$(0)'s
    ZeroPos = InStr(1, GetThemeProperty, Chr$(0))
    If ZeroPos > 0 Then
        GetThemeProperty = Left$(GetThemeProperty, ZeroPos - 1)
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)
    tmrProgress.Enabled = False
End Sub
Private Sub picPreview_Paint()
    Dim R As RECT
    DrawThemedWindow picPreview.hdc, FullRect
    SetRect R, 100, 100, 180, 125
    DrawButton picPreview.hdc, R, "Test"
End Sub
Private Sub tmrProgress_Timer()
    Static Value As Long, Add As Long
    Dim R As RECT
    If Add = 0 Then Add = 1
    Value = Value + 10 * Add
    If Value = 100 Or Value = 0 Then
        Add = -Add
    End If
    SetRect R, 200, 50, 280, 65
    DrawProgressBar picPreview.hdc, R, Value
End Sub
