VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Draw Focus Rect"
      Height          =   495
      Left            =   1238
      TabIndex        =   1
      Top             =   360
      Width           =   2205
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   570
      Left            =   533
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1005
      _Version        =   393216
      Appearance      =   1
      Max             =   500
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Click into ProgressBar and move the mouse..."
      Height          =   195
      Left            =   735
      TabIndex        =   3
      Top             =   2415
      Width           =   3210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   195
      Left            =   1268
      TabIndex        =   2
      Top             =   1485
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change the bar color of ProgressBar
Private Const WM_USER = &H400
Private Const CCM_FIRST       As Long = &H2000&
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'used with fnWeight
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
'used with SetBkMode
Const OPAQUE = 2
Const TRANSPARENT = 1

Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_CHARSTREAM = 4 ' Character-stream, PLP
Private Const DT_DISPFILE = 6 ' Display-file
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_INTERNAL = &H1000
Private Const DT_LEFT = &H0
Private Const DT_METAFILE = 5 ' Metafile, VDM
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_PLOTTER = 0 ' Vector plotter
Private Const DT_RASCAMERA = 3 ' Raster camera
Private Const DT_RASDISPLAY = 1 ' Raster display
Private Const DT_RASPRINTER = 2 ' Raster printer
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_OPTIONS = (DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_NOCLIP)

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Sub Command1_Click()
Dim lnghDC As Long
Dim lngX As Long
Dim lngY As Long
Dim R As RECT

    lnghDC = GetDC(ProgressBar1.hWnd)
    
    lngX = (ProgressBar1.Width \ Screen.TwipsPerPixelX) - 5
    lngY = (ProgressBar1.Height \ Screen.TwipsPerPixelY) - 5
    SetRect R, 2, 2, lngX, lngY
    DrawFocusRect lnghDC, R
    

End Sub


Private Sub ProgressBar1_Click()
    'Label1.Caption = "Valor: " & ProgressBar1.Value
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ProgressBar1_MouseMove Button, Shift, x, y
    End If
End Sub

Private Sub ProgressBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngPercent As Long
Dim lngValue As Long
Dim strText As String

Dim lnghDC As Long
Dim R As RECT
Dim lngX As Long
Dim lngY As Long
Dim lngFont As Long

    If Button = vbLeftButton Then
        With ProgressBar1
            lngPercent = (x * 100) \ .Width
            If lngPercent > 100 Then
                lngPercent = 100
            ElseIf lngPercent < 0 Then
                lngPercent = 0
            End If
            
            lngValue = (.Max * lngPercent) \ 100
            If lngValue <> .Value Then
                If lngValue > .Max Then
                    .Value = .Max
                ElseIf lngValue < .Min Then
                    .Value = .Min
                Else
                    .Value = lngValue
                End If
            
                'This work because the Min and Max of the ProgressBar are 0 to 255 ...
                sl_SetProgressBarColour ProgressBar1.hWnd, RGB(.Value, Abs(255 - .Value), 0)
                .Refresh
            End If
            
            
            Label1.Caption = "Value: " & ProgressBar1.Value
            
            strText = CStr(lngPercent) & " %"
            
            lnghDC = GetDC(.hWnd)
            lngX = (.Width \ Screen.TwipsPerPixelX) - 5  'transforma em pixel
            lngY = (.Height \ Screen.TwipsPerPixelY) - 5 'transforma em pixel
            
            
            lngFont = CreateMyFont(lnghDC, 9, 0, "Tahoma")
            SelectObject lnghDC, lngFont
            SetBkMode lnghDC, TRANSPARENT
            'TextOut lnghDC, 60, 6, strTexto, Len(strTexto)
            

            SetRect R, 3, 3, lngX + 1, lngY + 1
            SetTextColor lnghDC, &H0
            DrawText lnghDC, strText, Len(strText), R, DT_OPTIONS
            
            SetRect R, 2, 2, lngX, lngY
            SetTextColor lnghDC, &HC0FFFF
            DrawText lnghDC, strText, Len(strText), R, DT_OPTIONS
            
            DeleteObject lngFont
            
        End With
    End If
End Sub

Private Function CreateMyFont(hdc As Long, nSize As Integer, nDegrees As Long, FontName As String) As Long
    'Create a specified font
    CreateMyFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(hdc, LOGPIXELSY), 72), 0, nDegrees * 10, 0, FW_BOLD, False, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, FontName)
End Function

Private Sub sl_SetProgressBarColour(hwndProgBar As Long, ByVal clrref As Long)
   Call SendMessage(hwndProgBar, PBM_SETBARCOLOR, 0&, ByVal clrref)
End Sub

