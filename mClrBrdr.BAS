Attribute VB_Name = "mClrBrdr"
'****************************************************************************************
'Module:        mClrBrdr - BAS module
'Filename:      mClrBrdr.bas
'Author:        Jim Kahl
'Purpose:       Provide a colored border around the edge of a window based on the back
'               ground color of the form
'****************************************************************************************
'NOTE:          This was originally a tip I found on another web site that simply drew
'               a solid color for the border of the window, I like this better even if
'               it doesn't work the greatest for some colors and has a tendency to not
'               draw properly when resizing or repainting
'****************************************************************************************
Option Explicit

'****************************************************************************************
'API CONSTANTS
'****************************************************************************************
'PenStyle
Private Const PS_SOLID = 0
Private Const PS_INSIDEFRAME As Long = 6
'SystemMetrics
Private Const SM_CYBORDER As Long = 6
'OleTranslateColor
Private Const CLR_INVALID = 0

'****************************************************************************************
'TYPES
'****************************************************************************************
Public Type HSL
    Hue As Integer
    Saturation As Integer
    Luminance As Integer
End Type

'****************************************************************************************
'API FUNCTIONS
'****************************************************************************************
'USER32
Private Declare Function GetSystemMetrics Lib "user32.dll" ( _
                ByVal nIndex As Long) _
                As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
                ByVal hwnd As Long, _
                ByVal hDC As Long) _
                As Long
Private Declare Function CreatePen Lib "gdi32" ( _
                ByVal nPenStyle As Long, _
                ByVal nWidth As Long, _
                ByVal crColor As Long) _
                As Long
'GDI32
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
                ByVal hDC As Long, _
                ByVal hObject As Long) _
                As Long
Private Declare Function LineTo Lib "gdi32" ( _
                ByVal hDC As Long, _
                ByVal x As Long, _
                ByVal Y As Long) _
                As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" ( _
                ByVal clr As Long, _
                ByVal hpal As Long, _
                ByRef lpcolorref As Long) _
                As Long

'****************************************************************************************
'METHODS - PUBLIC
'****************************************************************************************
Public Sub DrawBorder(frmTarget As Form, ByVal Color As OLE_COLOR)
    Dim hDC As Long
    Dim hPen As Long
    Dim nLeft As Long
    Dim nRight As Long
    Dim nTop As Long
    Dim nBottom As Long
    Dim hNewPen As Long
    Dim nWid As Long
    Dim nRGB As Long
    Dim tHSL As HSL
    Dim nHilite As Long
    Dim nShadow As Long
    
    'translate
    nRGB = TranslateColor(Color)
    tHSL = RGBToHSL(nRGB)
    'use luminance to determine the lightness/darkness of the color
    'to choose which lines to use when providing shading
    Debug.Print tHSL.Luminance
    If tHSL.Luminance >= 100 Then
        'use these two lines for lighter colors
        nHilite = nRGB Or &H808080
        nShadow = nRGB And &HB0B0B0
    Else
        'use these two lines for darker colors
        nHilite = nRGB * 1.5
        nShadow = nRGB \ 1.5
    End If
    'border width
    nWid = GetSystemMetrics(SM_CYBORDER) * 5
    'get window DC
    hDC = GetWindowDC(frmTarget.hwnd)   'this is outside the form
    'create a pen
    hNewPen = CreatePen(PS_INSIDEFRAME, nWid, nHilite)
    'Initialize misc variables
    nLeft = 1
    nTop = 1
    nRight = frmTarget.Width / Screen.TwipsPerPixelX - 2
    nBottom = frmTarget.Height / Screen.TwipsPerPixelY - 2
    'select border pen
    hPen = SelectObject(hDC, hNewPen)
    'draw highlight color
    Call LineTo(hDC, nLeft, nBottom)
    Call LineTo(hDC, nLeft, nTop)
    Call LineTo(hDC, nRight, nTop)
    Call SelectObject(hDC, hPen)
    Call DeleteObject(hNewPen)
    'draw shadow color
    nWid = nWid \ 2
    hNewPen = CreatePen(PS_INSIDEFRAME, nWid, nShadow)
    hPen = SelectObject(hDC, hNewPen)
    Call LineTo(hDC, nRight, nBottom)
    Call LineTo(hDC, nLeft, nBottom)
    Call LineTo(hDC, nLeft + nWid, nBottom - nWid)
    Call LineTo(hDC, nRight - nWid, nBottom - nWid)
    Call LineTo(hDC, nRight - nWid, nTop + nWid)
    
    'select old pen
    Call SelectObject(hDC, hPen)
    Call DeleteObject(hNewPen)
    Call ReleaseDC(frmTarget.hwnd, hDC)
End Sub

Private Function TranslateColor( _
        ByVal clr As OLE_COLOR, _
        Optional hpal As Long = 0) As Long
        
    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Function RGBToHSL(ByVal RGBValue As Long) As HSL
    'NOTE: don't remember where I got this from copied from internet
    'some time ago
    ' by Paul - wpsjr1@syix.com, 20011120
    Dim R As Long
    Dim G As Long
    Dim B As Long
    Dim lMax As Long
    Dim lMin As Long
    Dim q As Single
    Dim lDifference As Long
    Static Lum(255) As Long
    Static QTab(255) As Single
    Static init As Long
    
    If init = 0 Then
        For init = 2 To 255 ' 0 and 1 are both 0
            Lum(init) = init * 100 / 255
        Next
        For init = 1 To 255
            QTab(init) = 60 / init
        Next init
    End If
    
    R = RGBValue And &HFF
    G = (RGBValue And &HFF00&) \ &H100&
    B = (RGBValue And &HFF0000) \ &H10000
    
    If R > G Then
        lMax = R
        lMin = G
    Else
        lMax = G
        lMin = R
    End If
    If B > lMax Then
        lMax = B
    ElseIf B < lMin Then
        lMin = B
    End If
    
    RGBToHSL.Luminance = Lum(lMax)
    
    lDifference = lMax - lMin
    If lDifference Then
        ' do a 65K 2D lookup table here for more speed if needed
        RGBToHSL.Saturation = (lDifference) * 100 / lMax
        q = QTab(lDifference)
        Select Case lMax
            Case R
                If B > G Then
                    RGBToHSL.Hue = q * (G - B) + 360
                Else
                    RGBToHSL.Hue = q * (G - B)
                End If
            Case G
                RGBToHSL.Hue = q * (B - R) + 120
            Case B
                RGBToHSL.Hue = q * (R - G) + 240
        End Select
    End If
End Function


