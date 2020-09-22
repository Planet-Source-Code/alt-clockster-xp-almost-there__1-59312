VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     cLogo
' Author:   Steve McMahon
' Date:     24 January 1999
'**************************************************************************************************
' Shows how to draw a gradient then draw vertical text onto it.
'**************************************************************************************************
' Visit vbAccelerator - advanced, free VB source code.
'     http://vbaccelerator.com
'
' Modified January, February, 2005.  Stripped out what I didn't want or need and formatted it
' to my liking.....
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' cLogo Constants
'**************************************************************************************************
Private Const FW_NORMAL = 400
' Remaining declared in mdlDeclates/Winsubhook2.tlb

'**************************************************************************************************
' cLogo Enums/Structs
'**************************************************************************************************
' None or declared in mdlDeclares/Winsubhook2.tlb

'**************************************************************************************************
' cLogo Win32 API
'**************************************************************************************************
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, _
     ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

'**************************************************************************************************
' cLogo Property Variables
'**************************************************************************************************
Private m_picCanvas As PictureBox
Private m_sCaption As String
Private m_bRGBStart(1 To 3) As Integer
Private m_oStartColor As OLE_COLOR
Private m_bRGBEnd(1 To 3) As Integer
Private m_oEndColor As OLE_COLOR

'**************************************************************************************************
' cLogo Properties
'**************************************************************************************************
Public Property Get Caption() As String
     Caption = m_sCaption
End Property ' Get Caption

Public Property Let Caption(ByVal sCaption As String)
     m_sCaption = sCaption
End Property ' Let Caption

Public Property Let DrawingObject(ByRef picThis As PictureBox)
     Set m_picCanvas = picThis
End Property ' DrawingObject

Public Property Get EndColor() As OLE_COLOR
     EndColor = m_oEndColor
End Property ' Get EndColor

Public Property Let EndColor(ByVal oColor As OLE_COLOR)
     Dim lColor As Long
     If (m_oEndColor <> oColor) Then
          m_oEndColor = oColor
          OleTranslateColor oColor, 0, lColor
          m_bRGBEnd(1) = lColor And &HFF&
          m_bRGBEnd(2) = ((lColor And &HFF00&) \ &H100)
          m_bRGBEnd(3) = ((lColor And &HFF0000) \ &H10000)
          If Not (m_picCanvas Is Nothing) Then Draw
     End If
End Property ' Let EndColor

Public Property Get StartColor() As OLE_COLOR
     StartColor = m_oStartColor
End Property ' Get StartColor

Public Property Let StartColor(ByVal oColor As OLE_COLOR)
Dim lColor As Long
    If (m_oStartColor <> oColor) Then
        m_oStartColor = oColor
        OleTranslateColor oColor, 0, lColor
        m_bRGBStart(1) = lColor And &HFF&
        m_bRGBStart(2) = ((lColor And &HFF00&) \ &H100)
        m_bRGBStart(3) = ((lColor And &HFF0000) \ &H10000)
        If Not (m_picCanvas Is Nothing) Then Draw
    End If
End Property ' Let StartColor

'**************************************************************************************************
' cLogo Public Methods/Subs
'**************************************************************************************************
Public Sub Draw()
     Dim lHeight As Long
     Dim lWidth As Long
     Dim lYStep As Long
     Dim lY As Long
     Dim bRGB(1 To 3) As Integer
     Dim tLF As LOGFONT
     Dim hFnt As Long
     Dim hFntOld As Long
     Dim lR As Long
     Dim rct As tRECT
     Dim hBr As Long
     Dim hDC As Long
     Dim dR(1 To 3) As Double
     hDC = m_picCanvas.hDC
     lHeight = m_picCanvas.Height '\ Screen.TwipsPerPixelY
     rct.Right = m_picCanvas.Width '\ Screen.TwipsPerPixelY
     ' Set a graduation of 255 pixels:
     lYStep = lHeight \ 255
     If (lYStep = 0) Then lYStep = 1
     rct.Bottom = lHeight
     bRGB(1) = m_bRGBStart(1)
     bRGB(2) = m_bRGBStart(2)
     bRGB(3) = m_bRGBStart(3)
     dR(1) = m_bRGBEnd(1) - m_bRGBStart(1)
     dR(2) = m_bRGBEnd(2) - m_bRGBStart(2)
     dR(3) = m_bRGBEnd(3) - m_bRGBStart(3)
     For lY = lHeight To 0 Step -lYStep
          ' Draw bar:
          rct.Top = rct.Bottom - lYStep
          hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
          FillRect hDC, rct, hBr
          DeleteObject hBr
          rct.Bottom = rct.Top
          ' Adjust colour:
          bRGB(1) = m_bRGBStart(1) + dR(1) * (lHeight - lY) / lHeight
          bRGB(2) = m_bRGBStart(2) + dR(2) * (lHeight - lY) / lHeight
          bRGB(3) = m_bRGBStart(3) + dR(3) * (lHeight - lY) / lHeight
     Next
     ' create the border
     lHeight = lHeight - 2
     pOLEFontToLogFont m_picCanvas.Font, hDC, tLF
     tLF.lfEscapement = 900
     hFnt = CreateFontIndirect(tLF)
     If (hFnt <> 0) Then
          hFntOld = SelectObject(hDC, hFnt)
          lR = TextOut(hDC, 0, lHeight, m_sCaption, Len(m_sCaption))
          SelectObject hDC, hFntOld
          DeleteObject hFnt
     End If
     m_picCanvas.Refresh
End Sub ' Draw

'**************************************************************************************************
' cLogo Private Methods/Subs
'**************************************************************************************************
Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
     Dim sFont As String
     Dim iChar As Integer
     ' Convert an OLE StdFont to a LOGFONT structure:
     With tLF
          sFont = fntThis.Name
          ' There is a quicker way involving StrConv and CopyMemory, but
          ' this is simpler!:
          For iChar = 1 To Len(sFont)
               .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
          Next iChar
          ' Based on the Win32SDK documentation:
          .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
          .lfItalic = fntThis.Italic
          If (fntThis.Bold) Then
               .lfWeight = FW_BOLD
          Else
               .lfWeight = FW_NORMAL
          End If
          .lfUnderline = fntThis.Underline
          .lfStrikeOut = fntThis.Strikethrough
     End With
End Sub ' pOLEFontToLogFont

'**************************************************************************************************
' cLogo Intrinsic Methods/Subs
'**************************************************************************************************
Private Sub Class_Initialize()
     StartColor = &H0
     EndColor = vbButtonFace
End Sub ' Class_Initialize
