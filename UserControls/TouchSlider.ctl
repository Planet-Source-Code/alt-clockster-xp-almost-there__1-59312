VERSION 5.00
Begin VB.UserControl TouchSlider 
   ClientHeight    =   120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   DrawStyle       =   5  'Transparent
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   8
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   60
   ToolboxBitmap   =   "TouchSlider.ctx":0000
End
Attribute VB_Name = "TouchSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' TouchSlider.ctl
'**************************************************************************************************
' Found this code here:  http://www.vbcode.com/asp/showsn.asp?theID=11561
' Author is listed as T-Prgrams but this code looks almost identical to the Cool Progress Bar
' submitted by Mario Flores on Planet-Source-Code.com.  You can find that submission
' here:  http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=53628&lngWId=1
'
' At any rate, I pared it down to suit my needs, formatted it to fit my style, and made many
' other modifications in porting it to this project.  I had already designed a slider control
' for my project but I just liked this one better. ;-)  Good work T-Prgrms, Mario, or whoever
' wrote this originally.
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' TouchSlider Constants
'**************************************************************************************************
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const CALLBACK_WINDOW = &H10000
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_CALCRECT As Long = &H400
Private Const MM_MIXM_CONTROL_CHANGE = &H3D1
Private Const MAXPNAMELEN = 32
Private Const MMSYSERR_NOERROR = 0
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or _
    MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or _
    MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const RGN_DIFF = 4

'**************************************************************************************************
' TouchSlider Enums
'**************************************************************************************************
Public Enum mvComponent
     [mvVolume] = 0
     [mvMute] = 1
End Enum ' mvComponent

Public Enum mvMute
     [IsNotMuted] = 0
     [IsMuted] = 1
End Enum ' mvMute

Public Enum tsProgressStyle
     [Segmented] = 0
     [Solid] = 1
End Enum ' tsProgressStyle

Public Enum eTSSound
     [UseInternalSound] = 0
     [UseExternalSound] = 1
End Enum ' eTSSound

Private Enum TRACKMOUSEEVENT_FLAGS
     TME_HOVER = &H1&
     TME_LEAVE = &H2&
     TME_QUERY = &H40000000
     TME_CANCEL = &H80000000
End Enum ' TRACKMOUSEEVENT_FLAGS

'**************************************************************************************************
' TouchSlider Structs
'**************************************************************************************************
Private Type POINTAPI
     x As Long
     y As Long
End Type ' POINTAPI

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

Private Type TRACKMOUSEEVENT_STRUCT
     cbSize As Long
     dwFlags As TRACKMOUSEEVENT_FLAGS
     hwndTrack As Long
     dwHoverTime As Long
End Type ' TRACKMOUSEEVENT_STRUCT

Private Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lMinimum As Long
     lMaximum As Long
     reserved(9) As Long
End Type ' MIXERCONTROL

Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    item As Long
    cbDetails As Long
    paDetails As Long
End Type ' MIXERCONTROLDETAILS

Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type ' MIXERCONTROLDETAILS_BOOLEAN

Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type ' MIXERCONTROLDETAILS_UNSIGNED

Private Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     dwType As Long
     dwDeviceID As Long
     wMid  As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
End Type ' MIXERLINE

Private Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type ' MIXERLINECONTROLS

'**************************************************************************************************
' TouchSlider API Declares
'**************************************************************************************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
     ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
     ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal fnStyle As Integer, _
     ByVal COLORREF As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
     ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, _
     ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, _
     ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, _
     ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, _
     ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
     lpRect As RECT) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
     ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
     ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, _
     ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, _
     ByVal nBkMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
     ByVal bRedraw As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" ( _
     lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" ( _
     lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
' mixer api
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" ( _
     ByVal hmxobj As Long, pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" ( _
     ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" ( _
     ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, _
     ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, _
     pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
' memory manipulation API
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, _
     ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, _
     struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
     
'**************************************************************************************************
' TouchSlider Events
'**************************************************************************************************
Public Event MouseEnter()
Public Event MouseLeave()
Public Event OnContextMenu()
Public Event OnMute(bValue As Boolean)
Public Event OnPositionChange(ByVal lOffsetY As Long, ByVal lHeightAdj As Long)
Public Event ValueChange(ByVal lValue As Long)

'**************************************************************************************************
' TouchSlider Default Property Constants
'**************************************************************************************************
Private Const m_def_Color = 33023
Private Const m_def_EnableSound = True
Private Const m_def_Height = 8
Private Const m_def_HeightAdjust = 0
Private Const m_def_Max = 100
Private Const m_def_Min = 0
Private Const m_def_offsetY = 10
Private Const m_def_ProgressStyle = 0
Private Const m_def_SegmentSize = 3
Private Const m_def_SegmentSpace = 1
Private Const m_def_Sound = 0

'**************************************************************************************************
' TouchSlider Module-Scoped Private Variables
'**************************************************************************************************
Implements WinSubHook2.iSubclass
Private m_sc As cSubclass
Private m_MemDC As Boolean
Private m_DimhDC As Long
Private m_IsDrag As Boolean
Private m_hBmp As Long
Private m_hBmpOld As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_bInCtrl As Boolean
Private m_bTrack As Boolean
Private m_bTrackUser32 As Boolean
Private m_hMixer As Long
Private m_mxc_vol As MIXERCONTROL
Private m_mxc_mute As MIXERCONTROL
Private fPercent As Double
Private TR As RECT
Private TBR As RECT
Private TSR As RECT
Private AT As RECT

'**************************************************************************************************
' TouchSlider Property Variables
'**************************************************************************************************
Private m_Color As OLE_COLOR
Private m_EnableSound As Boolean
Private m_hDC As Long
Private m_HeightAdjust As Long
Private m_hWnd As Long
Private m_Max As Long
Private m_Min As Long
Private m_Mute As Boolean
Private m_SliderOffsetY As Long
Private m_Sound As eTSSound
Private m_SoundPath As String
Private m_ProgressStyle  As tsProgressStyle
Private m_SegmentSize As Long
Private m_SegmentSpace As Long
Private m_Value As Long

'**************************************************************************************************
' TouchSlider Property Statements
'**************************************************************************************************
Public Property Get Color() As OLE_COLOR
     Color = m_Color
End Property ' Get Color

Public Property Let Color(ByVal New_Color As OLE_COLOR)
     ' store in property
     m_Color = pGetLongColor(New_Color)
     ' set in registry
     UpdateRegistry "TouchSliderForeColor", CStr(New_Color), CStr(m_def_Color)
     ' update display
     pDrawProgressBar
End Property ' Let Color

Public Property Get DefaultHeight() As Long
     DefaultHeight = m_def_Height
End Property ' Get DefaultHeight

Public Property Get EnableSound() As Boolean
     EnableSound = m_EnableSound
End Property ' Get EnableSound

Public Property Let EnableSound(New_EnableSound As Boolean)
     ' store in property
     m_EnableSound = New_EnableSound
     ' set registry
     UpdateRegistry "TouchSliderEnableSound", CStr(New_EnableSound), CStr(m_def_EnableSound)
End Property ' Let EnableSound

Public Property Get hDC() As Long
     hDC = m_hDC
End Property ' Get hDC

Public Property Let hDC(ByVal New_hDC As Long)
     ' dimension the memory DC
     m_hDC = pDimhDC(UserControl.ScaleWidth, UserControl.ScaleHeight)
     ' if we don't have a memory DC
     If m_hDC = 0 Then
          ' use the uc's DC
          m_hDC = UserControl.hDC   'On Fail...Do it Normally
     Else
          ' we have memory DC
          m_MemDC = True
     End If
End Property ' Let hDC

Public Property Get HeightAdjust() As Long
     ' return property
     HeightAdjust = m_HeightAdjust
End Property ' Get Height

Public Property Let HeightAdjust(New_HeightAdjust As Long)
     ' store property
     m_HeightAdjust = New_HeightAdjust
     ' set registry
     UpdateRegistry "TouchSliderHeightAdjust", CStr(New_HeightAdjust), CStr(m_def_HeightAdjust)
     ' raise the event
     RaiseEvent OnPositionChange(m_SliderOffsetY, m_HeightAdjust)
End Property ' Let HeightAdjust

Public Property Get hMixer() As Long
     hMixer = m_hMixer
End Property ' Get hMixer

Public Property Let hMixer(New_hMixer As Long)
     ' Do nothing
End Property ' Let hMixer

Public Property Get hwnd() As Long
     hwnd = m_hWnd
End Property ' Get hWnd

Public Property Let hwnd(ByVal New_hWnd As Long)
     m_hWnd = New_hWnd
End Property ' Let hWnd

Public Property Get Mute() As Boolean
     Mute = GetMasterMute(m_mxc_mute)
End Property ' Get Mute

Public Property Let Mute(New_Mute As Boolean)
     ' Did the value change
     If m_Mute <> New_Mute Then
          ' set the new value
          SetMasterMute New_Mute, m_mxc_mute
          ' update property
          m_Mute = New_Mute
          ' Raise Event
          RaiseEvent OnMute(m_Mute)
     End If
     ' always draw
     pDrawProgressBar
End Property ' Let Mute

Public Property Get ProgressStyle() As tsProgressStyle
     ProgressStyle = m_ProgressStyle
End Property ' Get ProgressStyle

Public Property Let ProgressStyle(ByVal New_ProgressStyle As tsProgressStyle)
     ' store in property
     m_ProgressStyle = New_ProgressStyle
     ' set registry
     UpdateRegistry "TouchSliderProgressStyle", CStr(New_ProgressStyle), _
          CStr(m_def_ProgressStyle)
     ' update display
     pDrawProgressBar
End Property ' Let ProgressStyle

Public Property Get SegmentSize() As Long
     SegmentSize = m_SegmentSize
End Property ' Get SegmentSize

Public Property Let SegmentSize(New_SegmentSize As Long)
     ' store in property
     m_SegmentSize = New_SegmentSize
     ' set registry
     UpdateRegistry "TouchSliderSegmentSize", CStr(New_SegmentSize), _
          CStr(m_def_SegmentSize)
     ' Update progress bar
     pDrawProgressBar
End Property ' Let SegmentSize

Public Property Get SegmentSpace() As Long
     SegmentSpace = m_SegmentSpace
End Property ' Get SegmentSpace

Public Property Let SegmentSpace(New_SegmentSpace As Long)
     ' store in property
     m_SegmentSpace = New_SegmentSpace
     ' set registry
     UpdateRegistry "TouchSliderSegmentSpace", CStr(New_SegmentSpace), _
          CStr(m_def_SegmentSpace)
     ' Update progress bar
     pDrawProgressBar
End Property ' Let SegmentSpace

Public Property Get SliderOffsetY() As Long
     SliderOffsetY = m_SliderOffsetY
End Property ' Get SliderOffsetY

Public Property Let SliderOffsetY(New_SliderOffsetY As Long)
     ' store in property
     m_SliderOffsetY = New_SliderOffsetY
     ' store in registry
     UpdateRegistry "TouchSliderOffsetY", CStr(New_SliderOffsetY), CStr(m_def_offsetY)
     ' raise the event
     RaiseEvent OnPositionChange(m_SliderOffsetY, m_HeightAdjust)
End Property ' Let SliderOffsetY

Public Property Get Sound() As eTSSound
     Sound = m_Sound
End Property ' Get Sound

Public Property Let Sound(New_Sound As eTSSound)
     ' store in property
     m_Sound = New_Sound
     ' store in registry
     UpdateRegistry "TouchSliderSound", CStr(New_Sound), CStr(m_def_Sound)
End Property ' Let Sound

Public Property Get SoundPath() As String
     SoundPath = m_SoundPath
End Property ' Get SoundPath

Public Property Let SoundPath(New_SoundPath As String)
     ' store in property
     m_SoundPath = New_SoundPath
     ' store in registry
     UpdateRegistry "TouchSliderSoundPath", CStr(New_SoundPath), ""
End Property ' Let SoundPath

Public Property Get Value() As Long
     Dim lVal As Long
     ' Get raw value
     lVal = GetMasterVolume(m_mxc_vol)
     ' Return property value
     Value = ((m_Value / 100) * m_Max) / IIf(m_Min > 0, m_Min, 1)
End Property ' Get Value

Public Property Let Value(ByVal New_Value As Long)
     Dim lValue As Long
     ' If New_Value is out of range exit without changes
     If (New_Value < m_def_Min Or New_Value > m_def_Max) Then Exit Property
     ' has the value changed?
     If New_Value <> m_Value Then
          ' calculate new value
          If New_Value > 0 Then lValue = m_mxc_vol.lMaximum * (New_Value / 100)
          ' set the new value
          SetMasterVolume lValue, m_mxc_vol
          ' Set property variable
          m_Value = New_Value
          ' Broadcast change
          PropertyChanged "Value"
     Else
          ' do nothinig
     End If
     pDrawProgressBar
End Property ' Let Value

'**************************************************************************************************
' TouchSlider Private Methods
'**************************************************************************************************
Private Sub pCalcBarSize()
     TR.Left = TR.Left + 3
     LSet TBR = TR
     fPercent = m_Value / 98
     If fPercent < 0# Then fPercent = 0#
     TBR.Right = TR.Left + (TR.Right - TR.Left) * fPercent
     TBR.Right = TBR.Right - ((TBR.Right - TBR.Left) Mod (m_SegmentSize + m_SegmentSpace))
     If TBR.Right < TR.Left Then TBR.Right = TR.Left
End Sub ' pCalcBarSize

Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
     Dim lhDCC As Long
     pDestroy
     lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
     If Not (lhDCC = 0) Then
          m_DimhDC = CreateCompatibleDC(lhDCC)
          If Not (m_DimhDC = 0) Then
               m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
               If Not (m_hBmp = 0) Then
                    m_hBmpOld = SelectObject(m_DimhDC, m_hBmp)
                    If Not (m_hBmpOld = 0) Then
                         m_lWidth = Width
                         m_lHeight = Height
                         DeleteDC lhDCC
                         Exit Sub
                    End If
               End If
          End If
          DeleteDC lhDCC
          pDestroy
     End If
End Sub ' pCreate

Private Sub pDestroy()
     If Not m_hBmpOld = 0 Then
          SelectObject m_DimhDC, m_hBmpOld
          m_hBmpOld = 0
     End If
     If Not m_hBmp = 0 Then
          DeleteObject m_hBmp
          m_hBmp = 0
     End If
     If Not m_DimhDC = 0 Then
          DeleteDC m_DimhDC
          m_DimhDC = 0
     End If
     m_lWidth = 0
     m_lHeight = 0
End Sub ' pDestroy

Private Function pDimhDC(Width As Long, Height As Long) As Long
     If m_DimhDC = 0 Then
          If (Width > 0) And (Height > 0) Then pCreate Width, Height
     Else
          If Width > m_lWidth Or Height > m_lHeight Then pCreate Width, Height
     End If
     pDimhDC = m_DimhDC
End Function ' pDimhDC

Public Sub pDraw(ByVal hDC As Long, Optional ByVal xSrc As Long = 0, _
     Optional ByVal ySrc As Long = 0, Optional ByVal WidthSrc As Long = 0, _
     Optional ByVal HeightSrc As Long = 0, Optional ByVal xDst As Long = 0, _
     Optional ByVal yDst As Long = 0)
     If WidthSrc <= 0 Then WidthSrc = m_lWidth
     If HeightSrc <= 0 Then HeightSrc = m_lHeight
     BitBlt hDC, xDst, yDst, WidthSrc, HeightSrc, m_DimhDC, xSrc, ySrc, vbSrcCopy
End Sub ' pDraw

Private Sub pDrawBar()
     Dim TempRect As RECT
     Dim ITemp As Long
     Dim lColor As Long
     If m_Mute Then
          lColor = m_Color And &H808080
     Else
          lColor = m_Color
     End If
     TempRect.Left = 4
     TempRect.Right = TBR.Right ' IIf(TBR.Right + 4 > TR.Right, TBR.Right - 4, TBR.Right)
     If TempRect.Right < TempRect.Left Then
          TempRect.Right = TempRect.Left + 1
     ElseIf TempRect.Right >= ScaleWidth - 3 Then
          TempRect.Right = ScaleWidth - 3
     End If
     TempRect.Top = 8
     TempRect.Bottom = TR.Bottom - 8
     pDrawGradient pShiftColorXP(lColor, 150), lColor, 4, 3, TempRect.Right, 6, m_hDC
     pDrawFillRectangle TempRect, lColor, m_hDC
     pDrawGradient lColor, pShiftColorXP(lColor, 150), 4, TempRect.Bottom - 2, _
          TempRect.Right, 6, m_hDC
End Sub ' pDrawBar

Private Sub pDrawBorder()
     Dim RTemp As RECT
     Dim hRgn1 As Long
     Dim hRgn2 As Long
     Dim hRgnCmb As Long
     TR.Left = TR.Left - 3
     Let RTemp = TR
     pDrawLine 3, 1, TR.Right - 2, 1, m_hDC, &HBEBEBE
     pDrawLine 2, TR.Bottom - 2, TR.Right - 2, TR.Bottom - 2, m_hDC, &HEFEFEF
     pDrawLine 1, 2, 1, TR.Bottom - 2, m_hDC, &HBEBEBE
     pDrawLine 2, 2, 2, TR.Bottom - 2, m_hDC, &HEFEFEF
     pDrawLine 2, 2, TR.Right - 2, 2, m_hDC, &HEFEFEF
     pDrawLine TR.Right - 2, 2, TR.Right - 2, TR.Bottom - 2, m_hDC, &HEFEFEF
     pDrawRectangle TR, pGetLongColor(&H686868), m_hDC
     ' top-left corner
     Call SetPixelV(m_hDC, 0, 1, pGetLongColor(&HA6ABAC))
     Call SetPixelV(m_hDC, 0, 2, pGetLongColor(&H7D7E7F))
     Call SetPixelV(m_hDC, 1, 0, pGetLongColor(&HA7ABAC))
     Call SetPixelV(m_hDC, 1, 1, pGetLongColor(&H777777))
     Call SetPixelV(m_hDC, 2, 0, pGetLongColor(&H7D7E7F))
     Call SetPixelV(m_hDC, 2, 2, pGetLongColor(&HBEBEBE))
     ' bottom-left corner
     Call SetPixelV(m_hDC, 1, TR.Bottom - 1, pGetLongColor(&HA6ABAC))
     Call SetPixelV(m_hDC, 2, TR.Bottom - 1, pGetLongColor(&H7D7E7F))
     Call SetPixelV(m_hDC, 0, TR.Bottom - 3, pGetLongColor(&H7D7E7F))
     Call SetPixelV(m_hDC, 0, TR.Bottom - 2, pGetLongColor(&HA7ABAC))
     Call SetPixelV(m_hDC, 1, TR.Bottom - 2, pGetLongColor(&H777777))
     ' top-right corner
     Call SetPixelV(m_hDC, TR.Right - 2, 1, pGetLongColor(&H686868))
     Call SetPixelV(m_hDC, TR.Right - 2, 0, pGetLongColor(&HA7ABAC))
     Call SetPixelV(m_hDC, TR.Right - 1, 1, pGetLongColor(&HA7ABAC))
     Call SetPixelV(m_hDC, TR.Right - 1, 2, pGetLongColor(&H7D7E7F))
     ' bottom-right corner
     Call SetPixelV(m_hDC, TR.Right - 1, TR.Bottom - 3, pGetLongColor(&H7D7E7F))
     Call SetPixelV(m_hDC, TR.Right - 1, TR.Bottom - 2, pGetLongColor(&HA7ABAC))
     Call SetPixelV(m_hDC, TR.Right - 2, TR.Bottom - 1, pGetLongColor(&HA7ABAC))
     ' round the border using regions
     hRgnCmb = CreateRectRgn(0, 0, TR.Right, TR.Bottom)
     hRgn2 = CreateRectRgn(0, 0, 0, 0)
     hRgn1 = CreateRectRgn(0, 0, 1, 1)
     CombineRgn hRgn2, hRgnCmb, hRgn1, RGN_DIFF
     DeleteObject hRgn1
       hRgn1 = CreateRectRgn(0, TR.Bottom, 1, TR.Bottom - 1)
     CombineRgn hRgnCmb, hRgn2, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     hRgn1 = CreateRectRgn(TR.Right - 1, 0, TR.Right, 1)
     CombineRgn hRgn2, hRgnCmb, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     hRgn1 = CreateRectRgn(TR.Right - 1, TR.Bottom, TR.Right, TR.Bottom - 1)
     CombineRgn hRgnCmb, hRgn2, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     DeleteObject hRgn2
     SetWindowRgn m_hWnd, hRgnCmb, True
     DeleteObject hRgnCmb
End Sub ' pDrawBorder

Private Sub pDrawFillRectangle(ByRef hRect As RECT, ByVal lColor As Long, ByVal lHDC As Long)
     Dim hBrush As Long
     hBrush = CreateSolidBrush(pGetLongColor(lColor))
     FillRect lHDC, hRect, hBrush
     DeleteObject hBrush
End Sub ' pDrawFillRectangle

Private Sub pDrawGradient(lEndColor As Long, lStartcolor As Long, ByVal x As Long, _
     ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal hDC As Long, _
     Optional bH As Boolean)
     On Error Resume Next
     ' Draw a Vertical Gradient in the current HDC
     Dim sR As Single
     Dim sG As Single
     Dim sB As Single
     Dim eR As Single
     Dim eG As Single
     Dim eB As Single
     Dim ni As Long
     lEndColor = pGetLongColor(lEndColor)
     lStartcolor = pGetLongColor(lStartcolor)
     sR = (lStartcolor And &HFF)
     sG = (lStartcolor \ &H100) And &HFF
     sB = (lStartcolor And &HFF0000) / &H10000
     eR = (lEndColor And &HFF)
     eG = (lEndColor \ &H100) And &HFF
     eB = (lEndColor And &HFF0000) / &H10000
     sR = (sR - eR) / IIf(bH, x2, y2)
     sG = (sG - eG) / IIf(bH, x2, y2)
     sB = (sB - eB) / IIf(bH, x2, y2)
     For ni = 0 To IIf(bH, x2, y2)
          If bH Then
               pDrawLine x + ni, y, x + ni, y2, hDC, RGB(eR + (ni * sR), eG + (ni * sG), _
                    eB + (ni * sB))
          Else
               pDrawLine x, y + ni, x2, y + ni, hDC, RGB(eR + (ni * sR), eG + (ni * sG), _
                    eB + (ni * sB))
          End If
     Next ni
End Sub ' pDrawGradient

Private Sub pDrawLine(ByVal x As Long, ByVal y As Long, ByVal Width As Long, _
     ByVal Height As Long, ByVal lHDC As Long, ByVal Color As Long)
     Dim hPen As Long
     Dim hPenOld As Long
     Dim Outline As Long
     Dim pt As POINTAPI
     hPen = CreatePen(0, 1, pGetLongColor(Color))
     hPenOld = SelectObject(lHDC, hPen)
     MoveToEx lHDC, x, y, pt
     LineTo lHDC, Width, Height
     SelectObject lHDC, hPenOld
     DeleteObject hPen
     DeleteObject hPenOld
End Sub ' pDrawLine

Private Sub pDrawProgressBar()
     If m_Value > 100 Then m_Value = 100
     ' Reference = Control Client Area
     GetClientRect m_hWnd, TR
     ' Fill the background
     pDrawFillRectangle TR, vbWhite, m_hDC
     ' Calculate Progress and Percent Values
     pCalcBarSize
      ' Draw bar
     If m_Value > 0 Then
          pDrawBar
          ' Draw Segment Spacing (This Will Generate the Blocks Effect)
          If m_ProgressStyle = 0 Then pDrawSegmentSpace
     End If
     ' Draw The XP Look Border
     pDrawBorder
     ' draw from DC
     If m_MemDC Then
          With UserControl
               pDraw .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .ScaleLeft, .ScaleTop
          End With
     End If
     MaskPicture = Image
End Sub ' pDrawProgressBar

Private Sub pDrawRectangle(ByRef BRect As RECT, ByVal Color As Long, ByVal hDC As Long)
     Dim hBrush As Long
     hBrush = CreateSolidBrush(Color)
     FrameRect hDC, BRect, hBrush
     DeleteObject hBrush
End Sub ' pDrawRectangle

Private Sub pDrawSegmentSpace()
     Dim i As Long
     Dim hBr As Long
     ' get brush handle
     hBr = CreateSolidBrush(vbWhite)
     ' set rectangle
     LSet TSR = TR
     For i = TBR.Left + m_SegmentSize To TBR.Right Step m_SegmentSize + m_SegmentSpace
          TSR.Left = i + 1
          TSR.Right = i + 1 + m_SegmentSpace
          FillRect m_hDC, TSR, hBr
     Next i
     ' delete the brush
     DeleteObject hBr
End Sub ' pDrawSegmentSpace

Private Function pGetLongColor(Color As Long) As Long
     If (Color And &H80000000) Then
          pGetLongColor = GetSysColor(Color And &H7FFFFFFF)
     Else
          pGetLongColor = Color
     End If
End Function ' pGetLongColor

Private Function pShiftColorXP(ByVal MyColor As Long, ByVal Base As Long) As Long
     Dim R As Long
     Dim G As Long
     Dim B As Long
     Dim Delta As Long
     R = (MyColor And &HFF)
     G = ((MyColor \ &H100) Mod &H100)
     B = ((MyColor \ &H10000) Mod &H100)
     Delta = &HFF - Base
     B = Base + B * Delta \ &HFF
     G = Base + G * Delta \ &HFF
     R = Base + R * Delta \ &HFF
     If R > 255 Then R = 255
     If G > 255 Then G = 255
     If B > 255 Then B = 255
     pShiftColorXP = R + 256& * G + 65536 * B
End Function ' pShiftColorXP

Private Sub RoundCorners(rc As RECT, ByVal m_hWnd As Long)
     Dim hRgn1 As Long
     Dim hRgn2 As Long
     Dim hRgnCmb As Long
     hRgnCmb = CreateRectRgn(0, 0, rc.Right, rc.Bottom)
     hRgn2 = CreateRectRgn(0, 0, 0, 0)
     hRgn1 = CreateRectRgn(0, 0, 1, 1)
     CombineRgn hRgn2, hRgnCmb, hRgn1, RGN_DIFF
     DeleteObject hRgn1
       hRgn1 = CreateRectRgn(0, rc.Bottom, 1, rc.Bottom - 1)
     CombineRgn hRgnCmb, hRgn2, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     hRgn1 = CreateRectRgn(rc.Right - 1, 0, rc.Right, 1)
     CombineRgn hRgn2, hRgnCmb, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     hRgn1 = CreateRectRgn(rc.Right - 1, rc.Bottom, rc.Right, rc.Bottom - 1)
     CombineRgn hRgnCmb, hRgn2, hRgn1, RGN_DIFF
     DeleteObject hRgn1
     DeleteObject hRgn2
     SetWindowRgn m_hWnd, hRgnCmb, True
     DeleteObject hRgnCmb
End Sub ' RoundCorners

'Track the mouse leaving the indicated window
Private Sub pTrackMouseLeave(ByVal lng_hWnd As Long)
     Dim tme As TRACKMOUSEEVENT_STRUCT
     If m_bTrack Then
          With tme
               .cbSize = Len(tme)
               .dwFlags = TME_LEAVE
               .hwndTrack = lng_hWnd
          End With
          If m_bTrackUser32 Then
               Call TrackMouseEvent(tme)
          Else
               Call TrackMouseEventComCtl(tme)
          End If
     End If
End Sub ' pTrackMouseLeave

Private Function IsFunctionExported(ByVal sFunction As String, _
     ByVal sModule As String) As Boolean
     Dim hmod As Long
     Dim bLibLoaded As Boolean
     hmod = GetModuleHandleA(sModule)
     If hmod = 0 Then
          hmod = LoadLibrary(sModule)
          If hmod Then bLibLoaded = True
     End If
     If hmod Then _
          If GetProcAddress(hmod, sFunction) Then IsFunctionExported = True
     If bLibLoaded Then Call FreeLibrary(hmod)
End Function ' IsFunctionExported

Public Function LoWord(ByVal Value As Long) As Integer
     ' This check is required to prevent an overflow error:
     If ((Value And &HFFFF&) > &H7FFF&) Then
          LoWord = (Value And &HFFFF&) - &H10000
     Else
          LoWord = Value And &HFFFF&
     End If
End Function ' LoWord

Public Function HiWord(ByVal Value As Long) As Integer
     HiWord = (Value And &HFFFF0000) \ &H10000
End Function ' HiWord

Private Function GetMasterVolume(mxc As MIXERCONTROL) As Long
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdu)
           hMem = GlobalAlloc(&H40, Len(mxcdu))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' Get the control value
     lRtn = mixerGetControlDetails(m_hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
     ' Copy the data into the control value buffer
     CopyStructFromPtr mxcdu, mxcd.paDetails, Len(mxcdu)
     ' free allocated memory
     GlobalFree (hMem)
     ' Return the function
     GetMasterVolume = mxcdu.dwValue
End Function ' GetMasterVolume

Private Function GetMasterMute(mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdb As MIXERCONTROLDETAILS_BOOLEAN
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdb)
           hMem = GlobalAlloc(&H40, Len(mxcdb))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
    ' Get the control value
    lRtn = mixerGetControlDetails(m_hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    ' Copy the data into the control value buffer
    CopyStructFromPtr mxcdb, mxcd.paDetails, Len(mxcdb)
    ' Free allocated memory
    GlobalFree (hMem)
    ' Return function
    GetMasterMute = IIf((Abs(mxcdb.fValue) = 1), True, False)
End Function ' GetMasterMute

Private Function GetMasterVolumeControl(lCtrlType As Long, mxc As MIXERCONTROL, _
     lControlType As Long) As Boolean
     Dim mxlc As MIXERLINECONTROLS
     Dim mxl As MIXERLINE
     Dim hMem As Long
     Dim lRtn As Long
     mxl.cbStruct = Len(mxl)
     mxl.dwComponentType = lCtrlType
     ' Obtain a line corresponding to the component type
     lRtn = mixerGetLineInfo(m_hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
     If (lRtn = MMSYSERR_NOERROR) Then
          With mxlc
               .cbStruct = Len(mxlc)
               .dwLineID = mxl.dwLineID
               .dwControl = lControlType
               .cControls = 1
               .cbmxctrl = Len(mxc)
          End With
          ' Allocate memory for the control
          hMem = GlobalAlloc(&H40, Len(mxc))
          mxlc.pamxctrl = GlobalLock(hMem)
          mxc.cbStruct = Len(mxc)
          ' Get the control
          lRtn = mixerGetLineControls(m_hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
          ' function succeeded?
          If (lRtn = MMSYSERR_NOERROR) Then
               GetMasterVolumeControl = True
               ' Copy the control into the destination structure
               CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
          End If
          GlobalFree (hMem)
     End If
End Function ' GetMasterVolumeControl

Private Function GetValue() As Long
     Dim lValue As Long
     On Error Resume Next
     ' convert value
     If m_mxc_vol.lMaximum > False Then
          lValue = m_mxc_vol.lMaximum * (GetValue / 100)
          SetMasterVolume lValue, m_mxc_vol
     End If
End Function ' GetValue

Private Function SetMasterVolume(lValue As Long, mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdu)
          ' Allocate a buffer for the control value buffer
           hMem = GlobalAlloc(&H40, Len(mxcdu))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' set value
     mxcdu.dwValue = lValue
     ' Copy the data into the control value buffer
     CopyPtrFromStruct mxcd.paDetails, mxcdu, Len(mxcdu)
     ' Set the control value
     lRtn = mixerSetControlDetails(m_hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     If (lRtn = MMSYSERR_NOERROR) Then SetMasterVolume = True
End Function ' SetMasterVolume

Private Function SetMasterMute(ByVal bValue As Boolean, mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdb As MIXERCONTROLDETAILS_BOOLEAN
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdb)
          ' Allocate a buffer for the control value buffer
           hMem = GlobalAlloc(&H40, Len(mxcdb))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' set value
     mxcdb.fValue = CLng(bValue)
     ' Copy the data into the control value buffer
     CopyPtrFromStruct mxcd.paDetails, mxcdb, Len(mxcdb)
     ' Set the control value
     lRtn = mixerSetControlDetails(m_hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     If (lRtn = MMSYSERR_NOERROR) Then SetMasterMute = True
End Function ' SetMasterMute

Private Sub UpdateRegistry(sProp As String, sValue As String, sDefValue As String)
     Dim lRtn As Long
     Dim hKey As Long
     ' Write values to registry
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Clockster\Clockster\Settings", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False And hKey Then
          ' write the values to the registry
          SetRegValue hKey, sProp, sValue, sDefValue
     End If
     ' close the key
     If hKey Then Call RegCloseKey(hKey)
End Sub ' UpdateRegistry

Public Sub SoundPlay()
     Dim lRtn As Long
     Static sndData() As Byte
     If Len(m_SoundPath) = False Or m_Sound = UseInternalSound Then
          sndData = LoadResData(101, "DEFAULT")
          Call PlaySoundData(sndData(0), 0, SND_ASYNC Or SND_MEMORY)
     Else
          If Len(m_SoundPath) Then _
               Call PlaySound(m_SoundPath, 0&, SND_ASYNC Or SND_FILENAME)
     End If
End Sub ' SoundPlay

Public Sub PositionTip()
     Dim waRect As tRECT
     Dim lTop As Long
     Dim lLeft As Long
     Dim lHeight As Long
     Dim lWidth As Long
     Dim tbPos As APPBARDATA
     ' Get the screen dimensions in waRECT
     SystemParametersInfo SPI_GETWORKAREA, 0, waRect, 0
     ' get taskbar position to determine where our form is located
     SHAppBarMessage ABM_GETTASKBARPOS, tbPos
     Select Case tbPos.uEdge
          Case ABE_LEFT
               ' get window position relative to the upper left corner of the screen
               lTop = ScaleY(frmClockFace.ClockRip.ClockClientTop, vbPixels, vbTwips)
               lLeft = ScaleX(waRect.Left, vbPixels, vbTwips)
               frmTip.Move lLeft, lTop
          Case ABE_TOP
               lWidth = ScaleX(waRect.Right, vbPixels, vbTwips)
               lTop = ScaleY(waRect.Top, vbPixels, vbTwips)
               ' We got the work area bounds.
               frmTip.Move lWidth - frmTip.Width, lTop
          Case ABE_RIGHT
               ' get window position relative to the upper left corner of the screen
               lTop = ScaleY(frmClockFace.ClockRip.ClockClientTop, vbPixels, vbTwips)
               lLeft = ScaleX(waRect.Right, vbPixels, vbTwips) - frmTip.Width
               frmTip.Move lLeft, lTop
          Case ABE_BOTTOM
               ' convert
               lWidth = ScaleX(waRect.Right, vbPixels, vbTwips)
               lHeight = ScaleY(waRect.Bottom, vbPixels, vbTwips)
               ' We got the work area bounds.
               frmTip.Move lWidth - frmTip.Width, lHeight - frmTip.Height
     End Select
End Sub ' SoundPlay

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Dim xPos As Long
     Dim yPos As Long
     Dim rcTemp As RECT
     Dim lValue As Long
     Select Case uMsg
          Case WM_CONTEXTMENU
               bHandled = True
               RaiseEvent OnContextMenu
          Case WM_MOUSEMOVE
               If Not m_bInCtrl Then
                    m_bInCtrl = True
                    Call pTrackMouseLeave(hwnd)
                    RaiseEvent MouseEnter
                    UserControl.SetFocus
               End If
               ' get x/y positions
               xPos = LoWord(lParam)
               If m_IsDrag Then
                    ' subtract off the left border from the value of x
                    xPos = xPos - 4
                    lValue = (xPos / (ScaleWidth - 7)) * 100
                    ' set the value
                    If lValue < 0 Then lValue = 0
                    If lValue > 100 Then lValue = 100
                    ' if the value has really changed, update label.  This code
                    ' helps relieve the flicker
                    If lValue <> m_Value Then
                         ' Show tip
                         If frmTip.TipEnabled Then
                              If m_Mute Then
                                   frmTip.lblTip = "Level:  " + CStr(lValue) + Chr(37) + _
                                        " - Muted"
                              Else
                                   frmTip.lblTip = "Level:  " + CStr(lValue) + Chr(37)
                              End If
                              ' refresh
                              frmTip.lblTip.Refresh
                         End If
                    End If
                    Value = lValue
               End If
          Case WM_MOUSELEAVE
               m_bInCtrl = False
               frmTip.Hide
               RaiseEvent MouseLeave
          Case WM_MOUSEWHEEL
               If m_bInCtrl Then
                    Select Case wParam
                         Case Is > False
                              ' up
                              lValue = m_Value + 5
                              ' make sure we stay at upper limit
                              If lValue > 100 Then lValue = 100
                              ' set value
                              Value = lValue
                         Case Else
                              ' down
                              lValue = m_Value - 5
                              ' ensure we stay at lower limit
                              If lValue < False Then lValue = False
                              ' set value
                              Value = lValue
                    End Select
                    ' Show tip
                    PositionTip
                    frmTip.Visible = True
                    frmTip.Show
                    If frmTip.TipEnabled Then
                         If m_Mute Then
                              frmTip.lblTip = "Level:  " + CStr(lValue) + Chr(37) + _
                                   " - Muted"
                         Else
                              frmTip.lblTip = "Level:  " + CStr(lValue) + Chr(37)
                         End If
                         ' refresh
                         frmTip.lblTip.Refresh
                    End If
                    Value = lValue
                    ' Raise event
                    RaiseEvent ValueChange(m_Value)
                    ' play sound
                    If m_EnableSound Then SoundPlay
                    UserControl.SetFocus
               End If
          Case WM_LBUTTONDOWN
               xPos = LoWord(lParam)
               m_IsDrag = True
               ' subtract off the left border from the value of x
               xPos = xPos - 4
               lValue = (xPos / (ScaleWidth - 7)) * 100
               ' set the value
               If lValue < 0 Then lValue = 0
               If lValue > 100 Then lValue = 100
               Value = lValue
               PositionTip
               frmTip.Visible = True
               frmTip.Show
          Case WM_LBUTTONUP
               ' play sound
               If m_EnableSound Then SoundPlay
               m_IsDrag = False
               UserControl.SetFocus
               frmTip.Hide
          Case MM_MIXM_CONTROL_CHANGE
               Mute = GetMasterMute(m_mxc_mute)
               lValue = GetMasterVolume(m_mxc_vol)
               ' Convert lValue to be within our limits
               Value = (lValue / m_mxc_vol.lMaximum) * 100
     End Select
End Sub ' iSubclass_Proc

'**************************************************************************************************
' UserControl Intrinsic Events
'**************************************************************************************************
Private Sub UserControl_DblClick()
     Dim bMute As Boolean
     bMute = Mute
     If bMute Then
          Mute = False
     Else
          Mute = True
     End If
End Sub ' UserControl_DblClick

Private Sub UserControl_Initialize()
     ' Set default Values
     m_Color = m_def_Color
     hDC = UserControl.hDC
     hwnd = UserControl.hwnd
     m_Max = 100
     m_Min = 0
     m_ProgressStyle = 0
     m_SegmentSize = 2
     m_SegmentSpace = 1
     m_Value = 0
     pDrawProgressBar
End Sub ' UserControl_Initialize

Private Sub UserControl_Paint()
     pDrawProgressBar
End Sub ' UserControl_Paint

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     Dim lRtn As Long
     Dim lValue As Long
     ' doing it here because this is the first place to get
     ' the ambient object....only want the subclass to start
     ' in run/execution mode(s)...
     ' begin the subclassing and mousetracking
     If Ambient.UserMode Then
          lRtn = mixerOpen(m_hMixer, 0, UserControl.hwnd, 0, CALLBACK_WINDOW)
          ' if successful
          If (lRtn = MMSYSERR_NOERROR) Then
               ' Get the master volume control
               lRtn = GetMasterVolumeControl(MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                    m_mxc_vol, MIXERCONTROL_CONTROLTYPE_VOLUME)
               ' get the value of volume
               If lRtn Then lValue = GetMasterVolume(m_mxc_vol)
               ' Convert lValue to be within our limits
               Value = (lValue / m_mxc_vol.lMaximum) * 100
               ' Are we muted?
               lRtn = GetMasterVolumeControl(MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, m_mxc_mute, _
                    MIXERCONTROL_CONTROLTYPE_MUTE)
               ' if successful, get control's value
               If lRtn Then Mute = GetMasterMute(m_mxc_mute)
          End If
          ' create subclasser object
          Set m_sc = New cSubclass
          m_bTrack = True
          m_bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
          If Not m_bTrackUser32 Then _
               If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then m_bTrack = False
          If m_bTrack Then
               ' add messages to trap
               With m_sc
                    .AddMsg WM_MOUSEMOVE, MSG_AFTER
                    .AddMsg WM_MOUSELEAVE, MSG_AFTER
                    .AddMsg WM_MOUSEWHEEL, MSG_BEFORE
                    .AddMsg WM_LBUTTONDOWN, MSG_AFTER
                    .AddMsg WM_LBUTTONUP, MSG_AFTER
                    .AddMsg WM_CONTEXTMENU, MSG_BEFORE
                    .AddMsg MM_MIXM_CONTROL_CHANGE, MSG_AFTER
               End With
          Else
               ' add messages to trap
               With m_sc
                    .AddMsg WM_MOUSEWHEEL, MSG_BEFORE
                    .AddMsg WM_LBUTTONDOWN, MSG_AFTER
                    .AddMsg WM_LBUTTONUP, MSG_AFTER
                    .AddMsg WM_CONTEXTMENU, MSG_BEFORE
                    .AddMsg MM_MIXM_CONTROL_CHANGE, MSG_AFTER
               End With
          End If
          ' start subclasser
          m_sc.Subclass UserControl.hwnd, Me
     End If
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     hDC = UserControl.hDC
     pDrawProgressBar
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     Dim sRtn As String
     Dim hKey As Long
     Dim lRtn As Long
     Dim lStyle As Long
     ' retrieve settings from registry
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Clockster\Clockster\Settings", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_QUERY_VALUE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False And hKey Then
          ' are we playing a sound when slider is changed
          sRtn = GetRegValue(hKey, "TouchSliderEnableSound")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               EnableSound = CBool(sRtn)
          Else ' set to default
               EnableSound = m_def_EnableSound
               ' set the registry
               UpdateRegistry "TouchSliderEnableSound", CStr(m_def_EnableSound), _
                    CStr(m_def_EnableSound)
          End If
          ' get slider forecolor
          sRtn = GetRegValue(hKey, "TouchSliderForeColor")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               Color = CLng(sRtn)
          Else ' set to default
               Color = m_def_Color
               ' set registry
               UpdateRegistry "TouchSliderForeColor", CStr(m_def_Color), _
                    CStr(m_def_Color)
          End If
          ' get slider forecolor
          sRtn = GetRegValue(hKey, "TouchSliderHeightAdjust")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               HeightAdjust = CLng(sRtn)
          Else ' set to default
               HeightAdjust = m_def_HeightAdjust
               ' set registry
               UpdateRegistry "TouchSliderHeightAdjust", CStr(m_def_HeightAdjust), _
                    CStr(m_def_HeightAdjust)
          End If
          ' get slider style
          sRtn = GetRegValue(hKey, "TouchSliderProgressStyle")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               ProgressStyle = CLng(sRtn)
          Else ' set to default
               ProgressStyle = m_def_ProgressStyle
               ' set registry
               UpdateRegistry "TouchSliderProgressStyle", CStr(m_def_ProgressStyle), _
                    CStr(m_def_ProgressStyle)
          End If
          ' Get size of segments
          sRtn = GetRegValue(hKey, "TouchSliderSegmentSize")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               SegmentSize = CLng(sRtn)
          Else ' set to default
               SegmentSize = m_def_SegmentSize
               ' set registry
               UpdateRegistry "TouchSliderSegmentSize", CStr(m_def_SegmentSize), _
                    CStr(m_def_SegmentSize)
          End If
          ' Get width of space between segments
          sRtn = GetRegValue(hKey, "TouchSliderSegmentSpace")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               SegmentSpace = CLng(sRtn)
          Else ' set to default
               SegmentSpace = m_def_SegmentSpace
               ' set registry
               UpdateRegistry "TouchSliderSegmentSpace", CStr(m_def_SegmentSpace), _
                    CStr(m_def_SegmentSpace)
          End If
          ' get the slider's y offset
          sRtn = GetRegValue(hKey, "TouchSliderOffsetY")
          ' if we have a value
          If Len(sRtn) Then
               ' set the property reg value
               SliderOffsetY = CLng(sRtn)
          Else ' set to default
               SliderOffsetY = m_def_offsetY
               ' set registry
               UpdateRegistry "TouchSliderOffsetY", CStr(m_def_offsetY), CStr(m_def_offsetY)
          End If
          ' get the slider's sound path
          sRtn = GetRegValue(hKey, "TouchSliderSound")
          ' if we have a value
          If Len(sRtn) Then
               ' set the property reg value
               Sound = CInt(sRtn)
          Else ' set to default
               Sound = m_def_Sound
               ' set registry
               UpdateRegistry "TouchSliderSound", CStr(m_def_Sound), CStr(m_def_Sound)
          End If
          ' get the slider's sound path
          sRtn = GetRegValue(hKey, "TouchSliderSoundPath")
          ' if we have a value
          If Len(sRtn) Then
               ' set the property reg value
               SoundPath = sRtn
          Else ' set to default
               SoundPath = ""
               ' set registry
               UpdateRegistry "TouchSliderSoundPath", "", ""
          End If
     End If
     ' close the key
     Call RegCloseKey(hKey)
     ' set window style so nothing writes on us
     ' Get the styles owned by the window already
     lStyle = GetWindowLong(UserControl.hwnd, GWL_STYLE)
     ' strip out styles passed to sub
     lStyle = lStyle And Not 0
     ' add style passed to sub
     lStyle = lStyle Or WS_CLIPCHILDREN
     ' Set the window style
     SetWindowLong UserControl.hwnd, GWL_STYLE, lStyle
End Sub ' UserControl_Show

Private Sub UserControl_Terminate()
     On Error Resume Next
     ' Close the mixer
     If m_hMixer Then mixerClose m_hMixer
     ' Destroy Temp DC
     pDestroy
     ' Destroy subclass object
     m_sc.UnSubclass
     Set m_sc = Nothing
End Sub ' UserControl_Terminate

