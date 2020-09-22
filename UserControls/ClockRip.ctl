VERSION 5.00
Begin VB.UserControl ClockRip 
   AutoRedraw      =   -1  'True
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ClockRip.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "ClockRip.ctx":0420
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   540
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   0
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "ClockRip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************************************************
' ClockMetrics Constant Declares
'**************************************************************************************************
Private Const VER_PLATFORM_WIN32_NT = 2

'**************************************************************************************************
' ClockMetrics Enum Declares
'**************************************************************************************************
Public Enum APPBAREDGE
     ABE_LEFT = 0
     ABE_TOP = 1
     ABE_RIGHT = 2
     ABE_BOTTOM = 3
End Enum ' APPBAREDGE

Public Enum eWinStyle
     NO_STYLE = &H0
     WS_CHILD = &H40000000
     WS_CLIPCHILDREN = &H2000000
End Enum ' eWinStyle

Public Enum eWinXPTest
     [IsNotXP]
     [IsXP]
End Enum ' eWinXPTest

'**************************************************************************************************
' ClockMetrics Struct Declares
'**************************************************************************************************
Private Type WINDOWINFO
     cbSize As Long
     rcWindow As tRECT
     rcClient As tRECT
     dwStyle As Long
     dwExStyle As Long
     cxWindowBorders As Long
     cyWindowBorders As Long
     atomWindowtype As Long
     wCreatorVersion As Long
End Type ' WINDOWINFO

'**************************************************************************************************
' ClockMetrics Win32 API Declares
'**************************************************************************************************
Private Declare Function CopyRect Lib "user32" (lpDestRect As tRECT, lpSourceRect As tRECT) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As tRECT, lpRect2 As tRECT) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
     ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
     ByVal ParentHwnd As Long, ByVal Firsthwnd As Long, ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long
Private Declare Function GetWindowInfo Lib "user32" (ByVal hwnd As Long, _
     ByRef pwi As WINDOWINFO) As Boolean
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
     ByVal hWndNewParent As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal crColor As Long) As Long
     
'**************************************************************************************************
' ClockMetrics Events
'**************************************************************************************************
Public Event OnAppBarPositionChange(tbPos As APPBAREDGE, ByVal lPosLeft As Long, _
     ByVal lPosTop As Long, ByVal lPosRight As Long, ByVal lPosBottom As Long, _
     lHeight As Long, lWidth As Long)
Public Event OnClockRectChange(ByVal lPosLeft As Long, ByVal lPosTop As Long, _
     ByVal lPosRight As Long, ByVal lPosBottom As Long, lHeight As Long, _
     lWidth As Long)
Public Event OnBackgroundChange(backImg As stdole.Picture)

'**************************************************************************************************
' ClockMetrics Module-Scoped Variables
'**************************************************************************************************
Private m_bTmrCreated As Boolean
Private m_oldAppBarPos As APPBAREDGE
Private m_oldClockClientRect As tRECT
Private m_oldParent As Long
Private m_tmr As cTimer
Implements WinSubHook2.iTimer

'**************************************************************************************************
' ClockMetrics Property Variables
'**************************************************************************************************
Private m_AppBarPos As APPBAREDGE
Private m_ClockClientLeft As Long
Private m_ClockClientTop As Long
Private m_ClockClientRect As tRECT
Private m_ClockClientRight As Long
Private m_ClockClientBottom As Long
Private m_ClockHeight As Long
Private m_ClockhWnd As Long
Private m_ClockWidth As Long
Private m_Error As String

'**************************************************************************************************
' ClockMetrics Read-Only Property Statements
'**************************************************************************************************
Public Property Get AppBarPos() As APPBAREDGE
     ' Return property
     AppBarPos = m_AppBarPos
End Property ' Get AppBarPos

Public Property Get ClockClientBottom() As Long
     ' Return property
     ClockClientBottom = m_ClockClientBottom
End Property ' Get ClockClientLeft

Public Property Get ClockClientLeft() As Long
     ' Return property
     ClockClientLeft = m_ClockClientLeft
End Property ' Get ClockClientLeft

Public Property Get ClockClientRect() As tRECT
     ' Return property
     LSet ClockClientRect = m_ClockClientRect
End Property ' Get ClockClientRect

Public Property Get ClockClientRight() As Long
     ' Return property
     ClockClientRight = m_ClockClientRight
End Property ' Get ClockClientRight

Public Property Get ClockClientTop() As Long
     ' Return property
     ClockClientTop = m_ClockClientTop
End Property ' Get ClockClientTop

Public Property Get ClockHeight() As Long
     ' Return property
     ClockHeight = m_ClockHeight
End Property ' Get ClockHeight

Public Property Get ClockhWnd() As Long
     ' Return property
     ClockhWnd = m_ClockhWnd
End Property ' Get ClockhWnd

Public Property Get ClockWidth() As Long
     ' Return property
     ClockWidth = m_ClockWidth
End Property ' Get ClockWidth

Public Property Get Error() As String
     ' Return property
     Error = m_Error
End Property ' Get Error

Public Property Let Error(New_Error As String)
     ' Do nothing....
End Property ' Let Error

Public Property Get Picture() As Picture
     Set Picture = picBuffer.Picture
End Property ' Get Picture

'**************************************************************************************************
' ClockMetrics Public Custom Methods
'**************************************************************************************************
Public Function AdoptClockParent(lhWnd As Long) As Long
     If m_ClockhWnd = False Then _
          m_ClockhWnd = GetClockHandle
     ' if we have the clock's hwnd then re-parent and store
     ' the old parent window handle so we can reset it
     If m_ClockhWnd Then
          ' reparent window passed in param
          m_oldParent = SetParent(lhWnd, m_ClockhWnd)
          ' now that the clock is a parent, set it's style
          ' to clipchildren
          SetWinStyle m_ClockhWnd, WS_CLIPCHILDREN, NO_STYLE
     End If
End Function ' AdoptClockParent

Public Sub DivorceClockParent(lhWnd As Long)
     ' re-parent to original parent
     Call SetParent(lhWnd, m_oldParent)
     ' bail on the clipchildren style
     SetWinStyle m_ClockhWnd, NO_STYLE, WS_CLIPCHILDREN
End Sub ' DivorceClockParent

Public Sub SetWinStyle(ByVal lhWnd As Long, ByVal lWinStyle As eWinStyle, _
     ByVal lWinStyleNot As eWinStyle)
     Dim lStyle As Long
     ' Get the styles owned by the window already
     lStyle = GetWindowLong(lhWnd, GWL_STYLE)
     ' strip out styles passed to sub
     lStyle = lStyle And Not lWinStyleNot
     ' add style passed to sub
     lStyle = lStyle Or lWinStyle
     ' Set the window style
     SetWindowLong lhWnd, GWL_STYLE, lStyle
End Sub ' SetWinStyle

'**************************************************************************************************
' ClockMetrics Private Custom Methods
'**************************************************************************************************
Private Sub DrawClockBackground()
     Dim lhWnd As Long
     Dim lHDC As Long
     Dim lLoop As Long
     ' Get the clock handle
     If m_ClockhWnd = False Then m_ClockhWnd = GetClockHandle
     ' Okay now do we have a clock handle?
     If m_ClockhWnd Then
          ' set local hwnd to clockhwnd
          lhWnd = m_ClockhWnd
          ' First get the clock's DC
          lHDC = GetDC(lhWnd)
          ' If successful
          If lHDC Then
               picBuffer.Cls
               ' what's our appbar orientation
               Select Case m_AppBarPos
                    Case ABE_TOP, ABE_BOTTOM
'                         ' I use this for testing
'                         For lLoop = 0 To m_ClockHeight Step 2
'                              SetPixel lHDC, 0, lLoop, vbRed
'                              SetPixel lHDC, 1, lLoop, vbRed
'                         Next
                         ' blit the clock window onto picBuffer
                         StretchBlt picBuffer.hDC, 0, 0, m_ClockWidth, m_ClockHeight, _
                              lHDC, 0, 0, 1, m_ClockHeight, vbSrcCopy
                    Case ABE_LEFT, ABE_RIGHT
'                         ' I use this for testing
'                         For lLoop = 0 To m_ClockWidth Step 2
'                              SetPixel lHDC, lLoop, 0, vbBlue
'                              SetPixel lHDC, lLoop, 1, vbBlue
'                         Next
                         ' blit the clock window onto picBuffer
                         StretchBlt picBuffer.hDC, 0, 0, m_ClockWidth, m_ClockHeight, _
                              lHDC, 0, 0, m_ClockWidth, 1, vbSrcCopy
               End Select
               ' update control
               picBuffer.Picture = picBuffer.Image
               picBuffer.Refresh
               ' raise event to public new picture
               RaiseEvent OnBackgroundChange(picBuffer.Picture)
               ' Release the dc
               ReleaseDC lhWnd, lHDC
          End If
     End If
End Sub ' DrawClockBackground

Private Function GetAppBarPosition() As APPBAREDGE
     Dim abd As APPBARDATA
     ' Get taskbar position
     SHAppBarMessage ABM_GETTASKBARPOS, abd
     ' Return edge
     GetAppBarPosition = abd.uEdge
End Function ' GetAppBarPosition

Private Function GetClockHandle() As Long
     GetClockHandle = FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), _
          0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString)
End Function ' GetClockHandle

Private Function GetWinInfo(ByVal lhWnd As Long, ByRef wi As WINDOWINFO) As Boolean
     Dim bRtn As Boolean
     ' initialize struct
     wi.cbSize = Len(wi)
     ' call api
     GetWinInfo = GetWindowInfo(lhWnd, wi)
End Function ' GetWinInfo

Public Function IsOSWinXP() As eWinXPTest
     Dim OS As tOSVERSIONINFO
     Dim lRtn As Long
     ' initialize OS struct
     OS.dwOSVersionInfoSize = Len(OS)
     ' call api
     lRtn = GetVersionEx(OS)
     If lRtn Then
          If (OS.dwPlatformId = VER_PLATFORM_WIN32_NT) And _
               (OS.dwMajorVersion = 5 And OS.dwMinorVersion = 1) Then
                    IsOSWinXP = 1
          Else
               IsOSWinXP = 0
          End If
     End If
End Function ' IsOSWinXP

'**************************************************************************************************
' ClockMetrics Implemented Interface Proc
'**************************************************************************************************
Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
     Dim bRtn As Boolean
     Dim lRtn As Long
     Dim wi As WINDOWINFO
     ' Get the clock's handle
     If m_ClockhWnd = False Then m_ClockhWnd = GetClockHandle
     ' If we have a handle
     If m_ClockhWnd Then
          ' Get the clock's window info
          bRtn = GetWinInfo(m_ClockhWnd, wi)
          ' if we get our data
          If bRtn Then
               ' pour into property variables
               With wi
                    m_ClockClientLeft = .rcClient.Left
                    m_ClockClientTop = .rcClient.Top
                    m_ClockClientRect = .rcClient
                    m_ClockClientRight = .rcClient.Right
                    m_ClockClientBottom = .rcClient.Bottom
                    m_ClockHeight = m_ClockClientBottom - m_ClockClientTop
                    m_ClockWidth = m_ClockClientRight - m_ClockClientLeft
               End With
               ' did any part of the rectangle change?
               lRtn = EqualRect(m_oldClockClientRect, m_ClockClientRect)
               ' Keep picBuffer in sync with clock size
               picBuffer.Move picBuffer.Left, picBuffer.Top, _
                    m_ClockWidth * Screen.TwipsPerPixelX, _
                    m_ClockHeight * Screen.TwipsPerPixelY
               picBuffer.Refresh
               ' process return
               If lRtn = False Then
                    ' rectangles are not equal
                    With m_ClockClientRect
                         RaiseEvent OnClockRectChange(.Left, .Top, .Right, .Bottom, _
                              m_ClockHeight, m_ClockWidth)
                    End With
                    ' set old equal to new
                    CopyRect m_oldClockClientRect, m_ClockClientRect
               End If
               ' no error
               m_Error = "NoError"
          Else
               ' GetWindowInfo api failed
               m_Error = "GetWinInfo Failed"
          End If
          ' Where's the taskbar?
          m_AppBarPos = GetAppBarPosition
          ' did position change?
          If m_AppBarPos <> m_oldAppBarPos Then
               ' raise clockrectchange event just for a redraw
               With m_ClockClientRect
                    'RaiseEvent OnClockRectChange(.Left, .Top, .Right, .Bottom, m_ClockHeight, m_ClockWidth)
                    ' raise event
                    RaiseEvent OnAppBarPositionChange(m_AppBarPos, .Left, .Top, .Right, .Bottom, _
                              m_ClockHeight, m_ClockWidth)
               End With
               ' Set old to new
               m_oldAppBarPos = m_AppBarPos
          End If
          ' Repaint background
          DrawClockBackground
     Else
          ' GetClockHandle Failed
          m_Error = "GetClockHandle Failed"
     End If
     DoEvents
End Sub ' iTimer_Proc

'**************************************************************************************************
' ClockMetrics UserControl Intrinsic Methods
'**************************************************************************************************
Private Sub UserControl_Initialize()
     Error = m_Error
End Sub ' UserControl_Initialize

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     ' If not in the IDE
     If Ambient.UserMode Then
          ' Set our timer object
          Set m_tmr = New cTimer
          ' start the timer
          m_tmr.TmrStart Me, 50
          ' timer created
          m_bTmrCreated = True
     End If
     ' set picbuffer's color
     picBuffer.BackColor = Ambient.BackColor
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     Height = 480
     Width = 480
End Sub ' UserControl_Resize

Private Sub UserControl_Terminate()
     ' if not in ide
     If m_bTmrCreated Then
          ' stop the timer
          m_tmr.TmrStop
          ' kill the timer object
          Set m_tmr = Nothing
     End If
End Sub ' UserControl_Terminate
