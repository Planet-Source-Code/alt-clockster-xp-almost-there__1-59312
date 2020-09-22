VERSION 5.00
Begin VB.UserControl CustomClock 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   10
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   59
   ToolboxBitmap   =   "CustomClock.ctx":0000
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   -135
      Top             =   -165
   End
End
Attribute VB_Name = "CustomClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' CustomClock.ctl - just a little companion control for the trayvolume project.
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' CustomClock Constant Declares
'**************************************************************************************************
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TWIPS_INCH As Long = 1440

'**************************************************************************************************
' CustomClock Enums
'**************************************************************************************************
Public Enum eCCChime
     [UseInternalChime]
     [UseExternalChime]
End Enum ' eCCChime

Public Enum eCCChimeInterval
     [Never]
     [Every Hour]
     [Every Half-Hour]
     [Every Quarter-Hour]
End Enum ' eCCChimeInterval

Public Enum eCCTimeDisplayFormat
     [12hr]
     [12hr-AMPM]
     [12hr-Seconds]
     [12hr-Seconds-AMPM]
     [24hr]
     [24hr-Seconds]
End Enum ' eTimeDisplayFormat

'**************************************************************************************************
' CustomClock Structs
'**************************************************************************************************
' Structs declared in mdlDeclares/Winsubhook2.tlb

'**************************************************************************************************
' CustomClock Win32 API
'**************************************************************************************************
' API used in this control declared in mdlDeclares/Winsubhook2.tlb
Public Event OnContextMenu()
Public Event OnPositionChange(ByVal lOffsetY As Long)

'**************************************************************************************************
' CustomClock Module-Scoped Variables
'**************************************************************************************************
Private m_bSoundPlayed As Boolean
Private m_IsSubclassed As Boolean
Private m_isWinXP As eWinXPTest
Private m_sOldTime As String
Private m_sc As cSubclass
Implements WinSubHook2.iSubclass
Private m_tmr As cTimer
Implements WinSubHook2.iTimer

'**************************************************************************************************
' CustomClock Default Property Constants
'**************************************************************************************************
Private Const m_def_Chime = 0
Private Const m_def_ChimeInterval = 3
Private Const m_def_ChimePath = "C:\"
Private Const m_def_fontname = "Tahoma"
Private Const m_def_fontbold = False
Private Const m_def_fontitalic = False
Private Const m_def_fontsize = 8.25
Private Const m_def_forecolor = 16711680
Private Const m_def_offsetY = 0
Private Const m_def_TimeDisplayFormat = 2

'**************************************************************************************************
' CustomClock Property Variables
'**************************************************************************************************
Private m_Chime As eCCChime
Private m_ChimeInterval As eCCChimeInterval
Private m_ChimePath As String
Private m_ClockMinute As String
Private m_ClockTime As String
Private m_ClockOffsetY As Long
Private m_TimeDisplayFormat As eCCTimeDisplayFormat

'**************************************************************************************************
' CustomClock Properties
'**************************************************************************************************
Public Property Get Chime() As eCCChime
     Chime = m_Chime
End Property ' Get Chime

Public Property Let Chime(New_Chime As eCCChime)
     ' store in property
     m_Chime = New_Chime
     ' store in registry
     UpdateRegistry "ClockChime", CStr(m_Chime), CStr(m_def_Chime)
End Property ' Let Chime

Public Property Get ChimeInterval() As eCCChimeInterval
     ChimeInterval = m_ChimeInterval
End Property ' Get ChimeInterval

Public Property Let ChimeInterval(New_ChimeInterval As eCCChimeInterval)
     ' store in property
     m_ChimeInterval = New_ChimeInterval
     ' store in registry
     UpdateRegistry "ClockChimeInterval", CStr(m_ChimeInterval), CStr(m_def_ChimeInterval)
End Property ' Let ChimeInterval

Public Property Get ChimePath() As String
     ChimePath = m_ChimePath
End Property ' Get ChimePath

Public Property Let ChimePath(New_ChimePath As String)
     ' store in property
     m_ChimePath = New_ChimePath
     ' store in registry
     UpdateRegistry "ClockChimePath", m_ChimePath, m_def_ChimePath
End Property ' Let ChimePath

Public Property Get ClockTime() As String
     ClockTime = m_ClockTime
End Property ' Get ClockTime

Public Property Get ClockOffsetY() As Long
     ClockOffsetY = m_ClockOffsetY
End Property ' Get ClockOffsetY

Public Property Let ClockOffsetY(New_ClockOffsetY As Long)
     ' store in property
     m_ClockOffsetY = New_ClockOffsetY
     ' store in registry
     UpdateRegistry "ClockOffsetY", CStr(m_ClockOffsetY), CStr(m_def_offsetY)
     ' raise the event
     RaiseEvent OnPositionChange(m_ClockOffsetY)
End Property ' Let ClockOffsetY

Public Property Get Font() As Font
     Set Font = UserControl.Font
     DrawClock
End Property ' Get Font

Public Property Set Font(ByVal New_Font As Font)
     ' update usercontrol
     Set UserControl.Font = New_Font
     ' update registry with font metrics
     UpdateRegistry "ClockFontName", UserControl.Font.Name, m_def_fontname
     UpdateRegistry "ClockFontBold", CStr(UserControl.Font.Bold), m_def_fontbold
     UpdateRegistry "ClockFontItalic", CStr(UserControl.Font.Italic), m_def_fontitalic
     UpdateRegistry "ClockFontSize", CStr(UserControl.Font.Size), m_def_fontsize
     ' update clock
     DrawClock
End Property ' Set Font

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = UserControl.ForeColor
End Property ' Get ForeColor

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
     ' update usercontrol
     UserControl.ForeColor = New_ForeColor
     ' store in registry
     UpdateRegistry "ClockForeColor", CStr(New_ForeColor), CStr(m_def_forecolor)
     ' update clock
     DrawClock
End Property ' Let ForeColor

Public Property Get hDC() As Long
     hDC = UserControl.hDC
End Property ' Get hDC

Public Property Let hDC(New_hDC As Long)
     ' do nothing...just want to see it in the property browser
End Property ' Let hDC

Public Property Get hwnd() As Long
     hwnd = UserControl.hwnd
End Property ' Get hWnd

Public Property Let hwnd(New_hWnd As Long)
     ' do nothing...just want to see it in the property browser
End Property

Public Property Get TimeDisplayFormat() As eCCTimeDisplayFormat
     TimeDisplayFormat = m_TimeDisplayFormat
End Property ' Get TimeDisplayFormat

Public Property Let TimeDisplayFormat(New_TimeDisplayFormat As eCCTimeDisplayFormat)
     ' store in property
     m_TimeDisplayFormat = New_TimeDisplayFormat
     ' store in registry
     UpdateRegistry "ClockTimeDisplayFormat", CStr(m_TimeDisplayFormat), _
          CStr(m_def_TimeDisplayFormat)
     ' update clock
     DrawClock
End Property ' Let TimeDisplayFormat

'**************************************************************************************************
' CustomClock Intrinsic Events
'**************************************************************************************************
Private Sub UserControl_DblClick()
     ' call the time/date control panel applet
     Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub ' UserControl_DblClick

Private Sub UserControl_Resize()
     ' redraw the clock based on the new control rectangle
     DrawClock
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     Dim sRtn As String
     Dim hKey As Long
     Dim lRtn As Long
     ' retrieve settings from registry
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Clockster\Clockster\Settings", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_QUERY_VALUE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False And hKey Then
          ' get chime (internal/external)
          sRtn = GetRegValue(hKey, "ClockChime")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               Chime = CLng(sRtn)
          Else ' set to default
               Chime = m_def_Chime
               ' set registry
               UpdateRegistry "ClockChime", CStr(m_def_Chime), CStr(m_def_Chime)
          End If
          ' get chime interval
          sRtn = GetRegValue(hKey, "ClockChimeInterval")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               ChimeInterval = CLng(sRtn)
          Else ' set to default
               ChimeInterval = m_def_ChimeInterval
               ' set registry
               UpdateRegistry "ClockChimeInterval", CStr(m_def_ChimeInterval), CStr(m_def_ChimeInterval)
          End If
          ' Get path to chime sound
          sRtn = GetRegValue(hKey, "ClockChimePath")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               ChimePath = sRtn
          Else ' set to default
               ChimePath = m_def_ChimePath
               ' set registry
               UpdateRegistry "ClockChimePath", CStr(m_def_ChimePath), CStr(m_def_ChimePath)
          End If
          ' Get clock font name
          sRtn = GetRegValue(hKey, "ClockFontName")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               UserControl.Font.Name = sRtn
          Else ' set to default
               UserControl.Font.Name = m_def_fontname
               ' set registry
               UpdateRegistry "ClockFontName", CStr(m_def_fontname), CStr(m_def_fontname)
          End If
          ' is font bold
          sRtn = GetRegValue(hKey, "ClockFontBold")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               UserControl.Font.Bold = CBool(sRtn)
          Else ' set to default
               UserControl.Font.Bold = m_def_fontbold
               ' set registry
               UpdateRegistry "ClockFontBold", CStr(m_def_fontbold), CStr(m_def_fontbold)
          End If
          ' is font italic
          sRtn = GetRegValue(hKey, "ClockFontItalic")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               UserControl.Font.Italic = CBool(sRtn)
          Else ' set to default
               UserControl.Font.Italic = m_def_fontitalic
               ' set registry
               UpdateRegistry "ClockFontItalic", CStr(m_def_fontitalic), CStr(m_def_fontitalic)
          End If
          ' get the font's size
          sRtn = GetRegValue(hKey, "ClockFontSize")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               UserControl.Font.Size = CSng(sRtn)
          Else ' set to default
               UserControl.Font.Size = m_def_fontsize
               ' set registry
               UpdateRegistry "ClockFontSize", CStr(m_def_fontsize), CStr(m_def_fontsize)
          End If
          ' get the clock's forecolor
          sRtn = GetRegValue(hKey, "ClockForeColor")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               UserControl.ForeColor = CLng(sRtn)
          Else ' set to default
               UserControl.ForeColor = m_def_forecolor
               ' set registry
               UpdateRegistry "ClockForeColor", CStr(m_def_forecolor), CStr(m_def_forecolor)
          End If
          ' get the clock's y offset
          sRtn = GetRegValue(hKey, "ClockOffsetY")
          ' if we have a value
          If Len(sRtn) Then
               ' set the property reg value
               ClockOffsetY = CLng(sRtn)
          Else ' set to default
               ClockOffsetY = m_def_offsetY
               ' set registry
               UpdateRegistry "ClockOffsetY", CStr(m_def_offsetY), CStr(m_def_offsetY)
          End If
          ' how do we display clock?
          sRtn = GetRegValue(hKey, "ClockTimeDisplayFormat")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               TimeDisplayFormat = CLng(sRtn)
          Else ' set to default
               TimeDisplayFormat = m_def_TimeDisplayFormat
               ' set registry
               UpdateRegistry "ClockTimeDisplayFormat", CStr(m_def_TimeDisplayFormat), _
                    CStr(m_def_TimeDisplayFormat)
          End If
     End If
     ' close the key
     Call RegCloseKey(hKey)
     ' draw clock
     DrawClock
     ' get current time
     m_ClockTime = ProcessTimeFormat
     ' If not in design
     If Ambient.UserMode Then
          ' create subclass object
          Set m_sc = New cSubclass
          ' add the message we want to trap
          m_sc.AddMsg WM_CONTEXTMENU, MSG_BEFORE
          ' begin subclasser
          m_sc.Subclass hwnd, Me
          ' enable timer
          'tmrUpdate.Enabled = True
          Set m_tmr = New cTimer
          m_tmr.TmrStart Me, 30
          '
          m_IsSubclassed = True
     End If
End Sub ' UserControl_Show

Private Sub UserControl_Terminate()
     Dim hKey As Long
     Dim lRtn As Long
     ' if subclassed
     If m_IsSubclassed Then
          ' stop timer
          m_tmr.TmrStop
          ' destroy timer object
          Set m_tmr = Nothing
          ' stop
          m_sc.UnSubclass
          ' destroy subclass object
          Set m_sc = Nothing
     End If
End Sub ' UserControl_Terminate

'**************************************************************************************************
' CustomClock Constituent Control Events
'**************************************************************************************************
Private Sub tmrUpdate_Timer()
     DrawClock
End Sub ' tmrUpdate_Timer

'**************************************************************************************************
' CustomClock Implemented Subclass Interface
'**************************************************************************************************
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Select Case uMsg
          Case WM_CONTEXTMENU
               bHandled = True
               RaiseEvent OnContextMenu
     End Select
End Sub ' iSubclass_Proc

Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
     DrawClock
End Sub ' iTimer_Proc

'**************************************************************************************************
' CustomClock Methods
'**************************************************************************************************
Private Function DrawClock()
     Dim rcClk As tRECT
     Dim lTxtHt As Long
     ' get the selected clock format
     m_ClockTime = ProcessTimeFormat
     ' did the clock caption change. If so, update our clock
     If m_ClockTime <> m_sOldTime Then
          ' reset m_sOldTime to the new time
          m_sOldTime = m_ClockTime
          ' Get window rectangle to pass to DrawText function
          rcClk.Top = -3
          rcClk.Right = ScaleWidth - 1
          rcClk.Bottom = ScaleHeight
          ' replace maskpicture
          Set MaskPicture = Nothing
          ' clear the display
          Cls
          ' output it to the control
          lTxtHt = DrawText(hDC, m_ClockTime, Len(m_ClockTime), rcClk, DT_BOTTOM Or DT_CENTER)
          ' set usercontrol height to that of text height
          Height = lTxtHt * Screen.TwipsPerPixelY
          ' refresh uc display
          Refresh
          ' set the maskimage so our clock will display on the transparent background
          Set MaskPicture = Image
          ' Get the minute
          m_ClockMinute = Minute(m_ClockTime)
          ' if minute is only one char then prepend a 0
          If Len(m_ClockMinute) = 1 Then m_ClockMinute = Chr(48) + m_ClockMinute
          ' is it chime time?
          Select Case m_ChimeInterval
               Case 1 ' interval is every hour
                    If m_ClockMinute = "00" Then
                         If m_bSoundPlayed = False Then
                              PlayChime
                              m_bSoundPlayed = True
                         End If
                    Else
                         m_bSoundPlayed = False
                    End If
               Case 2 ' interval is every 30 minutes
                    If m_ClockMinute = "00" Or m_ClockMinute = "30" Then
                         If m_bSoundPlayed = False Then
                              PlayChime
                              m_bSoundPlayed = True
                         End If
                    Else
                         m_bSoundPlayed = False
                    End If
               Case 3 ' inteval is every 15 minutes
                    If m_ClockMinute = "00" Or m_ClockMinute = "15" Or _
                         m_ClockMinute = "30" Or m_ClockMinute = "45" Then
                         If m_bSoundPlayed = False Then
                              PlayChime
                              m_bSoundPlayed = True
                         End If
                    Else
                         m_bSoundPlayed = False
                    End If
          End Select
     End If
     DoEvents
End Function ' DrawClock

Private Function ProcessTimeFormat() As String
     ProcessTimeFormat = Choose(m_TimeDisplayFormat + 1, Format(Time, "h:mm"), _
          Format(Time, "h:mm AMPM"), Format(Time, "h:mm:ss"), _
          Format(Time, "h:mm:ss AMPM"), Format(Time, "hh:mm"), _
          Format(Time, "hh:mm:ss"))
End Function ' ProcessTimeFormat

Private Sub PlayChime()
     Static sndData() As Byte
     Select Case m_Chime
          Case 0 ' play internal wav from resource file
               sndData = LoadResData(101, "CHIME")
               Call PlaySoundData(sndData(0), 0, SND_ASYNC Or SND_MEMORY)
          Case 1 ' play external wav file if have a path to one
               If Len(m_ChimePath) Then _
                    Call PlaySound(m_ChimePath, 0&, SND_ASYNC Or SND_FILENAME)
     End Select
End Sub ' PlayChime

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
