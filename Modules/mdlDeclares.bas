Attribute VB_Name = "mdlDeclares"
'**************************************************************************************************
' Name:     mDeclares.bas
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     ?
'**************************************************************************************************
' Requires: - None
'**************************************************************************************************
' Copyright © ? Steve McMahon for vbAccelerator
'**************************************************************************************************
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' mdlDeclares Constants
'**************************************************************************************************
Public Const ABM_GETTASKBARPOS = &H5
Public Const FW_BOLD = 700
Public Const LOGPIXELSY = 90
Public Const DT_BOTTOM = &H8
Public Const DT_CENTER = &H1
Public Const DT_CALCRECT = &H400
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0&
Public Const DT_VCENTER = &H4
Public Const GWL_STYLE = (-16)
Public Const LF_FACESIZE = 32
Public Const SPI_GETWORKAREA = 48
' registry constants
Public Const READ_CONTROL = &H20000
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And _
     (Not SYNCHRONIZE))
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or _
     KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
Public Const SND_FILENAME = &H20000
Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2

'**************************************************************************************************
' mdlDeclares Public/Shared Enums & Structs
'**************************************************************************************************
Public Type APPBARDATA
     cbSize As Long
     hwnd As Long
     uCallbackMessage As Long
     uEdge As Long
     rc As tRECT
     lParam As Long
End Type ' APPBARDATA

Public Type LOGFONT
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName(LF_FACESIZE) As Byte
End Type ' LOGFONT

Public Type tMenuItem
     bChecked As Boolean
     bComboBox As Boolean
     bCreated As Boolean
     bDefault As Boolean
     bDragOff As Boolean
     bEnabled As Boolean
     bIsAVBMenu As Boolean
     bMarkToDestroy As Boolean
     bMenuBarBreak As Boolean
     bMenuBreak As Boolean
     bOwnerDraw As Boolean
     bRadioCheck As Boolean
     bShowCheckAndIcon As Boolean
     bTextBox As Boolean
     bTitle As Boolean
     bVisible As Boolean
     hMenu As Long
     iShortCutShiftKey As Integer
     iShortCutShiftMask As Integer
     lActualID As Long
     lHeight As Long
     lID As Long
     lIndex As Long
     lItemData As Long
     lParentId As Long
     lParentIndex As Long
     lShortCutStartPos As Long
     lWidth As Long
     sAccelerator As String
     sCaption As String
     sHelptext As String
     sInputCaption As String
     sKey As String
     sShortCutDisplay As String
End Type ' tMenuItem

'**************************************************************************************************
' mdlDeclares Win32 Project-Scoped API
'**************************************************************************************************
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
     lpPoint As tPOINT) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias _
     "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As tRECT, _
     ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
     ByVal lpStr As String, ByVal nCount As Long, lpRect As tRECT, ByVal wFormat As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
     ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
     (ByVal ParentHwnd As Long, ByVal Firsthwnd As Long, ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, _
     ByVal nDenominator As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
     ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, _
     ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, _
     ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
     "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, _
     ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
     ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
     "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
     "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
     "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
     ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
     "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, _
          ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, _
     pData As APPBARDATA) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
     ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias _
     "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     lpvParam As Any, ByVal fuWinIni As Long) As Long
     
'**************************************************************************************************
' mdlDeclares Win32 Module-Scoped API
'**************************************************************************************************
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
     ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'**************************************************************************************************
' mdlDeclares Module-Level Variables
'**************************************************************************************************
Private m_hWnd() As Long
Private m_iCount As Long

Public m_PopMenuCreated As Boolean
Private m_sndData() As Byte

'**************************************************************************************************
' mdlDeclares Property Statements
'**************************************************************************************************
Public Property Get EnumerateWindowsCount() As Long
     EnumerateWindowsCount = m_iCount
End Property ' Get EnumerateWindowsCount

Public Property Get EnumerateWindowshWnd(ByVal iIndex As Long) As Long
     EnumerateWindowshWnd = m_hWnd(iIndex)
End Property ' Get EnumerateWindowshWnd

'**************************************************************************************************
' mdlDeclares cPopupMenu Utility Methods/Subs
'**************************************************************************************************
Public Function AutoStartAdd() As Long
     Dim hKey As Long
     Dim lRtn As Long
     Dim sPathApp As String
     ' get a key handle
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False Then
          ' construct path to app
          sPathApp = App.Path + Chr(92) + App.EXEName + ".exe"
          ' set the value
          lRtn = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, ByVal sPathApp, Len(sPathApp))
     End If
End Function ' AutoStartAdd

Public Function AutoStartDelete() As Long
     Dim hKey As Long
     Dim lRtn As Long
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     If lRtn = False Then AutoStartDelete = RegDeleteValue(hKey, App.EXEName)
End Function ' AutoStartDelete

Private Function ClassName(ByVal lhWnd As Long) As String
     Dim lLen As Long
     Dim sBuf As String
     lLen = 260
     sBuf = String$(lLen, 0)
     lLen = GetClassName(lhWnd, sBuf, lLen)
     If (lLen <> 0) Then ClassName = Left$(sBuf, lLen)
End Function ' ClassName

Public Function EnumerateWindows() As Long
     m_iCount = 0
     Erase m_hWnd
     EnumWindows AddressOf EnumWindowsProc, 0
End Function ' EnumerateWindows

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
     Dim sClass As String
     sClass = ClassName(hwnd)
     If sClass = "#32768" Then ' Menu Window Class Name
          If IsWindowVisible(hwnd) Then
               m_iCount = m_iCount + 1
               ReDim Preserve m_hWnd(1 To m_iCount) As Long
               m_hWnd(m_iCount) = hwnd
          End If
     End If
End Function ' EnumWindowsProc

Public Function IsAutoStart() As Boolean
     Dim hKey As Long
     Dim lType As Long
     Dim sValue As String
     ' set value
     sValue = App.EXEName
     ' If the key exists
     If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          0, KEY_READ, hKey) = False Then
          ' Look for the subkey named after the application
          If RegQueryValueEx(hKey, sValue, ByVal 0&, lType, ByVal 0&, ByVal 0&) = False Then
               IsAutoStart = True
               ' Close the registry key handle.
               RegCloseKey hKey
          End If
    End If
End Function ' IsAutoStart

Public Function ExistsRegEntries() As Boolean
     Dim hKey As Long
     Dim lType As Long
     Dim sValue As Long
     ' Set key were are looking for
     sValue = App.EXEName
     ' if the key exists
     If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Clockster\Clockster", _
          0, KEY_READ, hKey) = False Then
          ExistsRegEntries = True
          ' Close the handle
          RegCloseKey hKey
     End If
End Function ' ExistsRegEntries

Public Function GetRegValue(lhKey As Long, sSetting As String) As String
     Dim lRtn As Long
     Dim sBuff As String
     Dim lSize As Long
     ' Retrieve value from registry
     lRtn = RegQueryValueEx(lhKey, sSetting, 0&, REG_SZ, ByVal 0&, lSize)
     ' if successful
     If lRtn = False And lSize Then
          ' setup the buffer
          sBuff = String(lSize, Chr(0))
          ' receive the string
          lRtn = RegQueryValueEx(lhKey, sSetting, 0&, REG_SZ, ByVal sBuff, lSize)
          ' trim the null
          GetRegValue = Left$(sBuff, Len(sBuff) - 1)
     End If
End Function ' GetRegValue

Public Sub PlayResSound(lID As Integer, sSoundType As String)
    m_sndData = LoadResData(lID, sSoundType)
    sndPlaySound m_sndData(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub ' PlayResSound

Public Sub PlayWaveFile(ByVal FilePath As String)
     Dim lFlags As Long
     Dim lRtn As Long
     lFlags = SND_ASYNC Or SND_FILENAME
     lRtn = PlaySound(FilePath, 0&, lFlags)
End Sub ' PlayWaveFile

Public Function SetRegValue(lhKey As Long, sSetting As String, sValue As String, _
     Optional sDefault As String) As Long
     Dim sRegVal As String
     Dim lRtn As Long
     ' if value is present
     If Len(sValue) Then
          sRegVal = sValue
     Else ' use the default
          sRegVal = sDefault
     End If
     ' Call api
     SetRegValue = RegSetValueEx(lhKey, sSetting, 0&, REG_SZ, ByVal sRegVal, Len(sRegVal))
End Function ' SetRegValue
