VERSION 5.00
Begin VB.UserControl UpDown 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ToolboxBitmap   =   "ucUpDownBox.ctx":0000
   Begin VB.Timer tmr_Inc 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1425
      Top             =   135
   End
   Begin VB.TextBox txtValue 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' ucUpDownBox.ctl
'**************************************************************************************************
' User control:  ucUpDownBox.ctl
'                (Owner drawn version)
' Author:        Carles P.V.
' Submitted PSC:
' Dependencies:  None
' Last revision: 2003.12.10
'**************************************************************************************************
' LOG:
' - 2003.12.10:
'   Fixed up-down button rects. initialization
'   with odd font heights.
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' ucUpDownBox Constants
'**************************************************************************************************
Private Const DFC_SCROLL As Long = &H3
Private Const DFCS_SCROLLDOWN As Long = &H1
Private Const DFCS_PUSHED As Long = &H200
Private Const DFCS_SCROLLUP As Long = &H0
Private Const DFCS_INACTIVE As Long = &H100
Private Const SM_CXVSCROLL As Long = &H2
Private Const SM_SWAPBUTTON As Long = 23
Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Const SPI_GETKEYBOARDSPEED As Long = 10
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const VK_LBUTTON As Long = &H1
Private Const VK_RBUTTON As Long = &H2
Private Const VK_MBUTTON As Long = &H4
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28

'**************************************************************************************************
' ucUpDownBox Enums
'**************************************************************************************************
Public Enum ebBorderStyleConstants
     [None] = 0
     [3D] = 1
End Enum ' ebBorderStyleConstants

Private Enum eScrollDirCts
     eScrollUp = DFCS_SCROLLUP
     eScrollDn = DFCS_SCROLLDOWN
End Enum ' eScrollDirCts

'**************************************************************************************************
' ucUpDownBox Structs
'**************************************************************************************************
Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

'**************************************************************************************************
' ucUpDownBox API Declares
'**************************************************************************************************
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT2, _
     ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, _
     ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function PtInRect Lib "user32" (lpRect As RECT2, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, _
     ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
     ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, _
     ByVal fuWinIni As Long) As Long

'**************************************************************************************************
' ucUpDownBox Event Declares
'**************************************************************************************************
Public Event Change()
Public Event DownClick()
Public Event UpClick()

'**************************************************************************************************
' ucUpDownBox Module-Scoped Variables
'**************************************************************************************************
Private m_rButtonUp       As RECT2
Private m_rButtonDn       As RECT2
Private m_eButtonUpPushed As Boolean
Private m_eButtonDnPushed As Boolean


'**************************************************************************************************
' ucUpDownBox Property Variables
'**************************************************************************************************

Private m_Min   As Long
Private m_Max   As Long
Private m_Value As Long


Private m_lBarWidth       As Long
Private m_lKeyboardDelay  As Long
Private m_lKeyboardSpeed  As Long
Private m_bSwapButtons    As Boolean

'**************************************************************************************************
' ucUpDownBox Property Statements
'**************************************************************************************************
Public Property Get Alignment() As AlignmentConstants
     Alignment = txtValue.Alignment
End Property ' Get Alignment

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
     txtValue.Alignment = New_Alignment
End Property ' Let Alignment

Public Property Get BackColor() As OLE_COLOR
     BackColor = UserControl.BackColor
End Property ' Get BackColor

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
     UserControl.BackColor = New_BackColor
     txtValue.BackColor = New_BackColor
End Property ' Let BackColor

Public Property Get BorderStyle() As ebBorderStyleConstants
     BorderStyle = UserControl.BorderStyle
End Property ' BorderStyle

Public Property Let BorderStyle(ByVal New_BorderStyle As ebBorderStyleConstants)
     UserControl.BorderStyle = New_BorderStyle
End Property ' Let BorderStyle

Public Property Get Enabled() As Boolean
     Enabled = UserControl.Enabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     txtValue.Enabled = New_Enabled
     UserControl.Enabled = New_Enabled
     Call UserControl_Paint
End Property ' Let Enabled

Public Property Get Font() As Font
     Set Font = txtValue.Font
End Property ' Get Font

Public Property Set Font(ByVal New_Font As Font)
     Set txtValue.Font = New_Font
     Set UserControl.Font = New_Font
     UserControl_Resize
End Property ' Set Font

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = txtValue.ForeColor
End Property ' Get ForeColor

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
     txtValue.ForeColor = New_ForeColor
End Property ' Let ForeColor

Public Property Get Max() As Long
     Max = m_Max
End Property ' Get Max

Public Property Let Max(ByVal New_Max As Long)
     m_Max = New_Max
End Property ' Let Max

Public Property Get Min() As Long
     Min = m_Min
End Property ' Get Min

Public Property Let Min(ByVal New_Min As Long)
     m_Min = New_Min
End Property ' Let Min

Public Property Get Value() As Long
     Value = m_Value
End Property ' Get Value

Public Property Let Value(ByVal New_Value As Long)
     If (New_Value < m_Min) Then New_Value = m_Min
     If (New_Value > m_Max) Then New_Value = m_Max
     txtValue = New_Value
End Property ' Let Value

'**************************************************************************************************
' ucUpDownBox Private Methods
'**************************************************************************************************
Private Sub pvSelectContents()
     ' Select Text box contents
     txtValue.SelStart = 0
     txtValue.SelLength = Len(txtValue)
End Sub ' pvSelectContents

Private Sub pvDrawScrollButton(lpRect As RECT2, ByVal eScrollDir As eScrollDirCts, _
     ByVal bPushed As Boolean, Optional ByVal bDisabled As Boolean = 0)
     ' Draw scroll button
     Call DrawFrameControl(hDC, lpRect, DFC_SCROLL, eScrollDir Or -bPushed * DFCS_PUSHED Or _
          -bDisabled * DFCS_INACTIVE)
End Sub ' pvDrawScrollButton

'**************************************************************************************************
' ucUpDownBox Constinuent Control Events
'**************************************************************************************************
Private Sub tmr_Inc_Timer()
     ' First, check Text box contents
     If (Not IsNumeric(txtValue)) Then m_Value = 0
     ' Apply increment
     Select Case True
          Case m_eButtonUpPushed ' <+1>
               If (m_Value < m_Max) Then
                    m_Value = m_Value + 1
               Else
                    Exit Sub
               End If
          Case m_eButtonDnPushed '<-1>
               If (m_Value > m_Min) Then
                    m_Value = m_Value - 1
               Else
                    Exit Sub
               End If
     End Select
     ' Update Text box and select Text box contents
     txtValue.Text = m_Value
     Call pvSelectContents
     ' Change to repeat delay
     If (tmr_Inc.Interval = m_lKeyboardDelay) Then tmr_Inc.Interval = m_lKeyboardSpeed
     ' Check [up]/[Down] keys
     If (GetAsyncKeyState(VK_UP) = 0 And GetAsyncKeyState(VK_LBUTTON - m_bSwapButtons) = 0 And _
          m_eButtonUpPushed) Then _
          Call UserControl_MouseUp(vbLeftButton, 0, CSng(m_rButtonUp.x1), _
               CSng(m_rButtonUp.y1))
     If (GetAsyncKeyState(VK_DOWN) = 0 And GetAsyncKeyState(VK_LBUTTON - m_bSwapButtons) = 0 And _
          m_eButtonDnPushed) Then _
          Call UserControl_MouseUp(vbLeftButton, 0, CSng(m_rButtonDn.x1), _
               CSng(m_rButtonDn.y1))
     ' Raise <Change> event
     RaiseEvent Change
End Sub ' tmr_Inc_Timer

Private Sub txtValue_Change()
     Dim lOldValue As Long
     ' Check
     If (IsNumeric(txtValue)) Then
          ' Store old value
          lOldValue = m_Value
          ' Check Min/Max range
          If (txtValue >= m_Min And txtValue <= m_Max) Then m_Value = txtValue
          ' Changed [?]
          If (lOldValue <> m_Value) Then RaiseEvent Change
      Else
          ' Reset
          m_Value = 0
     End If
End Sub ' txtValue_Change

Private Sub txtValue_GotFocus()
     ' Select contents
     Call pvSelectContents
End Sub ' txtValue_GotFocus

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
     ' Key support
     Select Case KeyCode
          Case vbKeyUp
               KeyCode = 0
               Call UserControl_MouseDown(vbLeftButton, 0, _
                    CSng(m_rButtonUp.x1), CSng(m_rButtonUp.y1))
          Case vbKeyDown
               KeyCode = 0: Call UserControl_MouseDown(vbLeftButton, 0, _
                    CSng(m_rButtonDn.x1), CSng(m_rButtonDn.y1))
     End Select
End Sub ' txtValue_KeyDown

Private Sub txtValue_KeyPress(KeyAscii As Integer)
     ' Only numbers (allow [KeyBack] and [-])
     If ((KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45) Then KeyAscii = 0
End Sub ' txtValue_KeyPress

Private Sub txtValue_KeyUp(KeyCode As Integer, Shift As Integer)
     ' Key support
     Select Case KeyCode
          Case vbKeyUp
               KeyCode = 0: Call UserControl_MouseUp(vbLeftButton, 0, _
                    CSng(m_rButtonUp.x1), CSng(m_rButtonUp.y1))
          Case vbKeyDown: KeyCode = 0: Call UserControl_MouseUp(vbLeftButton, 0, _
               CSng(m_rButtonDn.x1), CSng(m_rButtonDn.y1))
     End Select
End Sub ' txtValue_KeyUp

Private Sub txtValue_LostFocus()
     ' Reset timer/buttons
     tmr_Inc.Enabled = False
     m_eButtonUpPushed = False
     m_eButtonDnPushed = False
     Call UserControl_Paint
End Sub ' txtValue_LostFocus

'**************************************************************************************************
' ucUpDownBox UserControl Intrinsic Events
'**************************************************************************************************
Private Sub UserControl_DblClick()
     If (GetAsyncKeyState(VK_RBUTTON + m_bSwapButtons) = 0 And _
          GetAsyncKeyState(VK_MBUTTON) = 0) Then
               ' Preserve second click
          Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
     End If
End Sub ' UserControl_DblClick

Private Sub UserControl_Initialize()
     Dim lStyle As Long
     ' Get system defaults (*)
     m_lBarWidth = GetSystemMetrics(SM_CXVSCROLL)
     m_bSwapButtons = GetSystemMetrics(SM_SWAPBUTTON)
     Call SystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, m_lKeyboardDelay, 0)
     Call SystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, m_lKeyboardSpeed, 0)
     m_lKeyboardDelay = 250 + 250 * m_lKeyboardDelay
     m_lKeyboardSpeed = 400 - 11.46 * m_lKeyboardSpeed
End Sub ' UserControl_Initialize

Private Sub UserControl_InitProperties()
     UserControl.BorderStyle = [3D]
     UserControl.BackColor = vbWindowBackground
     Set Font = Ambient.Font
     m_Min = 0
     m_Max = 100
     txtValue = m_Min
End Sub ' UserControl_InitProperties

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If (Button = vbLeftButton) Then
          Select Case True
               Case PtInRect(m_rButtonUp, x, y)
                    ' Button pushed
                    If (Not m_eButtonUpPushed Or m_eButtonDnPushed) Then
                         m_eButtonUpPushed = True
                         m_eButtonDnPushed = False
                         UserControl_Paint
                    End If
                    ' Turn on timer
                    tmr_Inc.Interval = m_lKeyboardDelay
                    tmr_Inc.Enabled = True
                    tmr_Inc_Timer
               Case PtInRect(m_rButtonDn, x, y)
                    ' Button pushed
                    If (Not m_eButtonDnPushed Or m_eButtonUpPushed) Then
                         m_eButtonDnPushed = True
                         m_eButtonUpPushed = False
                         UserControl_Paint
                    End If
                    ' Turn on timer
                    tmr_Inc.Interval = m_lKeyboardDelay
                    tmr_Inc.Enabled = True
                    tmr_Inc_Timer
          End Select
     End If
End Sub ' UserControl_MouseDown

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Select Case True
          Case m_eButtonUpPushed
               ' Turn off timer
               tmr_Inc.Enabled = False
               ' Button released
               m_eButtonUpPushed = False
               UserControl_Paint
               RaiseEvent UpClick
          Case m_eButtonDnPushed
               ' Turn off timer
               tmr_Inc.Enabled = False
               ' Button released
               m_eButtonDnPushed = False
               UserControl_Paint
               RaiseEvent DownClick
     End Select
End Sub ' UserControl_MouseUp

Private Sub UserControl_Paint()
     ' Draw scroll buttons
     Call pvDrawScrollButton(m_rButtonUp, eScrollUp, m_eButtonUpPushed, Not UserControl.Enabled)
     Call pvDrawScrollButton(m_rButtonDn, eScrollDn, m_eButtonDnPushed, Not UserControl.Enabled)
End Sub ' UserControl_Paint

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          UserControl.BackColor = .ReadProperty("BackColor", vbWindowBackground)
          UserControl.BorderStyle = .ReadProperty("BorderStyle", [3D])
          UserControl.Enabled = .ReadProperty("Enabled", True)
          Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
          txtValue.Alignment = .ReadProperty("Alignment", vbLeftJustify)
          txtValue.BackColor = .ReadProperty("BackColor", vbWindowBackground)
          txtValue.ForeColor = .ReadProperty("ForeColor", vbWindowText)
          Set txtValue.Font = .ReadProperty("Font", Ambient.Font)
          m_Min = .ReadProperty("Min", 0)
          m_Max = .ReadProperty("Max", 100)
          txtValue = .ReadProperty("Value", 0)
     End With
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     ' Adjust width
     If (ScaleWidth < 2 * m_lBarWidth) Then Width = (2 * m_lBarWidth + _
          (Width \ Screen.TwipsPerPixelX - ScaleWidth)) * Screen.TwipsPerPixelX
     ' Adjust height
     Height = ((TextHeight("") + 4) + (Height \ Screen.TwipsPerPixelY - ScaleHeight)) * _
          Screen.TwipsPerPixelY
     ' Relocate controls
     txtValue.Move 1, 1, ScaleWidth - m_lBarWidth - 2, ScaleHeight
     SetRect m_rButtonUp, ScaleWidth - m_lBarWidth, 0, ScaleWidth, ScaleHeight \ 2
     SetRect m_rButtonDn, ScaleWidth - m_lBarWidth, ScaleHeight \ 2 + (ScaleHeight Mod 2), _
          ScaleWidth, ScaleHeight
End Sub ' UserControl_Resize

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          .WriteProperty "Alignment", txtValue.Alignment, vbLeftJustify
          .WriteProperty "BackColor", txtValue.BackColor, vbWindowBackground
          .WriteProperty "BorderStyle", UserControl.BorderStyle, [3D]
          .WriteProperty "ForeColor", txtValue.ForeColor, vbWindowText
          .WriteProperty "Enabled", UserControl.Enabled, True
          .WriteProperty "Font", txtValue.Font, Ambient.Font
          .WriteProperty "Min", m_Min, 0
          .WriteProperty "Max", m_Max, 100
          .WriteProperty "Value", m_Value, 0
     End With
End Sub ' UserControl_WriteProperties
