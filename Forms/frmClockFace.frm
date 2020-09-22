VERSION 5.00
Begin VB.Form frmClockFace 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "ClocksterXP Face"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   5550
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClockFace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   Begin ClocksterXP.TouchSlider TouchSlider1 
      Height          =   120
      Left            =   0
      TabIndex        =   2
      Top             =   150
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   212
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00AEB388&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3555
      Left            =   5130
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   255
   End
   Begin ClocksterXP.CustomClock CustomClock1 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   344
   End
   Begin ClocksterXP.ClockRip ClockRip 
      Left            =   3510
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmClockFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  frmClockFace.frm - Base Clockster Interface
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
'  frmClockFace Win32 API Declares
'**************************************************************************************************
'  none or located in Winsubhook2.tlb

'**************************************************************************************************
'  frmClockFace Module-Scoped variables
'**************************************************************************************************
Implements WinSubHook2.iSubclass
Private m_sc As cSubclass
Private m_hMenuMain As Long
Private m_hMenuSub As Long
Private m_iMute As Long
Private WithEvents m_pop As cPopupMenu
Attribute m_pop.VB_VarHelpID = -1

'**************************************************************************************************
'  frmClockFace Intrinsic Events
'**************************************************************************************************
Private Sub Form_DblClick()
   Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub ' Form_DblClick

Private Sub Form_Load()
     ' set our window style to child
     ClockRip.SetWinStyle Me.hwnd, WS_CHILD, NO_STYLE
     ' set our parent
     ClockRip.AdoptClockParent Me.hwnd
     ' Create subclass object
     Set m_sc = New cSubclass
     ' add the message...here we only want the context menu message
     m_sc.AddMsg WM_CONTEXTMENU, MSG_BEFORE
     m_sc.AddMsg WM_SETTINGCHANGE, MSG_AFTER
     ' start subclassing
     m_sc.Subclass Me.hwnd, Me
     ' create popup menu
     Set m_pop = New cPopupMenu
     ' assign owner
     m_pop.hWndOwner = Me.hwnd
     ' Build menu
     Call menuCreate(m_pop)
End Sub ' Form_Load

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     CustomClock1.ToolTipText = Format$(Date, "Long Date")
End Sub ' Form_MouseMove

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim lIdx As Long
     If Button = vbRightButton Then
          m_pop.Restore "TrayVolume"
          lIdx = m_pop.ShowPopupMenu(0, 0)
     End If
End Sub ' Form_MouseUp

Private Sub Form_Unload(Cancel As Integer)
     On Error Resume Next
     ' unload options form just in case
     Unload frmOptions
     ' destroy it
     Set frmOptions = Nothing
     ' unload the tip form
     Unload frmTip
     ' destroy it
     Set frmTip = Nothing
     ' stop the subclasser
     m_sc.UnSubclass
     ' Destroy subclass object
     Set m_sc = Nothing
     ' can our parent
     ClockRip.DivorceClockParent Me.hwnd
End Sub ' Form_Unload

'**************************************************************************************************
'  frmClockFace Sited Control Events
'**************************************************************************************************
Private Sub ClockRip_OnAppBarPositionChange(tbPos As APPBAREDGE, ByVal lPosLeft As Long, _
     ByVal lPosTop As Long, ByVal lPosRight As Long, ByVal lPosBottom As Long, _
     lHeight As Long, lWidth As Long)
     ' now form position and width based on appbar location
     Select Case tbPos
          Case ABE_LEFT, ABE_RIGHT
               Move 0, 15
          Case ABE_TOP, ABE_BOTTOM
               Move 15, 0
     End Select
     ' Set position based on offsets
     CustomClock1.Move 0, CustomClock1.ClockOffsetY, lWidth, CustomClock1.Height
     TouchSlider1.Move 2, TouchSlider1.SliderOffsetY, lWidth - 4, _
          TouchSlider1.DefaultHeight + TouchSlider1.HeightAdjust
End Sub ' ClockRip_OnAppBarPositionChange

Private Sub ClockRip_OnBackgroundChange(backImg As stdole.Picture)
     Set Picture = ClockRip.Picture
     Refresh
End Sub ' ClockRip_OnBackgroundChange

Private Sub ClockRip_OnClockRectChange(ByVal lPosLeft As Long, ByVal lPosTop As Long, _
     ByVal lPosRight As Long, ByVal lPosBottom As Long, lHeight As Long, lWidth As Long)
     ' resize test box
     Move Me.Left, Me.Top, lWidth * Screen.TwipsPerPixelX, _
          lHeight * Screen.TwipsPerPixelY
     ' Set position based on offsets
     CustomClock1.Move 0, CustomClock1.ClockOffsetY, lWidth, CustomClock1.Height
     TouchSlider1.Move 2, TouchSlider1.SliderOffsetY, lWidth - 4, _
          TouchSlider1.DefaultHeight + TouchSlider1.HeightAdjust
End Sub ' ClockRip_OnClockRectChange

Private Sub CustomClock1_OnContextMenu()
     ' call the sub that already handles this
     Call Form_MouseUp(vbRightButton, 0, 0, 0)
End Sub ' CustomClock1_OnContextMenu

Private Sub CustomClock1_OnPositionChange(ByVal lOffsetY As Long)
     CustomClock1.Move CustomClock1.Left, CustomClock1.ClockOffsetY
End Sub ' CustomClock1_OnPositionChange

Private Sub TouchSlider1_OnContextMenu()
     ' call the sub that already handles this
     Call Form_MouseUp(vbRightButton, 0, 0, 0)
End Sub ' TouchSlider1_OnContextMenu

Private Sub TouchSlider1_OnMute(bValue As Boolean)
     ' keep the menu apprised of mute changes
     m_pop.Checked(m_iMute) = bValue
End Sub ' TouchSlider1_OnMute

Private Sub TouchSlider1_OnPositionChange(ByVal lOffsetY As Long, ByVal lHeightAdj As Long)
     Dim lTSHt As Long
     Dim lTSWt As Long
     lTSWt = TouchSlider1.Width
     lTSHt = TouchSlider1.DefaultHeight + lHeightAdj
     TouchSlider1.Move TouchSlider1.Left, TouchSlider1.SliderOffsetY, lTSWt, lTSHt
End Sub ' TouchSlider1_OnPositionChange

'**************************************************************************************************
'  frmClockFace Object Methods
'**************************************************************************************************
Private Sub m_pop_Click(ItemNumber As Long)
     ' process menu items
     Select Case ItemNumber
          Case 2 ' show about form
               frmAbout.Show
          ' case 4 through 7 are control panel/system applets
          Case 6
               Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
          Case 7
               Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,3")
          Case 8
               Call Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1")
          Case 9
               Call Shell("sndvol32.exe", vbNormalFocus)
          Case 10 ' show options dialog
               frmOptions.Show
          Case 11
               If m_pop.Checked(ItemNumber) Then
                    TouchSlider1.Mute = False
                    m_pop.Checked(ItemNumber) = False
               Else
                    TouchSlider1.Mute = True
                    m_pop.Checked(ItemNumber) = True
               End If
          Case 13 ' bail out
               Unload Me
     End Select
End Sub ' m_pop_Click

Private Sub m_pop_DrawItem(ByVal hDC As Long, ByVal lMenuIndex As Long, lLeft As Long, _
     lTop As Long, lRight As Long, lBottom As Long, ByVal bSelected As Boolean, _
     ByVal bChecked As Boolean, ByVal bDisabled As Boolean, bDoDefault As Boolean, _
     ByVal lhMenu As Long)
     Dim bRtn As Boolean
     Dim mLogo As cLogo
     Dim lLoop As Long
     Dim lHeight As Long
     Dim lSubHeight As Long
     ' Loop through the menu items and tally the total height
     For lLoop = 1 To m_pop.count
          ' Check if item is in the main menu:
          Select Case m_pop.hMenu(lLoop)
               Case m_hMenuMain
                    lHeight = lHeight + m_pop.MenuItemHeight(lLoop)
               Case m_hMenuSub
                    lSubHeight = lSubHeight + m_pop.MenuItemHeight(lLoop)
          End Select
     Next
     ' set the logo pic's height based on the menu we're dealing with
     Select Case lhMenu
          Case m_hMenuMain
               ' set the picturebox height to total menu items before blitting
               picSideBar.Height = lHeight
          Case m_hMenuSub
               ' set the picturebox height to total submenu items before blitting
               picSideBar.Height = lSubHeight
     End Select
     ' create the logo object and call the drawing methods
     Set mLogo = New cLogo
     ' set logo properties
     With mLogo
          .DrawingObject = picSideBar
          .StartColor = 11449224   ' the start gradient color
          .EndColor = 15001055     ' the end gradient color
          .Caption = "ClocksterXP" ' logo text
          .Draw                    ' call the class draw method
     End With
     ' We've painted the logo to the picbox so destroy the logo object
     Set mLogo = Nothing
     ' paint logo to the menu
     BitBlt hDC, lLeft, lTop, picSideBar.Width, lBottom - lTop, picSideBar.hDC, _
          0, lTop, vbSrcCopy
     ' account for the picture width when drawing
     lLeft = lLeft + picSideBar.Width + 2
     ' do default passed by ref so make sure is true
     bDoDefault = True
End Sub ' m_pop_DrawItem

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Dim wndRC As tRECT
     ' eat the context menu message
     If uMsg = WM_CONTEXTMENU Then
          bHandled = True
     ElseIf uMsg = WM_SETTINGCHANGE Then
          ' first get the client rectangle
          GetClientRect ClockRip.ClockhWnd, wndRC
          ' position clock based on offsets
          CustomClock1.Move 0, CustomClock1.ClockOffsetY, wndRC.Right, 12
          TouchSlider1.Move 2, TouchSlider1.SliderOffsetY, wndRC.Right - 4, _
               TouchSlider1.DefaultHeight + TouchSlider1.HeightAdjust
     End If
     DoEvents
End Sub ' iSubclass_Proc

Public Sub menuCreate(mPop As cPopupMenu)
     Dim bMuted As Boolean
     Dim lItm As Long
     Dim lSubItm As Long
     Dim lStart As Long
     Dim lEnd As Long
     Dim lLoop As Long
     Dim sStr As String * 30
     ' add our menu items
     With mPop
          ' clear previous
          .Clear
          lItm = .AddItem("-  About...")
          sStr = "About ClocksterXP"
          lItm = .AddItem(sStr)
          .OwnerDraw(lItm) = True
          ' add items to main menu
          lItm = .AddItem("-  Settings...")
          ' get the hMenu handle for the main menu.  We'll need this
          ' when we are logo drawing
          m_hMenuMain = mPop.hMenu(lItm)
          ' yep, all are owner drawn
          .OwnerDraw(lItm) = True
          ' add pivot for submenu
          sStr = "Related System Applets"
          lItm = .AddItem(sStr)
          .OwnerDraw(lItm) = True
               ' add our submenu header...indented to differentiate between
               ' main and sub menu
               lSubItm = .AddItem("-  Related System Applets...", , , lItm)
               ' get the hMenu handle for this submenu
               m_hMenuSub = mPop.hMenu(lSubItm)
               .OwnerDraw(lSubItm) = True
               ' add submenu items
               sStr = "Time/Date Settings"
               lSubItm = .AddItem(sStr, , , lItm)
               .OwnerDraw(lSubItm) = True
               sStr = "Regional Settings"
               lSubItm = .AddItem(sStr, , , lItm)
               .OwnerDraw(lSubItm) = True
               sStr = "Sounds And Multimedia"
               lSubItm = .AddItem(sStr, , , lItm)
               .OwnerDraw(lSubItm) = True
               sStr = "Volume Control Mixer"
               lSubItm = .AddItem(sStr, , , lItm)
               .OwnerDraw(lSubItm) = True
          ' Finished submenu, continue drawing main menu
          sStr = "ClocksterXP Settings"
          lItm = .AddItem(sStr)
          .OwnerDraw(lItm) = True
          ' continue main menu
          sStr = "Mute Volume"
          lItm = .AddItem(sStr)
          .OwnerDraw(lItm) = True
          .Checked(lItm) = TouchSlider1.Mute
          ' store the index for this menu item
          m_iMute = lItm
          lItm = .AddItem("-  Exit...")
          sStr = "Close ClocksterXP"
          lItm = .AddItem(sStr)
          .OwnerDraw(lItm) = True
          .Store "Clockster"
     End With
End Sub ' menuCreate
