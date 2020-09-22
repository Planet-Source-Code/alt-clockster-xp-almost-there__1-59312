VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     vbAccelerator PopupMenu Component
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     22 November 2002
'**************************************************************************************************
' Requires: pcMemDC.cls
'           pcNCMetrics.cls
'           pcStoreMenu.cls
'           cSubclass.cls
'**************************************************************************************************
' Copyright © 1998-2002 Steve McMahon for vbAccelerator
'**************************************************************************************************
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' cPopupMenu Constants
'**************************************************************************************************
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_LEFT = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const CLR_INVALID = -1
Private Const DT_LEFT = &H0
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_HILITE = &H80&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
Private Const MF_UNCHECKED = &H0&
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFT_RADIOCHECK = &H200&
Private Const MIIM_DATA = &H20&
Private Const MIIM_ID = &H2&
Private Const MIIM_STATE = &H1&
Private Const MIIM_SUBMENU = &H4&
Private Const MIIM_TYPE = &H10&
Private Const ODT_MENU = 1
Private Const OPAQUE = 2
Private Const TPM_NOANIMATION = &H4000&
Private Const TPM_RETURNCMD = &H100
Private Const TPM_VERTICAL = &H40
Private Const TRANSPARENT = 1

'**************************************************************************************************
' cPopupMenu Enums/Structs
'**************************************************************************************************
Private Type DRAWITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemAction As Long
     itemState As Long
     hwndItem As Long
     hDC As Long
     RcItem As tRECT
     ItemData As Long
End Type ' DRAWITEMSTRUCT

Private Type MEASUREITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemWidth As Long
     itemHeight As Long
     ItemData As Long
End Type ' MEASUREITEMSTRUCT

Private Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long
     cch As Long
End Type ' MENUITEMINFO

Private Type tMenuWindowHandle ' Window/handles
     hMenu As Long
     hWnd As Long
End Type ' tMenuWindowHandle

Private Type TPMPARAMS
     cbSize As Long
     rcExclude As tRECT
End Type ' TPMPARAMS

'**************************************************************************************************
' cPopupMenu Win32 API
'**************************************************************************************************
Private Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, _
     ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, _
     ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
     ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, _
     lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As tRECT, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, _
     ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
     ByVal lpNewItem As Long) As Long
Private Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, _
     ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
     ByVal lpString As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As tRECT, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
     ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" ( _
     ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, _
     lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
     ByVal crColor As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, _
     ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As tRECT) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, _
     ByVal n1 As Long, ByVal n2 As Long, ByVal hWnd As Long, lpTPMParams As TPMPARAMS) As Long
Private Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
' Remaining API declared in mdlDeclares/Winsubhook2.tlb
     
'**************************************************************************************************
' cPopupMenu Events
'**************************************************************************************************
Public Event Click(ItemNumber As Long)
Public Event DrawItem(ByVal hDC As Long, ByVal lMenuIndex As Long, ByRef lLeft As Long, _
     ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long, _
     ByVal bSelected As Boolean, ByVal bChecked As Boolean, ByVal bDisabled As Boolean, _
     bDoDefault As Boolean, ByVal lhMenu As Long)
Public Event InitPopupMenu(ParentItemNumber As Long)
Public Event ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Public Event MeasureItem(ByVal lMenuIndex As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Public Event MenuExit()
Public Event UnInitPopupMenu(ParentItemNumber As Long)

'**************************************************************************************************
' cPopupMenu Implements
'**************************************************************************************************
Implements WinSubHook2.iSubclass

'**************************************************************************************************
' cPopupMenu Module-Level Variables
'**************************************************************************************************
Private m_cMemDC As pcMemDC
Private m_cNCM As New pcNCMetrics
Private m_cStoredMenu() As pcStoreMenu
Private m_fntSymbol As StdFont
Private m_hSubMenus() As Long
Private m_hWndAttached As Long
Private m_iStoreCount As Long
Private m_iWndCount As Long
Private m_lLastMaxId As Long
Private m_lMenuItemHeight As Long
Private m_lSubMenuCount As Long
Private m_sc As cSubclass
Private m_tMI() As tMenuItem ' Array of menu items
Private m_tWnd() As tMenuWindowHandle

'**************************************************************************************************
' cPopupMenu Property Variables
'**************************************************************************************************
Private m_ActiveMenuBackColor As Long
Private m_ActiveMenuColor As Long
Private m_CurrentlyRestoredKey As String
Private m_fnt As StdFont
Private m_hWndOwner As Long
Private m_InactiveMenuColor As Long
Private m_MenuCount As Long
Private m_MenuBackgroundColor As Long

'**************************************************************************************************
' cPopupMenu Public Property Statements
'**************************************************************************************************
Public Property Get ActiveMenuBackgroundColor() As OLE_COLOR
     If m_ActiveMenuBackColor = CLR_INVALID Then
          ActiveMenuBackgroundColor = vbHighlight
     Else
          ActiveMenuBackgroundColor = m_ActiveMenuBackColor
     End If
End Property ' Get ActiveMenuBackgroundColor

Public Property Let ActiveMenuBackgroundColor(ByVal New_ActiveMenuBackgroundColor As OLE_COLOR)
     m_ActiveMenuBackColor = New_ActiveMenuBackgroundColor
End Property ' Let ActiveMenuBackgroundColor

Public Property Get ActiveMenuForeColor() As OLE_COLOR
     If m_ActiveMenuColor = CLR_INVALID Then
          ActiveMenuForeColor = vbHighlightText
     Else
          ActiveMenuForeColor = m_ActiveMenuColor
     End If
End Property ' Get ActiveMenuForeColor

Public Property Let ActiveMenuForeColor(ByVal New_ActiveMenuForeColor As OLE_COLOR)
     m_ActiveMenuColor = New_ActiveMenuForeColor
End Property ' Let ActiveMenuForeColor

Public Property Get Caption(ByVal lIndex As Long) As String
     Caption = m_tMI(lIndex).sCaption
End Property ' Get Caption

Public Property Let Caption(ByVal lIndex As Long, ByVal New_Caption As String)
   m_tMI(lIndex).sInputCaption = New_Caption
   m_tMI(lIndex).sCaption = New_Caption
   ResetMenuForRecalc lIndex
End Property ' Let Caption

Public Property Get Checked(ByVal lIndex As Long) As Boolean
     Dim tMII As MENUITEMINFO
     tMII.fMask = MIIM_STATE
     tMII.cbSize = LenB(tMII)
     GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     m_tMI(lIndex).bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
     Checked = m_tMI(lIndex).bChecked
End Property ' Get Checked

Public Property Let Checked(ByVal lIndex As Long, ByVal New_Checked As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     m_tMI(lIndex).bChecked = New_Checked
     If (New_Checked) Then
          lFlag = MF_CHECKED
          lFlagNot = 0
     Else
          lFlag = 0
          lFlagNot = MF_CHECKED
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let Checked

Public Property Get count() As Long
     count = m_MenuCount
End Property ' Get Count

Public Property Get CurrentlyRestoredKey() As String
     CurrentlyRestoredKey = m_CurrentlyRestoredKey
End Property ' Get CurrentlyRestoredKey

Public Property Get Default(ByVal lIndex As Long) As Boolean
     Default = m_tMI(lIndex).bDefault
End Property ' Get Default

Public Property Let Default(ByVal lIndex As Long, ByVal New_Default As Boolean)
     m_tMI(lIndex).bDefault = New_Default
     ' Modify this menu item to reflect the new properties:
     ResetMenuForRecalc lIndex
End Property ' Let Default

Public Property Get Enabled(ByVal lIndex As Long) As Boolean
     Enabled = m_tMI(lIndex).bEnabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal lIndex As Long, ByVal New_Enabled As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     m_tMI(lIndex).bEnabled = New_Enabled
     If (New_Enabled) Then
          lFlag = MF_ENABLED
          lFlagNot = MF_GRAYED
     Else
          lFlag = MF_DISABLED
          lFlagNot = MF_GRAYED
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let Enabled

Public Property Get Font() As IFont
     Dim lHDC As Long
     If m_fnt Is Nothing Then
          lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
          Set Font = m_cNCM.Font(lHDC, MenuFOnt)
          DeleteDC lHDC
     Else
          Set Font = m_fnt
     End If
End Property ' Get Font

Public Property Let Font(fntThis As IFont)
     pSetFont fntThis
End Property ' Let Font

Public Property Set Font(fntThis As IFont)
     pSetFont fntThis
End Property ' Set Font

Public Property Get Header(ByVal lIndex As Long) As Boolean
     Header = m_tMI(lIndex).bTitle
End Property ' Get Header

Public Property Let Header(ByVal lIndex As Long, ByVal New_Header As Boolean)
     m_tMI(lIndex).bTitle = New_Header
     If (New_Header) Then
          If Not m_tMI(lIndex).bDragOff Then
               ' disable the item
               pSetMenuFlag lIndex, MF_DISABLED, MF_GRAYED
          Else
               ' neither disabled nor grayed
               pSetMenuFlag lIndex, 0, MF_DISABLED Or MF_GRAYED
          End If
     End If
End Property ' Let Header

Public Property Get hMenu(ByVal lIndex As Long) As Long
     hMenu = m_tMI(lIndex).hMenu
End Property ' Get hMenu

Public Property Get hWndOwner() As Long
     hWndOwner = m_hWndOwner
End Property ' Get hWndOwner

Public Property Let hWndOwner(ByVal New_hWndOwner As Long)
     ' Clear up:
     Clear
     ' Set for new owner:
     m_hWndOwner = New_hWndOwner
End Property ' Let hWndOwner

Public Property Get IDForItem(ByVal lIndex As Long) As Long
     If lIndex > 0 And lIndex <= m_MenuCount Then IDForItem = m_tMI(lIndex).lActualID
End Property ' Get IDForItem

Public Property Get InactiveMenuForeColor() As OLE_COLOR
     If m_InactiveMenuColor = CLR_INVALID Then
          InactiveMenuForeColor = vbMenuText
     Else
          InactiveMenuForeColor = m_InactiveMenuColor
     End If
End Property ' Get InactiveMenuForeColor

Public Property Let InactiveMenuForeColor(ByVal New_InactiveMenuForeColor As OLE_COLOR)
     m_InactiveMenuColor = New_InactiveMenuForeColor
End Property ' Let InactiveMenuForeColor

Public Property Get IndexForKey(ByVal sKey As String) As Long
     Dim i As Long
     i = plMenuIndex(sKey)
     If i = -1 Then i = 0
     IndexForKey = i
End Property ' Get IndexForKey

Public Property Get ItemData(ByVal lIndex As Long) As Long
     ItemData = m_tMI(lIndex).lItemData
End Property ' Get ItemData

Public Property Let ItemData(ByVal lIndex As Long, ByVal lItemData As Long)
     Dim tMII As MENUITEMINFO
     m_tMI(lIndex).lItemData = lItemData
     If (lIndex > 0) And (lIndex <= m_MenuCount) Then
          tMII.cbSize = Len(tMII)
          tMII.fMask = MIIM_DATA
          GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
          tMII.dwItemData = lItemData
          SetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     End If
End Property ' Let ItemData

Public Property Get ItemForID(ByVal wID As Long) As Long
     Dim lIndex As Long
     For lIndex = 1 To m_MenuCount
          If m_tMI(lIndex).lActualID = wID Then
               ItemForID = lIndex
               Exit For
          End If
     Next
End Property ' Get ItemForID

Public Property Get ItemKey(ByVal lIndex As Long) As String
     ItemKey = m_tMI(lIndex).sKey
End Property ' Get ItemKey

Public Property Let ItemKey(ByVal lIndex As Long, ByVal sKey As String)
     m_tMI(lIndex).sKey = sKey
End Property ' Let ItemKey

Public Property Get ItemParentIndex(ByVal lIndex As Long) As Long
     ItemParentIndex = m_tMI(lIndex).lParentIndex
End Property ' Get ItemParentIndex

Public Property Get MenuBackgroundColor() As OLE_COLOR
     If m_MenuBackgroundColor = CLR_INVALID Then
          MenuBackgroundColor = vbWindowBackground
     Else
          MenuBackgroundColor = m_MenuBackgroundColor
     End If
End Property ' Get MenuBackgroundColor

Public Property Let MenuBackgroundColor(ByVal New_MenuBackgroundColor As OLE_COLOR)
     m_MenuBackgroundColor = New_MenuBackgroundColor
End Property ' Let MenuBackgroundColor

Public Property Get MenuItemHeight(ByVal lIndex As Long) As Long
     MenuItemHeight = m_tMI(lIndex).lHeight
End Property ' Get MenuItemHeight

Public Property Get MenuItemWidth(ByVal lIndex As Long) As Long
     MenuItemWidth = m_tMI(lIndex).lWidth
End Property ' Get MenuItemWidth

Public Property Get OwnerDraw(ByVal lIndex As Long) As Boolean
     OwnerDraw = m_tMI(lIndex).bOwnerDraw
End Property ' Get OwnerDraw

Public Property Let OwnerDraw(ByVal lIndex As Long, ByVal New_OwnerDraw As Boolean)
     m_tMI(lIndex).bOwnerDraw = New_OwnerDraw
End Property ' Let OwnerDraw

Public Property Get RadioCheck(ByVal lIndex As Long) As Boolean
     If (lIndex > 0) Then RadioCheck = m_tMI(lIndex).bRadioCheck
End Property ' Get RadioCheck

Public Property Let RadioCheck(ByVal lIndex As Long, ByVal New_RadioCheck As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     If New_RadioCheck Then _
          If m_tMI(lIndex).bChecked Then Checked(lIndex) = False
     m_tMI(lIndex).bRadioCheck = New_RadioCheck
     If (New_RadioCheck) Then
          lFlag = MFT_RADIOCHECK
          lFlagNot = 0
     Else
          lFlag = 0
          lFlagNot = MFT_RADIOCHECK
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let RadioCheck

Public Property Get Visible(ByVal lIndex As Long) As Boolean
     Visible = m_tMI(lIndex).bVisible
End Property ' Get Visible

Public Property Let Visible(ByVal lIndex As Long, ByVal New_Visible As Boolean)
     m_tMI(lIndex).bVisible = New_Visible
     ResetMenuForRecalc lIndex
End Property ' Let Visible

'**************************************************************************************************
' cPopupMenu Private Property Statements
'**************************************************************************************************
Private Property Get hFont() As Long
     Dim iFn As IFont
     Set iFn = Font
     hFont = iFn.hFont
End Property ' hFont

Private Property Get hFontBold() As Long
     Dim iFn As IFont
     Dim iFn2 As IFont
     Set iFn = Font
     iFn.Clone iFn2
     iFn2.Bold = True
     hFontBold = iFn2.hFont
End Property ' hFontBold

Private Property Get hFontSymbol() As Long
     Dim iFn As IFont
     Set iFn = m_fntSymbol
     hFontSymbol = iFn.hFont
End Property ' Get hFontSymbol

Private Property Get plStored(ByVal sKey As String) As Long
     Dim i As Long
     For i = 1 To m_iStoreCount
          If (m_cStoredMenu(i).Key = sKey) Then
               plStored = i
               Exit For
          End If
     Next
End Property ' Get plStored

'**************************************************************************************************
' cPopupMenu Public Methods
'**************************************************************************************************
Public Function AddItem(ByVal sCaption As String, Optional ByVal sHelptext As String = "", _
     Optional ByVal lItemData As Long = 0, Optional ByVal lParentIndex As Long = 0, _
     Optional ByVal bChecked As Boolean = False, Optional ByVal bEnabled As Boolean = True, _
     Optional ByVal sKey As String = "") As Long
     Dim lID As Long
     m_MenuCount = m_MenuCount + 1
     ReDim Preserve m_tMI(1 To m_MenuCount) As tMenuItem
     lID = plGetNewID()
     With m_tMI(m_MenuCount)
          .lID = lID
          .lActualID = lID
           pSetMenuCaption m_MenuCount, sCaption, (sCaption = "-")
          .sHelptext = sHelptext
           If (lParentIndex <> 0) Then .lParentId = m_tMI(lParentIndex).lActualID
          .lParentIndex = lParentIndex
          .lItemData = lItemData
          .bChecked = bChecked
          .bEnabled = bEnabled
          .bCreated = True
          .bVisible = True
          .bComboBox = False
          .bTextBox = False
          .bDragOff = False
          .bShowCheckAndIcon = False
          .sKey = sKey
     End With
     pAddNewMenuItem m_tMI(m_MenuCount)
     If (m_tMI(m_MenuCount).bTitle) Then
          Header(m_MenuCount) = True
          m_tMI(m_MenuCount).lHeight = (m_lMenuItemHeight + 6) * 3 \ 4
     Else
          If (m_tMI(m_MenuCount).sCaption = "-") Then
               m_tMI(m_MenuCount).lHeight = 6
          Else
               m_tMI(m_MenuCount).lHeight = m_lMenuItemHeight + 6
          End If
     End If
     ItemData(m_MenuCount) = lItemData
     AddItem = m_MenuCount
End Function ' AddItem

Public Sub Clear()
     Dim lMenu As Long
     m_lLastMaxId = &H800
     ' Clear up all submenus we have created:
     For lMenu = 1 To m_lSubMenuCount
          DestroyMenu m_hSubMenus(lMenu)
     Next
     m_lSubMenuCount = 0
     Erase m_hSubMenus
     ' Clear the main menu we have created:
     If (m_MenuCount > 0) Then DestroyMenu m_tMI(1).hMenu
     ' Clear up the array:
     m_MenuCount = 0
     Erase m_tMI
End Sub ' Clear

Public Sub CreateSubClass(hWndA As Long)
     If m_hWndAttached = hWndA Then
     Else
          With m_sc
                DestroySubClass
                .AddMsg WM_COMMAND, MSG_BEFORE
                .AddMsg WM_DESTROY, MSG_BEFORE
                .AddMsg WM_DRAWITEM, MSG_BEFORE
                .AddMsg WM_INITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_LBUTTONDOWN, MSG_BEFORE
                .AddMsg WM_MEASUREITEM, MSG_BEFORE
'                .AddMsg WM_MENUCHAR, MSG_BEFORE
                .AddMsg WM_MENUSELECT, MSG_BEFORE
                .AddMsg WM_MENURBUTTONUP, MSG_BEFORE
                .AddMsg WM_UNINITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_WININICHANGE, MSG_BEFORE
                .Subclass hWndA, Me
                m_hWndAttached = hWndA
          End With
     End If
End Sub ' CreateSubClass

Public Sub DestroySubClass()
     If Not (m_hWndAttached = 0) Then
          m_sc.UnSubclass
          m_hWndAttached = 0
     End If
End Sub ' DestroySubClass

Public Sub Restore(ByVal sKey As String)
     Dim lIndex As Long
     ' Restore the menu from the key sKey:
     If (sKey <> m_CurrentlyRestoredKey) Then
          lIndex = plStored(sKey)
          If (lIndex > 0) Then
               ' Clear any menu:
               Clear
                m_CurrentlyRestoredKey = ""
               ' Restore from storage:
               m_cStoredMenu(lIndex).Restore Me
               m_CurrentlyRestoredKey = sKey
          End If
     End If
End Sub ' Restore

Public Function ShowPopupMenu(ByVal lLeft As Long, ByVal lTop As Long, _
     Optional ByVal lExcludeLeft As Long = 0, Optional ByVal lExcludeTop As Long = 0, _
     Optional ByVal lExcludeRight As Long = 0, Optional ByVal lExcludeBottom As Long = 0, _
     Optional ByVal bTryToKeepLeft As Boolean = True) As Long
     Dim tP As tPOINT
     Dim tP2 As tPOINT
     Dim tPM As TPMPARAMS
     Dim lR As Long
     Dim lIndex As Long
     Dim lUN As Long
     Dim bIsSubclass As Boolean
     If count = 0 Then Exit Function
     tP.x = lLeft \ Screen.TwipsPerPixelX
     tP.y = lTop \ Screen.TwipsPerPixelY
     ClientToScreen m_hWndOwner, tP
     If Abs(lExcludeLeft - lExcludeRight) > 0 Or Abs(lExcludeTop - lExcludeBottom) > 0 Then
          tP2.x = lExcludeLeft \ Screen.TwipsPerPixelX
          tP2.y = lExcludeTop \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Left = tP2.x
          tPM.rcExclude.Top = tP2.y
          tP2.x = lExcludeRight \ Screen.TwipsPerPixelX
          tP2.y = lExcludeBottom \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Right = tP2.x
          tPM.rcExclude.Bottom = tP2.y
     End If
     tPM.cbSize = Len(tPM)
     lUN = TPM_RETURNCMD
     If Not (bTryToKeepLeft) Then lUN = lUN Or TPM_VERTICAL
     SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
     bIsSubclass = (m_hWndAttached <> 0)
     CreateSubClass m_hWndOwner
     lR = TrackPopupMenuEx(m_tMI(1).hMenu, lUN, tP.x, tP.y, m_hWndOwner, tPM)
     ' Find the index of the item with id lR within the menu:
     If lR > 0 Then
          lIndex = ItemForID(lR)
          ShowPopupMenu = lIndex
     End If
     If Not bIsSubclass Then
          If lR > 0 Then raiseClickEventSub lIndex
          DestroySubClass
     End If
     removeWindowHandles
     SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0
   ' The WM_COMMAND message is sent after this sub exits.
End Function ' ShowPopupMenu

Public Sub Store(ByVal sKey As String)
     Dim lIndex As Long
     Dim i As Long
     m_CurrentlyRestoredKey = ""
     ' Save the menu under the key sKey:
     lIndex = plStored(sKey)
     If (lIndex = 0) Then
          ' We need a new item
          m_iStoreCount = m_iStoreCount + 1
          ReDim Preserve m_cStoredMenu(1 To m_iStoreCount) As pcStoreMenu
          Set m_cStoredMenu(m_iStoreCount) = New pcStoreMenu
          lIndex = m_iStoreCount
     End If
     With m_cStoredMenu(lIndex)
          .Key = sKey
          .Store m_tMI(), m_MenuCount
     End With
     m_CurrentlyRestoredKey = sKey
End Sub ' Store

'**************************************************************************************************
' cPopupMenu Private Methods
'**************************************************************************************************
Private Sub addWindowHandle(ByVal lHDC As Long, ByVal hMenu As Long)
     Dim lhWnd As Long
     Dim i As Long
     Dim j As Long
     Dim k As Long
     Dim bFound As Boolean
     ' Works for W2k
     lhWnd = WindowFromDC(lHDC)
     ' else
     If lhWnd = 0 Then
          EnumerateWindows
          For i = 1 To EnumerateWindowsCount
               k = EnumerateWindowshWnd(i)
               bFound = False
               For j = 1 To m_iWndCount
                    If m_tWnd(j).hWnd = k Then bFound = True
               Next
               If Not bFound Then
                    lhWnd = k
                    Exit For
               End If
          Next
     End If
     If Not (lhWnd = 0) Then
          For i = 1 To m_iWndCount
               If m_tWnd(i).hMenu = hMenu Then Exit Sub
          Next
          ReDim Preserve m_tWnd(1 To m_iWndCount + 1) As tMenuWindowHandle
          m_iWndCount = m_iWndCount + 1
          With m_tWnd(m_iWndCount)
               .hMenu = hMenu
               .hWnd = lhWnd
          End With
     Else
      ' No handle for dc
     End If
End Sub ' addWindowHandle

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR) As Long
     Dim lCFrom As Long
     Dim lCTo As Long
     Dim lCRetR As Long
     Dim lCRetG As Long
     Dim lCRetB As Long
     lCFrom = TranslateColor(oColorFrom)
     lCTo = TranslateColor(oColorTo)
     lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
     If (lCRetR > 255) Then lCRetR = 255 Else If (lCRetR < 0) Then lCRetR = 0
     lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - _
          ((lCFrom \ &H100) And &HFF&)) \ 2
     If (lCRetG > 255) Then lCRetG = 255 Else If (lCRetG < 0) Then lCRetG = 0
     lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - _
          ((lCFrom \ &H10000) And &HFF&)) \ 2
     If (lCRetB > 255) Then lCRetB = 255 Else If (lCRetB < 0) Then lCRetB = 0
     BlendColor = RGB(lCRetR, lCRetG, lCRetB)
End Function ' BlendColor

Private Function DrawCustomEdge(ByVal hDC As Long, qrc As tRECT, ByVal edge As Long, _
     ByVal grfFlags As Long) As Long
     Dim junk As tPOINT
     Dim hPenOld As Long
     Dim hPen As Long
     If (qrc.Bottom > qrc.Top) Then
          hPen = CreatePen(PS_SOLID, 1, TranslateColor(&H808000)) ''vbHighlight))
     Else
          hPen = CreatePen(PS_SOLID, 1, TranslateColor(vb3DShadow))
     End If
     hPenOld = SelectObject(hDC, hPen)
     MoveToEx hDC, qrc.Left, qrc.Top, junk
     LineTo hDC, qrc.Right - 1, qrc.Top
     If (qrc.Bottom > qrc.Top) Then
          LineTo hDC, qrc.Right - 1, qrc.Bottom - 1
          LineTo hDC, qrc.Left, qrc.Bottom - 1
          LineTo hDC, qrc.Left, qrc.Top
     End If
     SelectObject hDC, hPenOld
     DeleteObject hPen
End Function ' DrawEdge

Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim bChecked As Boolean
     Dim bDefault As Boolean
     Dim bDisabled As Boolean
     Dim bDoDefault As Boolean
     Dim bHeader As Boolean
     Dim bHighlighted As Boolean
     Dim bRadioCheck As Boolean
     Dim bSeparator As Boolean
     Dim hBrush As Long
     Dim hFontInt As Long
     Dim hFntOld As Long
     Dim hFntSymOld As Long
     Dim lH As Long
     Dim lHDC As Long
     Dim lID As Long
     Dim lIndex As Long
     Dim lSelLeft As Long
     Dim lW As Long
     Dim tDIS As DRAWITEMSTRUCT
     Dim tMII As MENUITEMINFO
     Dim TR As tRECT
     Dim tSideRect As tRECT
     Dim tTR As tRECT
     Dim tWR As tRECT
     CopyMemory tDIS, ByVal lParam, Len(tDIS)
     If tDIS.CtlType = ODT_MENU Then
          lIndex = (plGetIndexForId(tDIS.itemID))
          If (lIndex > 0) Then
               If Not isVisible(lIndex) Then
                    DrawItem = True
                    Exit Function
               End If
               ' ensure the memory dc is big enough:
               m_cMemDC.Width = tDIS.RcItem.Right - tDIS.RcItem.Left + 2
               m_cMemDC.Height = tDIS.RcItem.Bottom - tDIS.RcItem.Top + 2
               lHDC = m_cMemDC.hDC
               LSet TR = tDIS.RcItem
               OffsetRect TR, -TR.Left, -TR.Top
               ' Get info about the menu item:
               tMII.cbSize = LenB(tMII)
               tMII.fMask = MIIM_TYPE Or MIIM_STATE 'Or MIIM_DATA
               ReDim B(0 To 128) As Byte
               tMII.dwTypeData = VarPtr(B(0))
               GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII
               bRadioCheck = m_tMI(lIndex).bRadioCheck
               bDisabled = Not (m_tMI(lIndex).bEnabled)
               bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED) Or bRadioCheck
               bHighlighted = ((tMII.fState And MFS_HILITE) = MFS_HILITE)
               bHeader = m_tMI(lIndex).bTitle
               bSeparator = isSeparator(lIndex)
               bDefault = m_tMI(lIndex).bDefault
               ' Fill background:
               TR.Bottom = TR.Bottom + 1
               TR.Right = TR.Right + 1
               LSet tTR = TR
               fillWithNormalBackground lHDC, tTR, tDIS.RcItem.Top
               LSet tSideRect = tTR
               tSideRect.Right = m_lMenuItemHeight '+ 8
               fillWithLighterControlColour lHDC, tSideRect, tDIS.RcItem.Top
               TR.Top = TR.Top + 1
               SetBkMode lHDC, TRANSPARENT
               ' set the appropriate font:
               If bDefault Then
                    hFntOld = SelectObject(lHDC, hFontBold)
               Else
                    hFntOld = SelectObject(lHDC, hFont)
               End If
               bDoDefault = True
               LSet tTR = tDIS.RcItem
               lW = tTR.Right - tTR.Left + 2
               lH = tTR.Bottom - tTR.Top + 1
               tTR.Top = tTR.Top - 1
               tTR.Bottom = tTR.Bottom + 1
               BitBlt tDIS.hDC, tDIS.RcItem.Left, tDIS.RcItem.Top, lW, lH, _
                    lHDC, 0, 0, vbSrcCopy
               RaiseEvent DrawItem(tDIS.hDC, lIndex, tTR.Left, tTR.Top, tTR.Right, _
                    tTR.Bottom, bHighlighted, bChecked, bDisabled, bDoDefault, m_tMI(lIndex).hMenu)
               BitBlt lHDC, 0, 0, lW, lH, tDIS.hDC, tDIS.RcItem.Left, _
                    tDIS.RcItem.Top, vbSrcCopy
               TR.Left = tTR.Left - tDIS.RcItem.Left
               TR.Top = tTR.Top - tDIS.RcItem.Top + 1
               ' ensure we have the window handle for the menu:
               addWindowHandle tDIS.hDC, m_tMI(lIndex).hMenu
               If (bDoDefault) Then
                    If Not (bSeparator) Or bHeader Then
                         ' Text item:
                         lID = tMII.dwItemData
                         If bChecked Then lSelLeft = TR.Left + 4 + (TR.Bottom - TR.Top + 1 - 4)
                         If bHighlighted And Not (bHeader Or bDisabled) Then
                              LSet tWR = TR
                              tWR.Left = tWR.Left + 1
                              tWR.Right = tWR.Right - 1
                              tWR.Bottom = tWR.Bottom - 1
                              fillWithLighterSelectedColour lHDC, tWR, _
                                   tDIS.RcItem.Top + tWR.Top
                              DrawCustomEdge lHDC, tWR, 0, 0
                         End If
                         If bDisabled Then
                              SetTextColor lHDC, TranslateColor(vb3DHighlight)
                         Else
                              If bHighlighted Then
                                   SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
                              Else
                                   SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                              End If
                         End If
                         LSet tWR = TR
                         If (bHeader) Then
                              ' no icons/checks
                         Else
                              tWR.Left = tWR.Left + 1
                              tWR.Right = tWR.Left + (tWR.Bottom - tWR.Top + 1 - 2)
                              ' Check:
                              If bChecked Then
                                   ' Colour in:
                                   If Not (bHighlighted) And Not (bDisabled) Then
                                        tWR.Top = tWR.Top + 1
                                        fillWithNormalBackground lHDC, tWR, _
                                             tDIS.RcItem.Top + tWR.Top
                                        tWR.Top = tWR.Top - 1
                                        If bDisabled Then _
                                             SetTextColor lHDC, TranslateColor(vb3DHighlight)
                                   End If
                                   If Not bDisabled Then _
                                        SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                                   tWR.Right = tWR.Right + 1
                                   tWR.Top = tWR.Top - 1
                                   tWR.Bottom = tWR.Bottom + 1
                                   ' Draw the appropriate symbol:
                                   hFontInt = SelectObject(lHDC, hFntOld)
                                   hFntSymOld = SelectObject(lHDC, hFontSymbol)
                                   If bRadioCheck Then
                                        pDrawText lHDC, "h", tWR, DT_VCENTER Or DT_CENTER Or _
                                             DT_SINGLELINE, bDisabled
                                   Else
                                        pDrawText lHDC, "b", tWR, DT_VCENTER Or DT_CENTER Or _
                                             DT_SINGLELINE, bDisabled
                                   End If
                                   SelectObject lHDC, hFntSymOld
                                   If bDefault Then
                                        hFntOld = SelectObject(lHDC, hFontBold)
                                   Else
                                        hFntOld = SelectObject(lHDC, hFont)
                                   End If
                              End If
                         End If
                         ' Draw text:
                         If bHeader Then
                              hFontInt = SelectObject(lHDC, hFntOld)
                              hFntSymOld = SelectObject(lHDC, m_cNCM.FontHandle(SMCaptionFont))
                              tWR.Left = tWR.Left - 1
                              tWR.Top = tWR.Top + 1
                              hBrush = CreateSolidBrush(TranslateColor(11449224))
                              SetTextColor lHDC, TranslateColor(&HFFFFFF)
                              FillRect lHDC, tWR, hBrush
                              DeleteObject hBrush
                              DrawText lHDC, m_tMI(lIndex).sCaption, -1, tWR, DT_SINGLELINE Or DT_TOP
                              SelectObject lHDC, hFntSymOld
                              hFntOld = SelectObject(lHDC, hFontInt)
                         Else
                              ' Not header
                              SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                              LSet tWR = TR
                              tWR.Left = TR.Left + 4 + (TR.Bottom - TR.Top + 1 - 4) + 2 + 1
                              tWR.Left = tWR.Left + 4
                              pDrawText lHDC, m_tMI(lIndex).sCaption, tWR, DT_LEFT Or _
                                   DT_SINGLELINE Or DT_VCENTER, bDisabled
                              If Len(m_tMI(lIndex).sShortCutDisplay) > 0 Then
                                   tWR.Left = tWR.Left + m_tMI(lIndex).lShortCutStartPos + 4 + 4
                                   pDrawText lHDC, m_tMI(lIndex).sShortCutDisplay, tWR, _
                                        DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
                              End If
                         End If
                    End If
                    SelectObject lHDC, hFntOld
                    If Not hFntOld = 0 Then SelectObject lHDC, hFntOld
                    DrawItem = True
               End If
               BitBlt tDIS.hDC, tDIS.RcItem.Left, tDIS.RcItem.Top, _
                    tDIS.RcItem.Right - tDIS.RcItem.Left + 1, _
                    tDIS.RcItem.Bottom - tDIS.RcItem.Top + 1, lHDC, _
                    0, 0, vbSrcCopy
          End If
     Else
          DrawItem = m_sc.CallOrigWndProc(WM_DRAWITEM, wParam, lParam)
     End If
End Function ' DrawItem

Private Sub fillWithLighterControlColour(ByVal lHDC As Long, TR As tRECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     SetBkMode lHDC, OPAQUE
     hBrush = CreateSolidBrush(BlendColor(9948620, MenuBackgroundColor))
     FillRect lHDC, TR, hBrush
     DeleteObject hBrush
     SetBkMode lHDC, TRANSPARENT
End Sub ' fillWithLighterControlColour

Private Sub fillWithLighterSelectedColour(ByVal lHDC As Long, TR As tRECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     SetBkMode lHDC, OPAQUE
     '11449224
     hBrush = CreateSolidBrush(BlendColor(11449224, MenuBackgroundColor))
     FillRect lHDC, TR, hBrush
     DeleteObject hBrush
   SetBkMode lHDC, TRANSPARENT
End Sub ' fillWithLighterSelectedColour

Private Sub fillWithNormalBackground(ByVal lHDC As Long, TR As tRECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     hBrush = CreateSolidBrush(TranslateColor(15001055)) ' MenuBackgroundColor))
     FillRect lHDC, TR, hBrush
     DeleteObject hBrush
End Sub ' fillWithNormalBackground

Private Function isSeparator(ByVal lIndex As Long) As Boolean
     If Trim$(m_tMI(lIndex).sCaption = "-") Then isSeparator = True
End Function ' isSeparator

Private Function isVisible(ByVal lIndex As Long) As Boolean
     If m_tMI(lIndex).bVisible Then
          isVisible = True
     Else
          isVisible = False
     End If
End Function ' isVisible

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim tMIS As MEASUREITEMSTRUCT
     Dim lIndex As Long
     Dim TR As tRECT
     Dim hFntOld As Long
     Dim hMenuThis As Long
     Dim l As Long
     Dim lID As Long
     Dim lMaxCaptionWidth As Long
     Dim lCaptionWidth As Long
     Dim lMaxShortCutWidth As Long
     Dim lMaxTotalWidth As Long
     Dim lTotalWidth As Long
     Dim lCount As Long
     Dim tMII As MENUITEMINFO
     CopyMemory tMIS, ByVal lParam, LenB(tMIS)
     If tMIS.CtlType = ODT_MENU Then
          lIndex = (plGetIndexForId(tMIS.itemID))
          If (lIndex <> 0) Then
               If Not isVisible(lIndex) Then
                    tMIS.itemHeight = 0
                    tMIS.itemWidth = 8
                    m_tMI(lIndex).lHeight = 0
                    m_tMI(lIndex).lWidth = 8
                    CopyMemory ByVal lParam, tMIS, LenB(tMIS)
                    Exit Function
               End If
               If Trim$(m_tMI(lIndex).sCaption = "-") Then
                    tMIS.itemWidth = 8
                    tMIS.itemHeight = 3
               Else
                    ' every item causes us to re-evaluate every other in the same menu...
                    hFntOld = SelectObject(m_cMemDC.hDC, hFont)
                    hMenuThis = m_tMI(lIndex).hMenu
                    lCount = GetMenuItemCount(hMenuThis)
                    ReDim lMenuIndex(1 To lCount) As Long
                    ReDim lShortCutWidth(1 To lCount) As Long
                    For l = 0 To lCount - 1
                         tMII.cbSize = Len(tMII)
                         tMII.fMask = MIIM_ID
                         lID = GetMenuItemInfo(hMenuThis, l, True, tMII)
                         lMenuIndex(l + 1) = (plGetIndexForId(tMII.wID))
                         If lMenuIndex(l + 1) > 0 Then
                              ' Get the width of this item:
                              If m_tMI(lMenuIndex(l + 1)).bDefault Then _
                                   hFntOld = SelectObject(m_cMemDC.hDC, hFontBold)
                              DrawText m_cMemDC.hDC, m_tMI(lMenuIndex(l + 1)).sCaption, _
                                   -1, TR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                              lCaptionWidth = TR.Right - TR.Left
                              ' add an additional space for the icon
                              If m_tMI(lMenuIndex(l + 1)).bShowCheckAndIcon Then _
                                   lCaptionWidth = lCaptionWidth + (m_lMenuItemHeight) + 4
                              If lCaptionWidth > lMaxCaptionWidth Then _
                                   lMaxCaptionWidth = lCaptionWidth
                              lCaptionWidth = lCaptionWidth + 4
                              lTotalWidth = lCaptionWidth
                              If Len(m_tMI(lMenuIndex(l + 1)).sShortCutDisplay) > 0 Then
                                   DrawText m_cMemDC.hDC, m_tMI(lMenuIndex(l + 1)).sShortCutDisplay, _
                                        -1, TR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                                   lShortCutWidth(l + 1) = TR.Right - TR.Left
                                   If lShortCutWidth(l + 1) > lMaxShortCutWidth Then _
                                        lMaxShortCutWidth = lShortCutWidth(l + 1)
                                   lTotalWidth = lTotalWidth + lShortCutWidth(l + 1)
                              End If
                              If lTotalWidth > lMaxTotalWidth Then lMaxTotalWidth = lTotalWidth
                              If m_tMI(lMenuIndex(l + 1)).bDefault Then _
                                   SelectObject m_cMemDC.hDC, hFntOld
                         End If
                    Next
                    SelectObject m_cMemDC.hDC, hFntOld
                    For l = 1 To lCount
                         If lMenuIndex(l) > 0 Then _
                              m_tMI(lMenuIndex(l)).lShortCutStartPos = lMaxTotalWidth - _
                                   lShortCutWidth(l) + 10 'lMaxCaptionWidt
                    Next
                    tMIS.itemWidth = 4 + m_lMenuItemHeight + 6 + lMaxTotalWidth + 4
                    If lMaxShortCutWidth > 0 Then tMIS.itemWidth = tMIS.itemWidth + 12
                    ' menu item height is always the same
                    tMIS.itemHeight = m_lMenuItemHeight + 6
                    If m_tMI(lIndex).bTitle Then tMIS.itemHeight = tMIS.itemHeight * 3 \ 4
               End If
               RaiseEvent MeasureItem(lIndex, tMIS.itemWidth, tMIS.itemHeight)
               m_tMI(lIndex).lWidth = tMIS.itemWidth
               m_tMI(lIndex).lHeight = tMIS.itemHeight
               CopyMemory ByVal lParam, tMIS, LenB(tMIS)
          End If
     Else
          MeasureItem = m_sc.CallOrigWndProc(WM_MEASUREITEM, wParam, lParam)
     End If
End Function ' MeasureItem

Private Sub pAddNewMenuItem(ByRef tMI As tMenuItem)
     Dim tMII As MENUITEMINFO
     Dim hMenu As Long
     Dim lPIndex As Long
     Dim lFlags As Long
     Dim lR As Long
     Dim hMenuNew As Long
     Dim bOk As Boolean
     Dim i As Long
     ' Find out where we're adding this item:
     With tMI
          If (.lParentId = 0) Then
               ' This is a new top level menu item:
               If (m_MenuCount = 1) Then
                    ' Create a new menu
                    hMenu = CreatePopupMenu()
               Else
                    ' Use the existing menu:
                    hMenu = m_tMI(1).hMenu
               End If
          Else
               ' We are adding to an existing menu:
               ' First we need to determine if there is already a sub menu for the parent item:
               lPIndex = plGetIndexForId(tMI.lParentId)
               If (lPIndex = 0) Then
                '  *** Couldn't find parent... *** "
               Else
                    ' Determine if the parent menu has a sub-menu:
                    tMII.fMask = MIIM_SUBMENU
                    tMII.cbSize = LenB(tMII)
                    GetMenuItemInfo m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, False, tMII
                    hMenu = tMII.hSubMenu
                    If (hMenu = 0) Then
                         ' We don't have a sub menu for this item so we're
                         ' going to have to add one:
                         ' Create the new menu item and store it's handle so we can clear up
                         ' again later:
                         hMenu = CreatePopupMenu()
                         If (hMenu <> 0) Then

                              ' Check if hMenu isn't an ID:
                              bOk = False
                              If Not (pbIDIsUnique(hMenu)) Then
                                   For i = 1 To 100
                                        hMenuNew = CreatePopupMenu()
                                        If (pbIDIsUnique(hMenuNew)) Then
                                             DestroyMenu hMenu
                                             hMenu = hMenuNew
                                             bOk = True
                                             Exit For
                                        Else
                                             DestroyMenu hMenuNew
                                        End If
                                   Next
                                   If Not bOk Then
                                        ' .. out of menu handles ...
                                        DestroyMenu hMenu
                                        Exit Sub
                                   End If
                              End If
                              m_lSubMenuCount = m_lSubMenuCount + 1
                              ReDim Preserve m_hSubMenus(1 To m_lSubMenuCount) As Long
                              m_hSubMenus(m_lSubMenuCount) = hMenu
                              ' Now set the parent item so it has a popup menu:
                              lFlags = plMenuFlags(lPIndex)
                              lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
                              lFlags = lFlags Or MF_POPUP
                              lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
                              lR = ModifyMenuByLong(m_tMI(lPIndex).hMenu, _
                                   m_tMI(lPIndex).lActualID, lFlags, hMenu, _
                                   m_tMI(lPIndex).lItemData)
                              ItemData(lPIndex) = m_tMI(lPIndex).lItemData
                                   ' WHen you add a sub menu to an item, its id becomes
                                   ' the sub menu handle
                                   m_tMI(lPIndex).lActualID = hMenu
                                   tMI.lParentId = hMenu
                              End If
                         End If
                    End If
               End If
               If (hMenu <> 0) Then
                    lFlags = plMenuFlags(m_MenuCount)
                    lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
                    lR = AppendMenuBylong(hMenu, lFlags, tMI.lID, tMI.lItemData)
               End If
          ' Store the hMenu for this item:
          .hMenu = hMenu
     End With
End Sub ' pAddNewMenuItem

Private Function pbIDIsUnique(ByVal lID As Long) As Boolean
     Dim bFound As Boolean
     Dim lMenu As Long
     For lMenu = 1 To m_MenuCount
          If (m_tMI(lMenu).lActualID = lID) Or (m_tMI(lMenu).lID = lID) Then
               bFound = True
               Exit For
          End If
     Next
     pbIDIsUnique = Not (bFound)
End Function ' pbIDIsUnique

Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, TR As tRECT, _
     ByVal dtFlags As Long, ByVal bDisabled As Boolean)
     Dim tWR As tRECT
     LSet tWR = TR
     If bDisabled Then SetTextColor lHDC, TranslateColor(vb3DShadow)
     DrawText lHDC, sText, -1, tWR, dtFlags
End Function ' pDrawText

Private Function plGetIndexForId(ByVal lItemId As Long) As Long
     Dim l As Long
     Dim lIndex As Long
     lIndex = 0
     For l = 1 To m_MenuCount
          If (m_tMI(l).lActualID = lItemId) Then
               lIndex = l
               Exit For
          End If
     Next
    plGetIndexForId = lIndex
End Function ' plGetIndexForId

Private Function plGetNewID() As Long
     Dim lID As Long
     If (m_lLastMaxId < m_MenuCount) Then
          m_lLastMaxId = m_MenuCount
     Else
          m_lLastMaxId = m_lLastMaxId + 1
     End If
     lID = m_lLastMaxId
     Do Until (pbIDIsUnique(lID))
          lID = lID + 1
          m_lLastMaxId = lID
     Loop
     plGetNewID = lID
End Function ' plGetNewID

Private Function plMenuFlags(ByVal lIndex As Long)
     Dim lFlags As Long
     With m_tMI(lIndex)
          If (.bChecked) Then
               lFlags = lFlags Or MF_CHECKED
          Else
               lFlags = lFlags Or MF_UNCHECKED
          End If
          If (.bEnabled) Then
               lFlags = lFlags Or MF_ENABLED
          Else
               lFlags = lFlags Or MF_GRAYED
          End If
          If Left$(Trim$(.sInputCaption), 1) = "-" Or (.bTitle And Not .bDragOff) Or _
               Not (isVisible(lIndex)) Then _
                    lFlags = lFlags Or MF_SEPARATOR
          If (m_tMI(lIndex).bMenuBarBreak) Then lFlags = lFlags Or MF_MENUBARBREAK
          If (m_tMI(lIndex).bMenuBreak) Then lFlags = lFlags Or MF_MENUBREAK
     End With
     plMenuFlags = lFlags
End Function ' plMenuFlags

Private Function plMenuIndex(ByVal vKey As Variant) As Long
     Dim i As Long
     ' Signal default
     plMenuIndex = -1
     ' Check for numeric key (i.e. index):
     If (IsNumeric(vKey)) Then
          i = CLng(vKey)
          If (i > 0) And (i <= m_MenuCount) Then plMenuIndex = i
     Else
          ' Check for string key:
          For i = 1 To m_MenuCount
               If (m_tMI(i).sKey = vKey) Then
                    plMenuIndex = i
                    Exit Function
               End If
          Next
     End If
End Function ' plMenuIndex

Private Sub pParseCaption(ByRef sCaption As String, ByVal sToken As String, ByRef bFlag As Boolean)
     Dim iPos As Long
     Dim iPos2 As Long
     Dim sCap As String
     iPos = InStr(sCaption, sToken)
     If (iPos <> 0) Then
          ' Check for double token (i.e. interpret as untokenised character):
          iPos2 = InStr(sCaption, sToken & sToken)
          If (iPos2 <> 0) Then
               bFlag = False
               If (iPos2 > 1) Then sCap = Left$(sCaption, iPos - 1)
               If (iPos2 + 1 < Len(sCaption)) Then sCap = sCap & Mid$(sCaption, iPos2 + 1)
          Else
               bFlag = True
               If (iPos > 1) Then sCap = Left$(sCaption, iPos - 1)
               If (iPos < Len(sCaption)) Then sCap = sCap & Mid$(sCaption, iPos + 1)
               sCaption = sCap
          End If
     Else
          bFlag = False
     End If
End Sub ' pParseCaption

Private Sub pSetFont(fntThis As IFont)
     Dim TR As tRECT
     Dim hFntOld As Long
     Dim i As Long
     Set m_fnt = fntThis
     m_fntSymbol.Size = Font.Size * 1.3
     hFntOld = SelectObject(m_cMemDC.hDC, hFont)
     DrawText m_cMemDC.hDC, "Xg", -1, TR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
     SelectObject m_cMemDC.hDC, hFntOld
     m_lMenuItemHeight = TR.Bottom - TR.Top + 2
     For i = 1 To m_MenuCount
          ResetMenuForRecalc i
     Next
End Sub ' pSetFont

Private Sub pSetMenuCaption(ByVal iItem As Long, ByVal sCaption As String, _
     ByVal bSeparator As Boolean)
     Dim sShortCut As String
     Dim iPos As Long
     m_tMI(iItem).sInputCaption = sCaption
     If (bSeparator) Then
          m_tMI(iItem).sCaption = "-"
     Else
          ' Check if this is a title:
          If (Left$(sCaption, 1) = "-") Then
               m_tMI(iItem).bTitle = True
               sCaption = Mid$(sCaption, 2)
          End If
          ' Check if this menu item will have a menu bar break:
          pParseCaption sCaption, "|", m_tMI(iItem).bMenuBarBreak
          ' Check if this menu item will be on the same line as
          ' the last one:
          pParseCaption sCaption, "^", m_tMI(iItem).bMenuBreak
          m_tMI(iItem).sCaption = sCaption
     End If
End Sub ' pSetMenuCaption

Private Sub pSetMenuFlag(ByVal lIndex As Long, ByVal lFlag As Long, ByVal lFlagNot As Long)
     Dim tMII As MENUITEMINFO
     Dim lFlags As Long
     lFlags = plMenuFlags(lIndex)
     lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
     tMII.fMask = MIIM_SUBMENU
     tMII.cbSize = LenB(tMII)
     GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     If (tMII.hSubMenu <> 0) Then lFlags = lFlags Or MF_POPUP
     lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
     lFlags = lFlags Or lFlag
     lFlags = lFlags And Not lFlagNot
     ModifyMenuByLong m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, lFlags, _
          m_tMI(lIndex).lActualID, m_tMI(lIndex).lItemData
End Sub ' pSetMenuFlag

Private Function RaiseClickEvent(lID As Long) As Boolean
     ' Return true from this if we have completely handled the
     ' click on our own:
     Dim lIndex As Long
     ' Find the Index of this menu id within our own array:
     lIndex = plGetIndexForId(lID)
     ' If we find it, then raise a click event for it:
     If (lIndex > 0) Then
          ' Send a click event with the index:
          raiseClickEventSub lIndex
          ' If this was one of the VB menu entries we have
          ' subclassed, we want to return false.  Then the
          ' click will filter through to the original Click
          ' event so your code should work as normal:
          On Error Resume Next ' 15/04/03: SPM: we may be unloading
          If Not (m_tMI(lIndex).bIsAVBMenu) Then RaiseClickEvent = True
     Else
          ' This is a problem.  We've got a click on
          ' a menu id which doesn't seem to be any
          ' of the menu items of the form.  It shouldn't
          ' happen, but return false anyway so we don't eat
          ' the message.
          RaiseClickEvent = False
   End If
End Function ' RaiseClickEvent

Private Sub raiseClickEventSub(ByVal lIndex As Long)
     RaiseEvent Click(lIndex)
End Sub ' raiseClickEventSub

Private Sub RaiseHighlightEvent(lID As Long)
     Dim lIndex As Long
     Dim sCaption As String
     Dim bSeparator As Boolean
     lIndex = plGetIndexForId(lID)
     If (lIndex > 0) Then
          sCaption = Trim$(m_tMI(lIndex).sCaption)
          If Len(sCaption) >= 1 Then _
               If Left$(sCaption, 1) = "-" Then bSeparator = True
         RaiseEvent ItemHighlight(lIndex, m_tMI(lIndex).bEnabled, bSeparator)
     End If
End Sub ' RaiseHighlightEvent

Private Sub RaiseInitMenuEvent(ByVal hMenu As Long, ByVal bState As Boolean)
     Dim lIndex As Long
     Dim lParentId As Long
     Dim bFound As Boolean
     ' Firstly, we need to find the index of an item in hMenu
     For lIndex = m_MenuCount To 1 Step -1
          If (m_tMI(lIndex).hMenu = hMenu) Then
               lParentId = m_tMI(lIndex).lParentId
               bFound = True
          End If
          If (bFound) Then
               If (m_tMI(lIndex).lActualID = lParentId) Then
                    If bState Then
                         RaiseEvent InitPopupMenu(lIndex)
                    Else
                         RaiseEvent UnInitPopupMenu(lIndex)
                    End If
                    Exit For
               End If
          End If
     Next
End Sub ' RaiseInitMenuEvent

Private Sub RaiseMenuExitEvent()
     ' raise the event:
     RaiseEvent MenuExit
End Sub ' RaiseMenuExitEvent

Private Sub removeWindowHandle(ByVal hMenu As Long)
     Dim i As Long
     Dim lIndex As Long
     For i = m_iWndCount To 1 Step -1
          If m_tWnd(i).hMenu = hMenu Then
               lIndex = i
               Exit For
          End If
     Next
     If lIndex > 0 Then
          If m_iWndCount > 1 Then
               For i = lIndex To m_iWndCount - 1
                    LSet m_tWnd(i) = m_tWnd(i + 1)
               Next
               m_iWndCount = m_iWndCount - 1
          Else
               m_iWndCount = 0
               Erase m_tWnd
          End If
     End If
End Sub ' removeWindowHandle

Private Sub removeWindowHandles()
     m_iWndCount = 0
     Erase m_tWnd
End Sub ' removeWindowHandles

Private Sub ResetMenuForRecalc(ByVal lIndex As Long)
     Dim tMI As MENUITEMINFO
     Dim hMenu As Long
     Dim hSubMenu As Long
     Dim lFlags As Long
     Dim lR As Long
     Dim lPosition As Long
     Dim iCount As Long
     Dim i As Long
     ' Modify the menu item:
     hMenu = m_tMI(lIndex).hMenu
     tMI.cbSize = Len(tMI)
     tMI.fMask = MIIM_SUBMENU
     GetMenuItemInfo hMenu, m_tMI(lIndex).lActualID, 0, tMI
     hSubMenu = tMI.hSubMenu
     iCount = GetMenuItemCount(hMenu)
     tMI.fMask = MIIM_ID
     For i = 0 To iCount - 1
          GetMenuItemInfo hMenu, i, 1, tMI
          If (tMI.wID <> m_tMI(lIndex).lActualID) Then
               lPosition = lPosition + 1
          Else
               Exit For
          End If
     Next
     ' remove it from the menu:
     RemoveMenu hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND
     ' Insert it back again at the corect position with the same ID etc:
     lFlags = plMenuFlags(lIndex)
     lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
     lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, _
          m_tMI(lIndex).lItemData)
     If (hSubMenu <> 0) Then
          ' If we had a submenu then put that back again:
          lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
          lFlags = lFlags Or MF_POPUP
          lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, hSubMenu, _
               m_tMI(lIndex).lItemData)
     End If
End Sub ' ResetMenuForRecalc

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
     ' Convert Automation color to Windows color
     If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function ' TranslateColor

'**************************************************************************************************
' cPopupMenu Implemented Properties/Methods
'**************************************************************************************************
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Dim lMenuId As Long
     Dim hMenu As Long
     Dim lItem As Long
     Dim lMenuCount As Long
     Dim lHiWord As Long
     Dim bEnabled As Boolean
     Dim bSeparator As Boolean
     Dim bFound As Boolean
     Dim iChar As Integer
     Dim lFlag As Long
     Dim bHandle As Boolean
     Dim lR As Long
     Select Case uMsg
          ' Handle Menu Select events:
          Case WM_MENUSELECT
               ' Extract the menu id and flags for the selected menu item:
               lHiWord = wParam \ &H10000
               lMenuId = wParam And &HFFFF&
               ' MenuId 0 corresponds to a separator on the system menu:
               If (lMenuId <> 0) Then
                    ' Extract separator & enabled/disabled from the flags
                    ' stored in the High Word of wParam:
                    bSeparator = ((lHiWord And MF_SEPARATOR) = MF_SEPARATOR)
                    bEnabled = ((lHiWord And MF_DISABLED) = MF_DISABLED) Or _
                         ((lHiWord And MF_GRAYED) = MF_GRAYED)
                    ' Menu handle is passed in as lParam:
                    hMenu = lParam
                    ' Now check if the message is a menu item higlight,
                    ' or whether it is indicating exit from the menu:
                    lMenuCount = GetMenuItemCount(hMenu)
                    For lItem = 0 To lMenuCount - 1
                         If (lMenuId = GetMenuItemID(hMenu, lItem)) Then
                              bFound = True
                              Exit For
                         End If
                    Next
                    ' Raise a highlight or menu exit as required:
                    If (bFound) Then
                         RaiseHighlightEvent lMenuId
                    Else
                         RaiseMenuExitEvent
                    End If
               End If
          ' Handle menu click events:
     Case WM_COMMAND
          ' Commands from menus are identified by an lParam of 0
          ' (otherwise it is set the hWnd of the control):
          bHandle = False
          If (lParam = 0) Or (lParam = m_hWndAttached) Then
               ' Low order word of the wParam item is the menu item id:
               lMenuId = (wParam And &HFFFF&)
               If (RaiseClickEvent(lMenuId)) Then
                    ' Don't send on the WM_COMMAND if the item
                    ' wasn't a VB menu, it might interfere
                    ' with some other control items!
               Else
                    'Handled
                    bHandled = True
               End If
          End If
     ' Draw Menu items:
     Case WM_DRAWITEM
          DrawItem wParam, lParam
          bHandled = True
     ' Measure Menu items prior to drawing them:
     Case WM_MEASUREITEM
          'iSubclass_WindowProc =
          Call MeasureItem(wParam, lParam)
          bHandled = True
     ' Handle accelerator (&key) messages in the menu:
'     Case WM_MENUCHAR
'          ' Check that this is my menu:
'          lFlag = wParam \ &H10000
'          If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
'               hMenu = lParam
'               iChar = (wParam And &HFFFF&)
'               ' See if this corresponds to an accelerator on the menu:
'               lR = plParseMenuChar(hMenu, iChar)
'          End If
     Case WM_INITMENUPOPUP, WM_UNINITMENUPOPUP
          ' Check the sys menu flag:
          If (lParam \ &H10000) > 0 Then
               ' System menu.
          Else
               hMenu = wParam
               ' Find the item which is the parent
               ' of this popup menu:
               RaiseInitMenuEvent hMenu, (uMsg = WM_INITMENUPOPUP)
               ' Win98+/2000 only
               If uMsg = WM_UNINITMENUPOPUP Then
                    removeWindowHandle hMenu
               End If
          End If
     Case WM_MENURBUTTONUP
          ' No action
     Case WM_WININICHANGE
          m_cNCM.GetMetrics
     Case WM_DESTROY
          Class_Terminate
     Case Else
          ' No action
   End Select
   DoEvents
End Sub

'**************************************************************************************************
' cPopupMenu Intrinsic Subs
'**************************************************************************************************
Private Sub Class_Initialize()
     Set m_sc = New cSubclass
     ' Control:
     m_lLastMaxId = &H800
     ' Stuff for drawing:
     Set m_cMemDC = New pcMemDC
     m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
     m_cMemDC.Height = 24
     m_ActiveMenuColor = CLR_INVALID
     m_InactiveMenuColor = CLR_INVALID
     m_MenuBackgroundColor = CLR_INVALID
     m_ActiveMenuBackColor = CLR_INVALID
     Set m_cNCM = New pcNCMetrics
     m_cNCM.GetMetrics
     Set m_fntSymbol = New StdFont
     m_fntSymbol.Name = "Marlett"
     m_fntSymbol.Size = Font.Size * 1.2
     pSetFont Nothing
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     Dim i As Long
     Set m_cMemDC = Nothing
     Clear
     Set m_cNCM = Nothing
     ' Clear up any stored menus:
     For i = 1 To m_iStoreCount
          Set m_cStoredMenu(i) = Nothing
     Next i
     Erase m_cStoredMenu
     m_iStoreCount = 0
     DestroySubClass
     Set m_sc = Nothing
End Sub ' Class_Terminate
