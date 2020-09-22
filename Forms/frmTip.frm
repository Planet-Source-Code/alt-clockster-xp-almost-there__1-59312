VERSION 5.00
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   154
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   45
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

Private Const m_def_backcolor = 14811135
Private Const m_def_enabled = True
Private Const m_def_fontbold = False
Private Const m_def_fontitalic = False
Private Const m_def_fontname = "Verdana"
Private Const m_def_fontsize = 8.25
Private Const m_def_forecolor = 0

Dim m_TipEnabled As Boolean

Public Property Get TipBackColor() As OLE_COLOR
     TipBackColor = frmTip.BackColor
End Property ' Get TipBackColor

Public Property Let TipBackColor(New_TipBackColor As OLE_COLOR)
     frmTip.BackColor = New_TipBackColor
     ' store in registry
     UpdateRegistry "TipBackColor", CStr(New_TipBackColor), CStr(New_TipBackColor)
     ' draw border
     DrawBorder
End Property ' Let TipBackColor

Public Property Get TipEnabled() As Boolean
     TipEnabled = m_TipEnabled
End Property ' Get TipEnabled

Public Property Let TipEnabled(New_TipEnabled As Boolean)
     m_TipEnabled = New_TipEnabled
     ' set registry
     UpdateRegistry "TipEnabled", CStr(New_TipEnabled), CStr(New_TipEnabled)
End Property ' Let TipEnabled

Public Property Get TipFontBold() As Boolean
     TipFontBold = lblTip.Font.Bold
End Property ' Get TipFontBold

Public Property Let TipFontBold(New_TipFontBold As Boolean)
     lblTip.Font.Bold = New_TipFontBold
     ' update registry
     UpdateRegistry "TipFontBold", CStr(New_TipFontBold), CStr(New_TipFontBold)
End Property ' Let TipFontBold

Public Property Get TipFontItalic() As Boolean
     TipFontItalic = lblTip.Font.Italic
End Property ' Get TipFontItalic

Public Property Let TipFontItalic(New_TipFontItalic As Boolean)
     lblTip.Font.Italic = New_TipFontItalic
     ' update registry
     UpdateRegistry "TipFontItalic", CStr(New_TipFontItalic), CStr(New_TipFontItalic)
End Property ' Let TipFontItalic

Public Property Get TipFontName() As String
     TipFontName = lblTip.Font.Name
End Property ' Get TipFontName

Public Property Let TipFontName(New_TipFontName As String)
     lblTip.Font.Name = New_TipFontName
     ' update registry
     UpdateRegistry "TipFontName", New_TipFontName, New_TipFontName
End Property ' Let TipFontName

Public Property Get TipFontSize() As Single
     TipFontSize = lblTip.Font.Size
End Property ' Get TipFontSize

Public Property Let TipFontSize(New_TipFontSize As Single)
     lblTip.Font.Size = New_TipFontSize
     ' update registry
     UpdateRegistry "TipFontSize", CStr(New_TipFontSize), CStr(New_TipFontSize)
End Property ' Let TipFontSize

Public Property Get TipForeColor() As OLE_COLOR
     TipForeColor = Me.ForeColor
End Property ' TipForeColor

Public Property Let TipForeColor(New_TipForeColor As OLE_COLOR)
     Me.ForeColor = New_TipForeColor
     lblTip.ForeColor = New_TipForeColor
     ' update registry
     UpdateRegistry "TipForeColor", CStr(New_TipForeColor), CStr(New_TipForeColor)
     ' draw the border
     DrawBorder
End Property ' Let TipForeColor

Private Sub Form_Load()
     Dim hKey As Long
     Dim lRtn As Long
     Dim sRtn As String
     ' retrieve settings from registry
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Clockster\Clockster\Settings", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_QUERY_VALUE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False And hKey Then
          ' get tooltip backcolor
          sRtn = GetRegValue(hKey, "TipBackColor")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               Me.BackColor = CLng(sRtn)
          Else ' set to default
               Me.BackColor = m_def_backcolor
               ' set registry
               UpdateRegistry "TipBackColor", CStr(m_def_backcolor), _
                    CStr(m_def_backcolor)
          End If
          ' is tip enabled?
          sRtn = GetRegValue(hKey, "TipEnabled")
          ' if we have a value
          If Len(sRtn) Then
               ' set variable to reg value
               TipEnabled = CBool(sRtn)
          Else ' set to default
               TipEnabled = m_def_enabled
               ' set registry
               UpdateRegistry "TipEnabled", CStr(m_def_enabled), _
                    CStr(m_def_enabled)
          End If
          ' get tooltip fontbold
          sRtn = GetRegValue(hKey, "TipFontBold")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               lblTip.Font.Bold = CBool(sRtn)
          Else ' set to default
               lblTip.Font.Bold = m_def_fontbold
               ' set registry
               UpdateRegistry "TipFontBold", CStr(m_def_fontbold), _
                    CStr(m_def_fontbold)
          End If
          ' Is the font italicized
          ' get tooltip fontbold
          sRtn = GetRegValue(hKey, "TipFontItalic")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               lblTip.Font.Italic = CBool(sRtn)
          Else ' set to default
               lblTip.Font.Italic = m_def_fontitalic
               ' set registry
               UpdateRegistry "TipFontItalic", CStr(m_def_fontitalic), _
                    CStr(m_def_fontitalic)
          End If
          ' what's the name of the font?
          sRtn = GetRegValue(hKey, "TipFontName")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               lblTip.Font.Name = sRtn
          Else ' set to default
               lblTip.Font.Name = m_def_fontname
               ' set registry
               UpdateRegistry "TipFontName", CStr(m_def_fontname), _
                    CStr(m_def_fontname)
          End If
          ' What's the size of the font?
          sRtn = GetRegValue(hKey, "TipFontSize")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               lblTip.Font.Size = CSng(sRtn)
          Else ' set to default
               lblTip.Font.Size = m_def_fontsize
               ' set registry
               UpdateRegistry "TipFontSize", CStr(m_def_fontsize), _
                    CStr(m_def_fontsize)
          End If
          ' get tooltip fontbold
          sRtn = GetRegValue(hKey, "TipForeColor")
          ' if we have a value
          If Len(sRtn) Then
               ' set property to reg value
               lblTip.ForeColor = CLng(sRtn)
               Me.ForeColor = CLng(sRtn)
          Else ' set to default
               lblTip.ForeColor = m_def_forecolor
               Me.ForeColor = m_def_forecolor
               ' set registry
               UpdateRegistry "TipForeColor", CStr(m_def_forecolor), _
                    CStr(m_def_forecolor)
          End If
     End If
     ' draw the border
     DrawBorder
End Sub ' Form_Load

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

Private Sub DrawBorder()
     ' Draw the border
     Me.Cls
     Me.Line (0, 0)-(ScaleWidth, 0)
     Me.Line (0, 0)-(0, ScaleHeight)
     Me.Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight)
     Me.Line (0, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1)
End Sub ' DrawBorder

