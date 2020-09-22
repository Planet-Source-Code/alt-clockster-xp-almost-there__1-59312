VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00AEB388&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clockster Settings"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00E4E5DF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4215
      Index           =   0
      Left            =   315
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   1
      Top             =   15
      Width           =   2220
      Begin VB.PictureBox picSettings 
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   4
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2115
         Width           =   2220
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "     Visual Display Settings"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image imgArrow 
            Height          =   105
            Index           =   4
            Left            =   1995
            Picture         =   "frmOptions.frx":030A
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.PictureBox picSettings 
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1485
         Width           =   2220
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "     Sound Settings"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   76
            Top             =   0
            Width           =   1305
         End
         Begin VB.Image imgArrow 
            Height          =   105
            Index           =   3
            Left            =   1995
            Picture         =   "frmOptions.frx":064C
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.PictureBox picSettings 
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1230
         Width           =   2220
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "     Visual Display Settings"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image imgArrow 
            Height          =   105
            Index           =   2
            Left            =   1995
            Picture         =   "frmOptions.frx":098E
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.PictureBox picSettings 
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
         Width           =   2220
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "     Chime Settings"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1290
         End
         Begin VB.Image imgArrow 
            Appearance      =   0  'Flat
            Height          =   105
            Index           =   1
            Left            =   1995
            Picture         =   "frmOptions.frx":0CD0
            Top             =   60
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.PictureBox picSettings 
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   148
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   345
         Width           =   2220
         Begin VB.Label lblSettings 
            AutoSize        =   -1  'True
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "     Visual Display Settings"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   70
            Top             =   -15
            Width           =   1815
         End
         Begin VB.Image imgArrow 
            Height          =   105
            Index           =   0
            Left            =   1995
            Picture         =   "frmOptions.frx":1012
            Top             =   60
            Width           =   60
         End
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   75
         ScaleHeight     =   94
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2685
         Width           =   2055
      End
      Begin VB.Label lblProps 
         BackColor       =   &H00AEB388&
         Caption         =   "ClocksterXP Zoomed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   5
         Top             =   2400
         Width           =   2220
      End
      Begin VB.Label lblProps 
         BackColor       =   &H00AEB388&
         Caption         =   " Tooltip Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   2220
      End
      Begin VB.Label lblProps 
         BackColor       =   &H00AEB388&
         Caption         =   " TouchSlider Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   885
         Width           =   2220
      End
      Begin VB.Label lblProps 
         BackColor       =   &H00AEB388&
         Caption         =   " Clock Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2220
      End
   End
   Begin VB.PictureBox picSideBar 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000005&
      Height          =   4215
      Index           =   0
      Left            =   15
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   210
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00AEB388&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4320
      Index           =   1
      Left            =   2625
      ScaleHeight     =   288
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   330
      TabIndex        =   68
      Top             =   15
      Width           =   4950
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   2
         Left            =   0
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
         Begin VB.TextBox txtTSSegSpace 
            BackColor       =   &H00E4E5DF&
            Enabled         =   0   'False
            Height          =   255
            Left            =   2700
            TabIndex        =   39
            Top             =   1410
            Width           =   360
         End
         Begin VB.PictureBox picTSForeColor 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2700
            ScaleHeight     =   195
            ScaleWidth      =   300
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
            Top             =   810
            Width           =   360
         End
         Begin VB.OptionButton optTSSegments 
            BackColor       =   &H00E4E5DF&
            Caption         =   "VolumeBar Is Segmented"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   33
            Top             =   420
            Width           =   2145
         End
         Begin VB.OptionButton optTSSegments 
            BackColor       =   &H00E4E5DF&
            Caption         =   "VolumeBar Is Solid Color"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   32
            Top             =   150
            Width           =   2055
         End
         Begin VB.TextBox txtTSSegSize 
            BackColor       =   &H00E4E5DF&
            Enabled         =   0   'False
            Height          =   255
            Left            =   2700
            TabIndex        =   37
            Top             =   1110
            Width           =   360
         End
         Begin ClocksterXP.UpDown UpDown2 
            Height          =   315
            Left            =   2700
            TabIndex        =   43
            Top             =   2295
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   15001055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   -50
            Max             =   50
         End
         Begin ClocksterXP.UpDown udAdjSliderHeight 
            Height          =   315
            Left            =   2700
            TabIndex        =   41
            Top             =   1905
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   15001055
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   -50
            Max             =   50
         End
         Begin VB.Label lblAdjSliderHeight 
            BackStyle       =   0  'Transparent
            Caption         =   "Increment/Decrement TouchSlider Height:"
            Height          =   390
            Left            =   150
            TabIndex        =   40
            Top             =   1860
            Width           =   2370
         End
         Begin VB.Label lblTSOffsetY 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset TouchSlider On Y-Axis: *"
            Height          =   195
            Left            =   150
            TabIndex        =   42
            Top             =   2340
            Width           =   2400
         End
         Begin VB.Label lblInfo2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":1354
            Height          =   630
            Left            =   150
            TabIndex        =   44
            Top             =   2835
            Width           =   4245
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   7
            X1              =   7
            X2              =   289
            Y1              =   181
            Y2              =   181
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   6
            X1              =   10
            X2              =   291
            Y1              =   115
            Y2              =   115
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   5
            X1              =   10
            X2              =   291
            Y1              =   47
            Y2              =   47
         End
         Begin VB.Label lblTSSegSpace 
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "Space Between Segments:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   38
            Top             =   1455
            Width           =   2100
         End
         Begin VB.Label lblTSForeColor 
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "Volume TouchSlider Bar Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   855
            Width           =   2190
         End
         Begin VB.Label lblTSSegSize 
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "TouchSlider Segment Size (2 to 5):"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   36
            Top             =   1155
            Width           =   2520
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   0
         Left            =   0
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   225
         Width           =   4500
         Begin ClocksterXP.UpDown UpDown1 
            Height          =   315
            Left            =   2565
            TabIndex        =   79
            Top             =   2100
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   556
            BackColor       =   15001055
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   -50
            Max             =   50
         End
         Begin VB.TextBox txtClkFontSize 
            BackColor       =   &H00E4E5DF&
            Height          =   255
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   600
            Width           =   645
         End
         Begin VB.TextBox txtClkFontStyle 
            BackColor       =   &H00E4E5DF&
            Height          =   255
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   600
            Width           =   1800
         End
         Begin VB.TextBox txtClkFontName 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   150
            Width           =   3315
         End
         Begin VB.CommandButton cmdClkFont 
            Caption         =   "Change &Font"
            Height          =   315
            Left            =   3210
            TabIndex        =   17
            Top             =   990
            Width           =   1185
         End
         Begin VB.PictureBox picClkTextColor 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1050
            ScaleHeight     =   255
            ScaleWidth      =   300
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
            Top             =   990
            Width           =   360
         End
         Begin VB.ComboBox cboClkDisplayFormat 
            BackColor       =   &H00E4E5DF&
            Height          =   315
            ItemData        =   "frmOptions.frx":13FC
            Left            =   2550
            List            =   "frmOptions.frx":1412
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1650
            Width           =   1875
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   1
            X1              =   10
            X2              =   292
            Y1              =   174
            Y2              =   174
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmOptions.frx":145C
            Height          =   630
            Left            =   150
            TabIndex        =   21
            Top             =   2760
            Width           =   4245
         End
         Begin VB.Label lblClkOffsetY 
            BackStyle       =   0  'Transparent
            Caption         =   "Offset Clock On Y-Axis *:"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   2175
            Width           =   1860
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   0
            X1              =   10
            X2              =   292
            Y1              =   99
            Y2              =   99
         End
         Begin VB.Label lblClkFontSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Size:"
            Height          =   195
            Left            =   2970
            TabIndex        =   13
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblClkFontStyle 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Style:"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   600
            Width           =   795
         End
         Begin VB.Label lblClkFontName 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Name:"
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   150
            Width           =   840
         End
         Begin VB.Label lblClkTextColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   1050
            Width           =   900
         End
         Begin VB.Label lblClkDisplayFormat 
            BackStyle       =   0  'Transparent
            Caption         =   "Choose Clock Display Format:"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   1740
            Width           =   2190
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   4
         Left            =   0
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
         Begin VB.CheckBox chkEnableTip 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Enable Volume Level Tooltip On Volume Change"
            Height          =   225
            Left            =   150
            TabIndex        =   54
            Top             =   150
            Value           =   1  'Checked
            Width           =   3750
         End
         Begin VB.PictureBox picTipBackColor 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1050
            ScaleHeight     =   195
            ScaleWidth      =   300
            TabIndex        =   65
            ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
            Top             =   2085
            Width           =   360
         End
         Begin VB.TextBox txtTipFontStyle 
            BackColor       =   &H00E4E5DF&
            Height          =   255
            Left            =   1050
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1770
         End
         Begin VB.TextBox txtTipFontSize 
            BackColor       =   &H00E4E5DF&
            Height          =   255
            Left            =   3720
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   1200
            Width           =   645
         End
         Begin VB.PictureBox picTipTextColor 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1050
            ScaleHeight     =   195
            ScaleWidth      =   300
            TabIndex        =   62
            ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
            Top             =   1635
            Width           =   360
         End
         Begin VB.CommandButton cmdTipFont 
            Caption         =   "Change &Font"
            Height          =   315
            Left            =   3210
            TabIndex        =   63
            Top             =   1605
            Width           =   1185
         End
         Begin VB.TextBox txtTipFontName 
            BackColor       =   &H00E4E5DF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   750
            Width           =   3315
         End
         Begin VB.Label lblTipBackColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   64
            Top             =   2100
            Width           =   870
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   2
            X1              =   10
            X2              =   292
            Y1              =   35
            Y2              =   35
         End
         Begin VB.Label lblFontColor 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Color:"
            Height          =   195
            Left            =   150
            TabIndex        =   61
            Top             =   1650
            Width           =   900
         End
         Begin VB.Label lblTipFontName 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Name:"
            Height          =   195
            Left            =   150
            TabIndex        =   55
            Top             =   750
            Width           =   840
         End
         Begin VB.Label lblTipFontStyle 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Style:"
            Height          =   195
            Left            =   150
            TabIndex        =   57
            Top             =   1215
            Width           =   795
         End
         Begin VB.Label lblTipFontSize 
            BackStyle       =   0  'Transparent
            Caption         =   "Font Size:"
            Height          =   195
            Left            =   2970
            TabIndex        =   59
            Top             =   1230
            Width           =   720
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   1
         Left            =   0
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
         Begin VB.CommandButton cmdSoundPathChime 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3765
            Picture         =   "frmOptions.frx":14F6
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2220
            Width           =   270
         End
         Begin VB.ComboBox cboClkChimeInterval 
            Appearance      =   0  'Flat
            BackColor       =   &H00E4E5DF&
            Height          =   315
            ItemData        =   "frmOptions.frx":18BA
            Left            =   150
            List            =   "frmOptions.frx":18CA
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtSoundPathChime 
            BackColor       =   &H00E4E5DF&
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2190
            Width           =   3915
         End
         Begin VB.CommandButton cmdTestSound 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   4110
            TabIndex        =   30
            ToolTipText     =   "Test Sound"
            Top             =   2190
            Width           =   270
         End
         Begin VB.OptionButton optChime 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Choose Custom Chime"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   26
            Top             =   1545
            Width           =   2160
         End
         Begin VB.OptionButton optChime 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Use Default App Chime"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   25
            Top             =   1245
            Width           =   2160
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   3
            X1              =   10
            X2              =   291
            Y1              =   70
            Y2              =   70
         End
         Begin VB.Label lblChimeInterval 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Chime Interval:"
            Height          =   195
            Left            =   150
            TabIndex        =   23
            Top             =   150
            Width           =   1845
         End
         Begin VB.Label lblSelect 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Select Sound:"
            Height          =   195
            Left            =   150
            TabIndex        =   27
            Top             =   1905
            Width           =   990
         End
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E4E5DF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3525
         Index           =   3
         Left            =   0
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   4500
         Begin VB.CheckBox chkTSEnableSound 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Enable Sound On Volume Change"
            Height          =   240
            Left            =   150
            TabIndex        =   46
            Top             =   150
            Width           =   2685
         End
         Begin VB.CommandButton cmdTSSoundPath 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   5.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3765
            Picture         =   "frmOptions.frx":1920
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   1545
            Width           =   270
         End
         Begin VB.CommandButton cmdTestSound 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   9.75
               Charset         =   2
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   4125
            TabIndex        =   52
            ToolTipText     =   "Test Sound"
            Top             =   1515
            Width           =   270
         End
         Begin VB.TextBox txtTSSoundPath 
            BackColor       =   &H00E4E5DF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   1515
            Width           =   3915
         End
         Begin VB.OptionButton optTSSound 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Choose Custom Sound"
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   48
            Top             =   900
            Width           =   2010
         End
         Begin VB.OptionButton optTSSound 
            BackColor       =   &H00E4E5DF&
            Caption         =   "Use App Default"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   47
            Top             =   615
            Value           =   -1  'True
            Width           =   2010
         End
         Begin VB.Line lnSeperator 
            BorderColor     =   &H00AEB388&
            BorderWidth     =   3
            Index           =   8
            X1              =   10
            X2              =   292
            Y1              =   33
            Y2              =   33
         End
         Begin VB.Label lblTSSoundPath 
            BackColor       =   &H00E4E5DF&
            BackStyle       =   0  'Transparent
            Caption         =   "Select Sound:"
            Enabled         =   0   'False
            Height          =   195
            Left            =   150
            TabIndex        =   49
            Top             =   1260
            Width           =   990
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   360
         Left            =   3420
         TabIndex        =   67
         Top             =   3885
         Width           =   1110
      End
      Begin VB.CheckBox chkAutoStart 
         BackColor       =   &H00AEB388&
         Caption         =   "Add registry entry to run ClocksterXP automatically when the system starts"
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   0
         TabIndex        =   66
         Top             =   3825
         Width           =   3300
      End
      Begin VB.Label lblOptons 
         BackColor       =   &H00AEB388&
         BackStyle       =   0  'Transparent
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   225
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  frmOptions.frm - Base Clockster Interface
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
'  frmOptions Constant Declares
'**************************************************************************************************
Private Const BOLD_FONTTYPE = &H100
Private Const CC_RGBINIT As Long = &H1
Private Const CC_ANYCOLOR As Long = &H100
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_LIMITSIZE = &H2000&

'**************************************************************************************************
'  frmOptions Struct Declares
'**************************************************************************************************
Private Type CHOOSECOLORSTRUCT
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As Long
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type ' CHOOSECOLORSTRUCT

Private Type LOGFONT
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

Private Type OPENFILENAME
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type ' OPENFILENAME

Private Type TCHOOSEFONT
     lStructSize As Long
     hWndOwner As Long
     hDC As Long
     lpLogFont As Long
     iPointSize As Long
     flags As Long
     rgbColors As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As Long
     hInstance As Long
     lpszStyle As String
     nFontType As Integer
     iAlign As Integer
     nSizeMin As Long
     nSizeMax As Long
End Type ' TCHOOSEFONT

'**************************************************************************************************
'  frmOptions Win32 API Declares
'**************************************************************************************************
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" ( _
     lpcc As CHOOSECOLORSTRUCT) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" ( _
     pChoosefont As TCHOOSEFONT) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
     "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias _
     "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, _
     ByVal dx As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
     ByVal dwRop As Long) As Long
     
'**************************************************************************************************
'  frmOptions Module-Scoped Variables
'**************************************************************************************************
Dim m_clkhWnd As Long
Dim dwCustClrs(0 To 15) As Long
Dim m_LastChimePath As String
Dim m_LastSoundPath As String
Dim m_menuidx As Integer
Dim m_tmr As cTimer
Implements WinSubHook2.iTimer

'**************************************************************************************************
'  frmOptions Constituent Control Events
'**************************************************************************************************
Private Sub cboClkChimeInterval_Click()
     frmClockFace.CustomClock1.ChimeInterval = cboClkChimeInterval.ListIndex
End Sub ' cboClkChimeInterval_Click

Private Sub cboClkChimeInterval_GotFocus()
     SendMessage cboClkChimeInterval.hwnd, &H14F, True, 0&
End Sub ' cboClkChimeInterval_GotFocus

Private Sub cboClkDisplayFormat_Click()
     frmClockFace.CustomClock1.TimeDisplayFormat = cboClkDisplayFormat.ListIndex
End Sub ' cboClkDisplayFormat_Click

Private Sub cboClkDisplayFormat_GotFocus()
     SendMessage cboClkDisplayFormat.hwnd, &H14F, True, 0&
End Sub ' cboClkDisplayFormat_GotFocus

Private Sub chkAutoStart_Click()
     If chkAutoStart Then
          AutoStartAdd
          chkAutoStart.Caption = "Uncheck to disallow ClocksterXP to run  automatically " + _
               "when the system starts"
     Else
          AutoStartDelete
          chkAutoStart.Caption = "Add registry entry to run ClocksterXP automatically " + _
               "when the system starts"
     End If
End Sub ' chkAutoStart_Click

Private Sub chkEnableTip_Click()
     frmTip.TipEnabled = CBool(chkEnableTip)
End Sub ' chkEnableTip_Click

Private Sub chkTSEnableSound_Click()
     frmClockFace.TouchSlider1.EnableSound = CBool(Abs(chkTSEnableSound.Value))
End Sub ' chkTSEnableSound_Click

Private Sub cmdClkFont_Click()
     Dim sFont As StdFont
     ' Set local font = to clock font
     Set sFont = frmClockFace.CustomClock1.Font
     ' Show font selection dialog
     ShowFont sFont, hwnd
     With sFont
          ' Get selections...font name
          txtClkFontName = .Name
          ' Get font size
          txtClkFontSize = CStr(.Size)
          ' Process other effects
          If .Bold And .Italic Then
               ' if bold and italic
               txtClkFontStyle = "Bold Italic"
          ElseIf .Bold And Not (.Italic) Then
               ' if only bold
               txtClkFontStyle = "Bold"
          ElseIf Not (.Bold) And .Italic Then
               ' if only italic
               txtClkFontStyle = "Italic"
          Else ' If no effects
               txtClkFontStyle = "Regular"
          End If
     End With
     ' set clock font = to local font
     frmClockFace.CustomClock1.Font = sFont
     ' Destroy stdfont object
     Set sFont = Nothing
End Sub

Private Sub cmdClose_Click()
     Unload Me
End Sub ' cmdClose_Click

Private Sub cmdSoundPathChime_Click()
     Dim lRtn As Long
     Dim sPathFile As String
     Dim lWidth As Long
     Dim lPos As Long
     Dim OFName As OPENFILENAME
     ' First, get the rectangle for the text box
     lWidth = (txtSoundPathChime.Width \ Screen.TwipsPerPixelX) * 0.9
     ' Intialize openfilename struct
     With OFName
          .lStructSize = Len(OFName)
          ' Set the parent window
          .hWndOwner = Me.hwnd
          ' Set the application's instance
          .hInstance = App.hInstance
          ' Select a filter
          .lpstrFilter = "Wav Files (*.wav)" + Chr$(0) + "*.wav" + Chr$(0) + _
               "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
          ' create a buffer for the file
          .lpstrFile = Space$(254)
          ' set the maximum length of a returned file
          .nMaxFile = 255
          ' Create a buffer for the file title
          .lpstrFileTitle = Space$(254)
          ' Set the maximum length of a returned file title
          .nMaxFileTitle = 255
          ' Set the initial directory
          If Len(m_LastChimePath) Then
               .lpstrInitialDir = m_LastChimePath
          Else
               .lpstrInitialDir = "C:\"
          End If
          ' Set the title
          .lpstrTitle = "Select Chime Sound"
          ' No flags
          .flags = 0
          ' Show the 'Open File'-dialog
          If GetOpenFileName(OFName) Then
               ' Trim the string
               sPathFile = Trim$(.lpstrFile)
               ' save the path
               m_LastChimePath = sPathFile
               ' Store long path in tooltip
               txtSoundPathChime.ToolTipText = sPathFile
               txtSoundPathChime.Tag = sPathFile
               ' save sound path
               frmClockFace.CustomClock1.ChimePath = txtSoundPathChime.Tag
               ' compact to fit in textbox
               Call CompactPath(hDC, sPathFile, lWidth)
               ' output to textbox
               txtSoundPathChime = Trim$(sPathFile)
          End If
     End With
     txtSoundPathChime.SetFocus
End Sub ' cmdSoundPathChime_Click

Private Sub cmdTestSound_Click(Index As Integer)
     Dim sFile As String
     Select Case Index
          Case 0
               ' get file name from usercontrol
               sFile = frmClockFace.TouchSlider1.SoundPath
               If optTSSound(0) Then
                   PlayResSound 101, "DEFAULT"
               Else
                    If Len(sFile) Then PlayWaveFile sFile
               End If
          Case 1
               ' get file name from usercontrol
               sFile = frmClockFace.CustomClock1.ChimePath
               If optChime(0) Then
                    PlayResSound 101, "CHIME"
               Else
                    If Len(sFile) Then PlayWaveFile sFile
               End If
     End Select
End Sub ' cmdTestSound_Click

Private Sub cmdTipFont_Click()
     Dim sFont As StdFont
     ' Set local font = to clock font
     Set sFont = frmTip.lblTip.Font
     ' Show font selection dialog
     ShowFont sFont, hwnd
     With sFont
          ' Get selections...font name
          txtTipFontName = .Name
          ' Get font size
          txtTipFontSize = CStr(.Size)
          ' Process other effects
          If .Bold And .Italic Then
               ' if bold and italic
               txtTipFontStyle = "Bold Italic"
          ElseIf .Bold And Not (.Italic) Then
               ' if only bold
               txtTipFontStyle = "Bold"
          ElseIf Not (.Bold) And .Italic Then
               ' if only italic
               txtTipFontStyle = "Italic"
          Else ' If no effects
               txtTipFontStyle = "Regular"
          End If
     End With
     ' set clock font = to local font
     frmTip.lblTip.Font = sFont
     ' Destroy stdfont object
     Set sFont = Nothing
End Sub ' cmdTipFont_Click

Private Sub cmdTSSoundPath_Click()
     Dim lRtn As Long
     Dim sPathFile As String
     Dim lWidth As Long
     Dim lPos As Long
     Dim OFName As OPENFILENAME
     ' First, get the rectangle for the text box
     lWidth = txtTSSoundPath.Width * 0.9
     ' Intialize openfilename struct
     With OFName
          .lStructSize = Len(OFName)
          ' Set the parent window
          .hWndOwner = Me.hwnd
          ' Set the application's instance
          .hInstance = App.hInstance
          ' Select a filter
          .lpstrFilter = "Wav Files (*.wav)" + Chr$(0) + "*.wav" + Chr$(0) + _
               "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
          ' create a buffer for the file
          .lpstrFile = Space$(254)
          ' set the maximum length of a returned file
          .nMaxFile = 255
          ' Create a buffer for the file title
          .lpstrFileTitle = Space$(254)
          ' Set the maximum length of a returned file title
          .nMaxFileTitle = 255
          ' Set the initial directory
          If Len(m_LastChimePath) Then
               .lpstrInitialDir = m_LastSoundPath
          Else
               .lpstrInitialDir = "C:\"
          End If
          ' Set the title
          .lpstrTitle = "Select Volume Change Sound"
          ' No flags
          .flags = 0
          ' Show the 'Open File'-dialog
          If GetOpenFileName(OFName) Then
               ' Trim the string
               sPathFile = Trim$(.lpstrFile)
               ' save the path
               m_LastChimePath = sPathFile
               ' Store long path in tooltip
               txtTSSoundPath.ToolTipText = sPathFile
               txtTSSoundPath.Tag = sPathFile
               ' save sound path
               frmClockFace.TouchSlider1.SoundPath = txtTSSoundPath.Tag
               ' compact to fit in textbox
               Call CompactPath(hDC, sPathFile, lWidth)
               ' output to textbox
               txtTSSoundPath = Trim$(sPathFile)
          End If
     End With
     txtTSSoundPath.SetFocus
End Sub ' cmdTSSoundPath_Click

Private Sub lblSettings_Click(Index As Integer)
     Call picSettings_Click(Index)
End Sub ' lblSettings_Click

Private Sub optChime_Click(Index As Integer)
     If optChime(0) Then
          txtSoundPathChime.Enabled = False
          cmdSoundPathChime.Enabled = False
     Else
          txtSoundPathChime.Enabled = True
          cmdSoundPathChime.Enabled = True
     End If
     ' set the property
     frmClockFace.CustomClock1.Chime = Index
End Sub ' optChime_Click

Private Sub optTSSegments_Click(Index As Integer)
     frmClockFace.TouchSlider1.ProgressStyle = Index
     If Index = False Then
          txtTSSegSize.Enabled = True
          txtTSSegSpace.Enabled = True
          lblTSSegSize.Enabled = True
          lblTSSegSpace.Enabled = True
     Else
          txtTSSegSize.Enabled = False
          txtTSSegSpace.Enabled = False
          lblTSSegSize.Enabled = False
          lblTSSegSpace.Enabled = False
     End If
End Sub ' optTSSegments_Click

Private Sub optTSSound_Click(Index As Integer)
     frmClockFace.TouchSlider1.Sound = Index
     If Index = 0 Then
          txtTSSoundPath.Enabled = False
          cmdTSSoundPath.Enabled = False
          lblTSSoundPath.Enabled = False
     Else
          txtTSSoundPath.Enabled = True
          cmdTSSoundPath.Enabled = True
          lblTSSoundPath.Enabled = True
     End If
End Sub ' optTSSound_Click

Private Sub picClkTextColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picClkTextColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmClockFace.CustomClock1.ForeColor = lColor
     End If
End Sub ' picClkTextColor_Click

Private Sub picSettings_Click(Index As Integer)
     Dim lLoop As Long
     For lLoop = 0 To picSettings.UBound
          If Index = lLoop Then
               picSettings(lLoop).BackColor = 14080451
               picSettings(lLoop).SetFocus
               imgArrow(lLoop).Visible = True
               picOptions(lLoop).Visible = True
               m_menuidx = Index
          Else
               picSettings(lLoop).BackColor = 15001055
               imgArrow(lLoop).Visible = False
               picOptions(lLoop).Visible = False
          End If
     Next
End Sub ' picSettings_Click

Private Sub picSettings_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyUp Then
          If m_menuidx = 0 Then
               m_menuidx = 4
               Call picSettings_Click(m_menuidx)
          Else
               m_menuidx = m_menuidx - 1
               Call picSettings_Click(m_menuidx)
          End If
     ElseIf KeyCode = vbKeyDown Then
          If m_menuidx = 4 Then
               m_menuidx = 0
               Call picSettings_Click(m_menuidx)
          Else
               m_menuidx = m_menuidx + 1
               Call picSettings_Click(m_menuidx)
          End If
     ElseIf KeyCode = vbKeyReturn Then
          m_menuidx = m_menuidx + 1
          If m_menuidx > 4 Then m_menuidx = 0
          Call picSettings_Click(m_menuidx)
     End If
End Sub ' picSettings_KeyUp

Private Sub picTipBackColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picTipBackColor.BackColor = lColor
          ' set tip form's backcolor
          frmTip.TipBackColor = lColor
     End If
End Sub ' picTipBackColor_Click

Private Sub picTipTextColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picTipTextColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmTip.TipForeColor = lColor
     End If
End Sub ' picTipTextColor_Click

Private Sub picTSForeColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picTSForeColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmClockFace.TouchSlider1.Color = lColor
     End If
End Sub ' picTSForeColor_Click

Private Sub txtTSSegSize_Change()
     If IsNumeric(txtTSSegSize) Then
          If Val(txtTSSegSize) > 0 And Val(txtTSSegSize) < 5 Then
               frmClockFace.TouchSlider1.SegmentSize = CLng(txtTSSegSize)
          Else
               txtTSSegSize = ""
          End If
     Else
          ' restore
          txtTSSegSize = CStr(frmClockFace.TouchSlider1.SegmentSize)
     End If
End Sub ' txtTSSegSize_Change

Private Sub txtTSSegSpace_Change()
     If IsNumeric(txtTSSegSpace) Then
          If Val(txtTSSegSpace) > 0 And Val(txtTSSegSize) < 5 Then
               frmClockFace.TouchSlider1.SegmentSpace = CLng(txtTSSegSpace)
          Else
               txtTSSegSpace = ""
          End If
     Else
          ' restore
          txtTSSegSpace = CStr(frmClockFace.TouchSlider1.SegmentSpace)
     End If
End Sub ' txtTSSegSpace_Change

Private Sub udAdjSliderHeight_Change()
     frmClockFace.TouchSlider1.HeightAdjust = udAdjSliderHeight.Value
End Sub ' udAdjSliderHeight_Change

Private Sub UpDown1_Change()
     frmClockFace.CustomClock1.ClockOffsetY = UpDown1.Value
End Sub ' UpDown1_Change

Private Sub UpDown2_Change()
     frmClockFace.TouchSlider1.SliderOffsetY = UpDown2.Value
End Sub ' UpDown2_Change

'**************************************************************************************************
'  frmOptions Intrinsic Events
'**************************************************************************************************
Private Sub Form_Load()
     Dim lHDC As Long
     Dim lHt As Long
     Dim lWt As Long
     Dim rc As tRECT
     Dim mLogo As cLogo
     Dim sPath As String
     Dim lLoop As Long
     ' Set form's position
     SetPosition
     ' create the logo object and call the drawing methods
     Set mLogo = New cLogo
     ' set logo properties
     With mLogo
          .DrawingObject = picSideBar(0)
          .StartColor = 11449224        ' the start gradient color
          .EndColor = 11449224          ' the end gradient color
          .Caption = "ClocksterXP"      ' logo text
          .Draw                         ' call the class draw method
     End With
     ' Get clock properties
     With frmClockFace.CustomClock1
          picClkTextColor.BackColor = .ForeColor
          txtClkFontName = .Font.Name
          txtClkFontSize = .Font.Size
          ' Set up clock font effect options
          If .Font.Bold And .Font.Italic Then
               ' text is bold and italicized
               txtClkFontStyle = "Bold Italic"
          ElseIf .Font.Bold And Not (.Font.Italic) Then
               ' text is only bold
               txtClkFontStyle = "Bold"
          ElseIf Not (.Font.Bold) And .Font.Italic Then
               ' text is only italicized
               txtClkFontStyle = "Italic"
          Else
               ' text has no effects
               txtClkFontStyle = "Regular"
          End If
          ' set display format
          cboClkDisplayFormat.ListIndex = .TimeDisplayFormat
          ' chime settings
          cboClkChimeInterval.ListIndex = .ChimeInterval
          ' internal or external chime
          optChime(.Chime).Value = True
          ' get a chime path
          sPath = .ChimePath
          ' extract the path from the property
          sPath = ExtractFilePath(sPath)
          ' store it in last path variable
          m_LastChimePath = sPath
          If Len(sPath) Then
               Call CompactPath(hDC, sPath, txtSoundPathChime.Width * 0.9)
               txtSoundPathChime = sPath
          Else
               m_LastChimePath = "C:\"
          End If
          ' get x and y offsets
          UpDown1.Value = .ClockOffsetY
     End With
     ' Get slider props
     With frmClockFace.TouchSlider1
          chkTSEnableSound = CInt(Abs(.EnableSound))
          optTSSegments(.ProgressStyle) = True
          picTSForeColor.BackColor = .Color
          txtTSSegSize = CStr(.SegmentSize)
          txtTSSegSpace = CStr(.SegmentSpace)
          optTSSound(.Sound) = True
          ' get a volume change sound path
          sPath = .SoundPath
          ' extract the path from the property
          sPath = ExtractFilePath(sPath)
          ' store it in last path variable
          m_LastSoundPath = sPath
          If Len(sPath) Then
               Call CompactPath(hDC, sPath, txtTSSoundPath.Width * 0.9)
               txtTSSoundPath = sPath
          Else
               m_LastChimePath = "C:\"
          End If
          ' get height adjustment
          udAdjSliderHeight.Value = .HeightAdjust
          ' get offset
          UpDown2.Value = .SliderOffsetY
     End With
     ' Get tip properties
     With frmTip
          picTipTextColor.BackColor = .lblTip.ForeColor
          picTipBackColor.BackColor = .BackColor
          chkEnableTip = CInt(Abs(.TipEnabled))
          txtTipFontName = .lblTip.Font.Name
          txtTipFontSize = .lblTip.Font.Size
          ' Set up clock font effect options
          If .lblTip.Font.Bold And .lblTip.Font.Italic Then
               ' text is bold and italicized
               txtTipFontStyle = "Bold Italic"
          ElseIf .lblTip.Font.Bold And Not (.lblTip.Font.Italic) Then
               ' text is only bold
               txtTipFontStyle = "Bold"
          ElseIf Not (.lblTip.Font.Bold) And .lblTip.Font.Italic Then
               ' text is only italicized
               txtTipFontStyle = "Italic"
          Else
               ' text has no effects
               txtTipFontStyle = "Regular"
          End If
     End With
     ' We've painted the logo to the picbox so destroy the logo object
     Set mLogo = Nothing
     ' get the clock's DC from the clockripper control
     m_clkhWnd = frmClockFace.ClockRip.ClockhWnd
     ' enable timer
     Set m_tmr = New cTimer
     ' start timer
     m_tmr.TmrStart Me, 30
     ' Is application autostarting?
     chkAutoStart = CInt(Abs(IsAutoStart))
     ' show form
     Me.Show
     ' set first option
     Call picSettings_Click(0)
End Sub ' Form_Load

Private Sub Form_Unload(Cancel As Integer)
     Dim lLoop As Long
     ' stop the timer
     m_tmr.TmrStop
     ' destroy timer object
     Set m_tmr = Nothing
End Sub ' Form_Unload

'**************************************************************************************************
'  frmOptions Private Methods
'**************************************************************************************************
Private Function AutoStartAdd() As Long
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

Private Function AutoStartDelete() As Long
     Dim hKey As Long
     Dim lRtn As Long
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     If lRtn = False Then AutoStartDelete = RegDeleteValue(hKey, App.EXEName)
End Function ' AutoStartDelete

Private Function BytesToStr(ab() As Byte) As String
     BytesToStr = StrConv(ab, vbUnicode)
End Function ' BytesToStr

Private Sub CompactPath(ByVal lHDC As Long, ByRef sPathFile As String, ByVal lWidth As Long)
     Call PathCompactPath(hDC, sPathFile, lWidth)
End Sub ' TruncatePath

Private Function ExtractFilePath(ByVal vStrFullPath As String) As String
     Dim iPos As Integer
     iPos = InStrRev(vStrFullPath, Chr(92))
     ExtractFilePath = Left$(vStrFullPath, iPos)
End Function ' ExtractFilePath

Private Function IsArrayEmpty(va As Variant) As Boolean
     Dim v As Variant
     On Error Resume Next
     v = va(LBound(va))
     IsArrayEmpty = (Err <> 0)
End Function ' IsArrayEmpty

Private Function IsAutoStart() As Boolean
     Dim hKey As Long
     Dim lType As Long
     Dim sValue As String
     ' set value
     sValue = App.EXEName
     ' If the key exists
     If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          0, KEY_READ, hKey) = False Then
          ' Look for the subkey named after the application
          If RegQueryValueEx(hKey, sValue, ByVal 0&, lType, ByVal 0&, _
               ByVal 0&) = False Then
               IsAutoStart = True
               ' Close the registry key handle.
               RegCloseKey hKey
          End If
    End If
End Function ' IsAutoStart

Private Sub SetPosition()
     Dim lHeight As Long
     Dim lLeft As Long
     Dim lTop As Long
     Dim lWidth As Long
     Dim rcWA As tRECT
     Dim rcWnd As tRECT
     Dim tbPos As APPBARDATA
     ' Get the screen dimensions in rcWA
     SystemParametersInfo SPI_GETWORKAREA, 0, rcWA, 0
     ' get taskbar position to determine where our form is located
     SHAppBarMessage ABM_GETTASKBARPOS, tbPos
     Select Case tbPos.uEdge
          Case ABE_LEFT
               lLeft = rcWA.Left * Screen.TwipsPerPixelX
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
          Case ABE_TOP
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = rcWA.Top * Screen.TwipsPerPixelY
               Move lLeft, lTop
          Case ABE_RIGHT
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
          Case ABE_BOTTOM
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
     End Select
     DoEvents
End Sub ' SetPosition

Private Function ShowColor() As Long
     Dim cc As CHOOSECOLORSTRUCT
     With cc
          'set the flags based on the check and option buttons
          .flags = CC_ANYCOLOR
          .flags = .flags Or CC_RGBINIT
          .rgbResult = frmClockFace.CustomClock1.ForeColor
          'size of structure
          .lStructSize = Len(cc)
          'owner of the dialog
          .hWndOwner = hwnd
          'assign the custom colour selections
          .lpCustColors = VarPtr(dwCustClrs(0))
     End With
     If ChooseColor(cc) = 1 Then
          ' Return function
          ShowColor = cc.rgbResult
     Else
          ShowColor = -1
     End If
End Function ' ShowColor

Private Function ShowFont(CurFont As Font, Optional Owner As Long = -1, _
     Optional Color As Long = vbBlack, Optional MinSize As Long = 0, _
     Optional MaxSize As Long = 0, Optional flags As Long = 0) As Boolean
     Const PointsPerTwip = 1440 / 72
     Dim cf As TCHOOSEFONT
     Dim m_lApiReturn As Long
     Dim m_lExtendedError As Long
     Dim fnt As LOGFONT
     m_lApiReturn = 0
     m_lExtendedError = 0
     ' Unwanted Flags bits
     Const CF_FontNotSupported = CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE
     ' Flags can get reference variable or constant with bit flags
     ' Must have some fonts
     If (flags And CF_PRINTERFONTS) = 0 Then flags = flags Or CF_SCREENFONTS
     ' Color can take initial color, receive chosen color
     If Color <> vbBlack Then flags = flags Or CF_EFFECTS
     ' MinSize can be minimum size accepted
     If MinSize Then flags = flags Or CF_LIMITSIZE
     ' MaxSize can be maximum size accepted
     If MaxSize Then flags = flags Or CF_LIMITSIZE
     ' Put in required internal flags and remove unsupported
     flags = (flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
     ' Initialize LOGFONT variable
     With fnt
          .lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
          .lfWeight = CurFont.Weight
          .lfItalic = CurFont.Italic
          .lfUnderline = CurFont.Underline
          .lfStrikeOut = CurFont.Strikethrough
     End With
     ' Other fields zero
     StrToBytes fnt.lfFaceName, CurFont.Name
     ' Initialize TCHOOSEFONT variable
     With cf
          .lStructSize = Len(cf)
           If Owner <> -1 Then .hWndOwner = Owner
          .lpLogFont = VarPtr(fnt)
          .iPointSize = CurFont.Size * 10
          .flags = flags
          .rgbColors = Color
          .nSizeMin = MinSize
          .nSizeMax = MaxSize
     End With
     ' All other fields zero
     m_lApiReturn = CHOOSEFONT(cf)
     Select Case m_lApiReturn
          Case 1
               ' Success
               ShowFont = True
               flags = cf.flags
               Color = cf.rgbColors
               CurFont.Bold = cf.nFontType And BOLD_FONTTYPE
               CurFont.Italic = fnt.lfItalic
               CurFont.Strikethrough = fnt.lfStrikeOut
               CurFont.Underline = fnt.lfUnderline
               CurFont.Weight = fnt.lfWeight
               CurFont.Size = cf.iPointSize / 10
               CurFont.Name = BytesToStr(fnt.lfFaceName)
          Case 0
               ' Cancelled
               ShowFont = False
          Case Else
               ShowFont = False
     End Select
End Function ' ShowFont

Private Sub StrToBytes(ab() As Byte, s As String)
     Dim cab As Long
     If IsArrayEmpty(ab) Then
          ' Assign to empty array
          ab = StrConv(s, vbFromUnicode)
     Else
          ' Copy to existing array, padding or truncating if necessary
          cab = UBound(ab) - LBound(ab) + 1
          If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
          CopyMemoryStr ab(LBound(ab)), s, cab
     End If
End Sub ' StrToBytes

'**************************************************************************************************
'  frmOptions Implemented iTimer Interface
'**************************************************************************************************
Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
     Dim rc As tRECT
     Dim lHDC As Long
     Cls
     ' Get the clock's rectangle
     GetClientRect m_clkhWnd, rc
     ' now get the clock's DC
     lHDC = GetWindowDC(m_clkhWnd)
     ' if we have a handle
     If lHDC Then
          ' blit the clock window to the picturebox
          StretchBlt picPreview.hDC, 0, 0, picPreview.ScaleWidth, picPreview.ScaleHeight, _
               lHDC, 0, 0, rc.Right, rc.Bottom, vbSrcCopy
          ' set the image
          picPreview.Picture = picPreview.Image
          ' release the dc
          ReleaseDC m_clkhWnd, lHDC
     End If
End Sub

