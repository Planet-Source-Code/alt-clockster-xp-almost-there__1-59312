VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Clockster"
   ClientHeight    =   3645
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6090
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4770
      TabIndex        =   0
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "More acknowledgements forthcoming..."
      Height          =   240
      Left            =   1005
      TabIndex        =   9
      Top             =   2715
      Width           =   4395
   End
   Begin VB.Label lblBirch 
      Caption         =   "Copyright © 1996-2005 VBnet, Randy Birch, All Rights Reserved"
      Height          =   240
      Left            =   1005
      TabIndex        =   8
      Top             =   2430
      Width           =   4845
   End
   Begin VB.Label lblMcMahon 
      Height          =   615
      Left            =   1005
      TabIndex        =   7
      Top             =   1695
      Width           =   4875
   End
   Begin VB.Label lblCode 
      Caption         =   "Source Code Contributers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1005
      TabIndex        =   6
      Top             =   1425
      Width           =   2835
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Copyright © 2005, Alan Tucker, All Rights Reserved."
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1005
      TabIndex        =   5
      Top             =   1125
      Width           =   4830
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -60
      X2              =   6135
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1005
      TabIndex        =   2
      Top             =   825
      Width           =   4860
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title:  "
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1005
      TabIndex        =   3
      Top             =   225
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -60
      X2              =   6135
      Y1              =   3135
      Y2              =   3120
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   195
      Left            =   1005
      TabIndex        =   4
      Top             =   525
      Width           =   4875
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = "Direct access to your master volume via the system clock interface."
    lblMcMahon = "Copryight © 2002-2005, VBAccelerator.com." + vbCrLf + _
    "This product includes software developed by vbAccelerator (http://vbaccelerator.com/)."
    lblBirch = "Copyright © 1996-2005 VBnet, Randy Birch, All Rights Reserved"
End Sub

