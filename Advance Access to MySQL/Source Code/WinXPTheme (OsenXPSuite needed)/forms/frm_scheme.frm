VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.1#0"; "osenxpsuite.ocx"
Begin VB.Form Frm_Scheme 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Change Colorscheme"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2070
   ScaleWidth      =   2535
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite.OsenXPButton OsenXPButton1 
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   1590
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frm_scheme.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   0   'False
      GradientColor1  =   14854529
      GradientColor2  =   16777215
   End
   Begin VB.ComboBox CboScheme 
      Height          =   315
      ItemData        =   "frm_scheme.frx":001C
      Left            =   390
      List            =   "frm_scheme.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1140
      Width           =   1755
   End
   Begin osenxpsuite.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Change Colorscheme"
      TitleTop        =   7
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
      EnableCloseButton=   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your favorite colorscheme:"
      Height          =   465
      Left            =   300
      TabIndex        =   2
      Top             =   600
      Width           =   1860
   End
End
Attribute VB_Name = "Frm_Scheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboScheme_Click()
    MyScheme = CboScheme.ListIndex
    Me.OsenXPForm1.ColorScheme = MyScheme
End Sub

Private Sub Form_Load()
    Me.OsenXPForm1.ColorScheme = MyScheme
    CboScheme.ListIndex = MyScheme
End Sub

Private Sub OsenXPButton1_Click()
    Unload Me
End Sub
