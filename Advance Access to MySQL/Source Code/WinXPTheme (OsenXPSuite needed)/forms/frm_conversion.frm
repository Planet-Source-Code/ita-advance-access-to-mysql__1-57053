VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.1#0"; "osenxpsuite.ocx"
Begin VB.Form Frm_Conversion 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Advance Access to MySQL"
   ClientHeight    =   5760
   ClientLeft      =   3930
   ClientTop       =   1725
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_conversion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   2  'CenterScreen
   Begin osenxpsuite.OsenXPFrame fra_progress 
      Height          =   3300
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   1410
      Visible         =   0   'False
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5821
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BorderColor     =   12570832
      Begin osenxpsuite.OsenXPCheckBox ChkDropTables 
         Height          =   255
         Left            =   2700
         TabIndex        =   44
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         Value           =   1
         Caption         =   "Drop Table(s)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Drop Table(s)"
         BackColor       =   14215660
      End
      Begin osenxpsuite.OsenXPButton cmdSelect 
         Height          =   315
         Left            =   5490
         TabIndex        =   42
         Top             =   2970
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Select &All"
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
         MICON           =   "frm_conversion.frx":058A
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
      Begin osenxpsuite.OsenXPCheckBox ChkNoData 
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "Copy Only Structure"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Copy Only Structure"
         BackColor       =   14215660
      End
      Begin osenxpsuite.OsenXPListBox LstTables 
         Height          =   2445
         Left            =   600
         TabIndex        =   37
         Top             =   480
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   4313
         Appearance      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSelected    =   16576
         BackSelected    =   7381139
         BackSelectedG1  =   16777215
         BackSelectedG2  =   8632490
         ItemHeight      =   20
         ItemHeightAuto  =   0   'False
         ItemOffset      =   2
         ItemTextLeft    =   17
         SelectModeStyle =   2
         Lstyle          =   1
         ColorScheme     =   1
         ShowHeader      =   -1  'True
         Columns         =   2
         ShowGridLines   =   -1  'True
         AlternateRowColors=   -1  'True
         CT1             =   "Table name"
         CA1             =   0
         CW1             =   300
         CT2             =   "Rows"
         CA2             =   2
         CW2             =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select MS Access table(s) to convert into MySQL format"
         Height          =   195
         Left            =   600
         TabIndex        =   39
         Top             =   150
         Width           =   4005
      End
   End
   Begin osenxpsuite.OsenXPButton CmdFinish 
      Height          =   345
      Left            =   3600
      TabIndex        =   20
      Top             =   5250
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "&Finish"
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
      MICON           =   "frm_conversion.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin VB.PictureBox PicWelcome 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      Picture         =   "frm_conversion.frx":05C2
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   1
      Top             =   450
      Width           =   7440
      Begin osenxpsuite.OsenXPOptionButton OptScenario 
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   3960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OptScenario"
         BackColor       =   16777215
         AutoChangeBackColor=   0   'False
      End
   End
   Begin osenxpsuite.OsenXPFrame fra_progress 
      Height          =   2865
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   83325
      Visible         =   0   'False
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   5054
      Caption         =   "Source Database (MS Access):"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   12570832
      Image           =   "frm_conversion.frx":254F4
      Icon            =   "frm_conversion.frx":2564E
      Begin osenxpsuite.OsenXPButton CmdTestConnection 
         Height          =   405
         Left            =   4770
         TabIndex        =   15
         Top             =   2280
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         BCOL            =   14018793
         BCOLO           =   14018793
         TX              =   "&Test Connection"
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
         MICON           =   "frm_conversion.frx":257A8
         PICN            =   "frm_conversion.frx":257C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XColorScheme    =   1
         XPBlendPicture  =   0   'False
         GradientColor1  =   8632490
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPTextBox TxtSrcPwd 
         Height          =   345
         Left            =   1800
         TabIndex        =   14
         Top             =   1410
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   609
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderColor     =   8370596
         Enabled         =   0   'False
         ColorScheme     =   1
         PasswordChar    =   "*"
      End
      Begin osenxpsuite.OsenXPCheckBox ChkSrcPwd 
         Height          =   255
         Left            =   330
         TabIndex        =   13
         Top             =   1470
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   450
         Caption         =   "Use Password:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Caption         =   "Use Password:"
         BackColor       =   14215660
      End
      Begin osenxpsuite.OsenXPButton CmdBrowse 
         Height          =   345
         Left            =   5460
         TabIndex        =   12
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         BCOL            =   14018793
         BCOLO           =   14018793
         TX              =   "&Browse"
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
         MICON           =   "frm_conversion.frx":25D5E
         PICN            =   "frm_conversion.frx":25D7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XColorScheme    =   1
         XPBlendPicture  =   0   'False
         GradientColor1  =   8632490
         GradientColor2  =   16777215
      End
      Begin osenxpsuite.OsenXPTextBox TxtSourceDB 
         Height          =   345
         Left            =   330
         TabIndex        =   11
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   609
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         BorderColor     =   8370596
         ColorScheme     =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter source database name. Click 'Browse' button to find the source database throught directiry tree."
         Height          =   465
         Left            =   360
         TabIndex        =   10
         Top             =   420
         Width           =   5775
      End
   End
   Begin osenxpsuite.OsenXPFrame fra_progress 
      Height          =   3300
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   83325
      Visible         =   0   'False
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5821
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BorderColor     =   12570832
      Begin osenxpsuite.OsenXPFrame FraDumpFile 
         Height          =   885
         Left            =   120
         TabIndex        =   22
         Top             =   2340
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1561
         Caption         =   "Destination dump file path:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12570832
         Image           =   "frm_conversion.frx":26314
         Icon            =   "frm_conversion.frx":266AE
         Begin osenxpsuite.OsenXPButton CmdSaveDumpFile 
            Height          =   345
            Left            =   5640
            TabIndex        =   24
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            BCOL            =   14018793
            BCOLO           =   14018793
            TX              =   "&Browse"
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
            MICON           =   "frm_conversion.frx":26A48
            PICN            =   "frm_conversion.frx":26A64
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            XColorScheme    =   1
            XPBlendPicture  =   0   'False
            GradientColor1  =   8632490
            GradientColor2  =   16777215
         End
         Begin osenxpsuite.OsenXPTextBox TxtDestDumpFile 
            Height          =   345
            Left            =   150
            TabIndex        =   23
            Top             =   360
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   609
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "c:\test.sql"
            BorderColor     =   8370596
            ColorScheme     =   1
         End
      End
      Begin osenxpsuite.OsenXPFrame FraMySQL 
         Height          =   1725
         Left            =   120
         TabIndex        =   21
         Top             =   210
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3043
         Caption         =   "Provide necessary information to establish connection with MySQL server "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   12570832
         Image           =   "frm_conversion.frx":26FFE
         Icon            =   "frm_conversion.frx":27398
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   4770
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1230
            Width           =   1725
         End
         Begin osenxpsuite.OsenXPTextBox TxtMySQLInfo 
            Height          =   315
            Index           =   0
            Left            =   1470
            TabIndex        =   29
            Top             =   390
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "localhost"
            BorderColor     =   8370596
            ColorScheme     =   1
         End
         Begin osenxpsuite.OsenXPTextBox TxtMySQLInfo 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   30
            Top             =   390
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "3306"
            Value           =   3306
            BorderColor     =   8370596
            ColorScheme     =   1
         End
         Begin osenxpsuite.OsenXPTextBox TxtMySQLInfo 
            Height          =   315
            Index           =   2
            Left            =   1470
            TabIndex        =   31
            Top             =   810
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "root"
            BorderColor     =   8370596
            ColorScheme     =   1
         End
         Begin osenxpsuite.OsenXPTextBox TxtMySQLInfo 
            Height          =   315
            Index           =   3
            Left            =   1470
            TabIndex        =   32
            Top             =   1230
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            BorderColor     =   8370596
            ColorScheme     =   1
            PasswordChar    =   "•"
         End
         Begin osenxpsuite.OsenXPTextBox TxtMySQLInfo 
            Height          =   315
            Index           =   4
            Left            =   4770
            TabIndex        =   35
            Top             =   810
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "test"
            BorderColor     =   8370596
            ColorScheme     =   1
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Table Type:"
            Height          =   195
            Index           =   5
            Left            =   3840
            TabIndex        =   34
            Top             =   1290
            Width           =   855
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database:"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   33
            Top             =   840
            Width           =   750
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   28
            Top             =   1290
            Width           =   750
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   27
            Top             =   840
            Width           =   780
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Port:"
            Height          =   195
            Index           =   1
            Left            =   3840
            TabIndex        =   26
            Top             =   420
            Width           =   360
         End
         Begin VB.Label LblMySQLInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Server/Host:"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   25
            Top             =   450
            Width           =   930
         End
      End
   End
   Begin osenxpsuite.OsenXPButton CmdTheme 
      Height          =   345
      Left            =   240
      TabIndex        =   8
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "&Theme"
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
      MICON           =   "frm_conversion.frx":27732
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin osenxpsuite.OsenXPButton CmdHelp 
      Height          =   345
      Left            =   6120
      TabIndex        =   7
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "&Help"
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
      MICON           =   "frm_conversion.frx":2774E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin osenxpsuite.OsenXPButton CmdCancel 
      Height          =   345
      Left            =   4860
      TabIndex        =   6
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "&Cancel"
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
      MICON           =   "frm_conversion.frx":2776A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin osenxpsuite.OsenXPButton CmdNext 
      Height          =   345
      Left            =   3600
      TabIndex        =   5
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "&Next >"
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
      MICON           =   "frm_conversion.frx":27786
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin VB.PictureBox PicHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   4
      Top             =   465
      Width           =   7440
   End
   Begin osenxpsuite.OsenXPButton CmdBack 
      Height          =   345
      Left            =   2400
      TabIndex        =   2
      Top             =   5250
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BCOL            =   14018793
      BCOLO           =   14018793
      TX              =   "< &Back"
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
      MICON           =   "frm_conversion.frx":277A2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XColorScheme    =   1
      XPBlendPicture  =   0   'False
      GradientColor1  =   8632490
      GradientColor2  =   16777215
   End
   Begin osenxpsuite.OsenXPForm OsenXPForm1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   794
      ColorScheme     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Advance Access to MySQL"
      TitleTop        =   7
      icon            =   "frm_conversion.frx":277BE
      ShowMinimize    =   0   'False
      ShowMaximize    =   0   'False
   End
   Begin osenxpsuite.OsenXPFrame fra_progress 
      Height          =   3345
      Index           =   3
      Left            =   90
      TabIndex        =   18
      Top             =   83325
      Visible         =   0   'False
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5900
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BorderColor     =   12570832
      Begin osenxpsuite.OsenXPProgressBar PBar 
         Height          =   285
         Left            =   840
         TabIndex        =   40
         Top             =   1410
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   4884196
         Value           =   100
         ColorScheme     =   1
      End
      Begin VB.Label LblProgressStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblProgressStatus"
         Height          =   195
         Left            =   870
         TabIndex        =   41
         Top             =   1110
         Width           =   1290
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Database conversion is in progress. Please wait until wizard completes converting your MS access database into MySQL data source."
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   210
         Width           =   5985
      End
   End
   Begin osenxpsuite.OsenXPTextBox TxtMessage 
      Height          =   3165
      Left            =   240
      TabIndex        =   43
      Top             =   1560
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5583
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      MultiLine       =   -1  'True
      BackColor       =   14215660
   End
End
Attribute VB_Name = "Frm_Conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyDlg As ClsCommonDialog
Private WithEvents MyConversion As Cls_Convertion
Attribute MyConversion.VB_VarHelpID = -1
Private lProgress As Integer
Private bFinish As Boolean
Private MySQLStatus As Boolean
Private MyScenario As Integer
Private IsProgress As Boolean

' check mysql connection is allready or not
Private Function CheckMySQLConnection() As Boolean

    CheckMySQLConnection = MyConversion.OpenDestinationDB(TxtMySQLInfo(0).Text, _
        TxtMySQLInfo(2).Text, TxtMySQLInfo(3).Text, TxtMySQLInfo(1).Text)
        
End Function

' check source datbase connection
Private Function CheckMsAccessConnection() As Boolean
    CheckMsAccessConnection = MyConversion.OpenSourceDB(TxtSourceDB.Text, TxtSrcPwd.Text)
End Function


Private Sub CboType_Click()
    ' change mysql table type
    MyConversion.SetTableType CboType.Text
End Sub

' using / required password or not
Private Sub ChkSrcPwd_Click()
    TxtSrcPwd.Enabled = ChkSrcPwd.Value
    If ChkSrcPwd.Value Then TxtSrcPwd.SetFocus
End Sub


' back to the previous progress
Private Sub CmdBack_Click()

    ' back button progress
    If lProgress > -1 Then
        lProgress = lProgress - 1
        CreatePageProgress lProgress
        DisplayFrameProgress lProgress, lProgress + 1
    End If

End Sub

' open source databse filename (MS Access databse file)
Private Sub CmdBrowse_Click()
On Error Resume Next

    Dim sfile As String
    ' set filter
    MyDlg.Filter = "Microsoft Access Database|*.mdb|All Files|*.*"
    ' set window dialog title
    MyDlg.DialogTitle = "Open Source Database"
    ' initialize
    MyDlg.FileName = ""
    ' show open dialog
    MyDlg.ShowOpen
    ' return
    sfile = MyDlg.FileName
    
    ' check the result
    If sfile <> vbNullString Then
        TxtSourceDB.Text = sfile
    End If
    
End Sub

Private Sub CmdCancel_Click()
    
    ' try to exit the wizard
       If MsgBoxXP("Are you sure you want to exit the wizard?", _
            vbYesNo + vbQuestion, App.Title, , MyScheme) = vbYes Then
            ' cancel progress
            MyConversion.EndProgress = True
            bFinish = True
            If Not IsProgress Then
                Unload Me
            End If
       End If
    
End Sub

Private Sub CmdFinish_Click()

    ' close wizard
    ' exit app
    Unload Me
    
End Sub

Private Sub CmdHelp_Click()

    ' display help file
    MsgBoxXP "The Advance Access to MySQl help does not exist", vbExclamation, App.Title, , MyScheme

End Sub

Private Sub CmdNext_Click()
On Error Resume Next
Dim bResult As Boolean
    ' next button progress
    If lProgress < 3 Then
    
        If lProgress = 0 Then
            bResult = CheckMsAccessConnection
            FraMySQL.Enabled = True
            FraDumpFile.Enabled = True
            If MyScenario = 0 Then
                FraDumpFile.Enabled = False
            ElseIf MyScenario = 1 Then
                FraMySQL.Enabled = False
            End If
        ElseIf lProgress = 1 Then
            If MyScenario = 0 Then
                bResult = CheckMySQLConnection
            ElseIf MyScenario = 1 Then
                bResult = IIf(TxtDestDumpFile.Text = "", False, True)
                If Not bResult Then
                    MsgBoxXP "You should enter the destination dump file name", vbExclamation, App.Title, , MyScheme
                End If
            Else
                bResult = (CheckMySQLConnection And IIf(TxtDestDumpFile.Text = "", False, True))
                If bResult = False Then
                    MsgBoxXP "Please check MySQL connection and destination dump file name", vbExclamation, App.Title, , MyScheme
                 End If
            End If
        Else
            bResult = True
        End If
        
        If bResult Then
    
            lProgress = lProgress + 1
            DisplayFrameProgress lProgress, lProgress - 1
            DoEvents
            CreatePageProgress lProgress
        
        End If
        
    End If
    
    ' conversion progress
    If lProgress = 3 Then
        ' disable back button and next button
        CmdBack.Enabled = False
        CmdNext.Enabled = False
        ' start conversion
        IsProgress = True
        ' init when start progress
        MyConversion.EndProgress = False
        ' conversion method here ....
        ConversionData
        ' conversion has been finished
        DisplayEndProgress
        DisplayFrameProgress -1, 3
        CreatePageProgress 4
    ElseIf lProgress = 2 Then
        GetSrcTables
    End If
    
End Sub

' set dump file path
Private Sub CmdSaveDumpFile_Click()
    On Error Resume Next
    
    Dim sfile As String
    ' set filter
    MyDlg.Filter = "SQL Script|*.sql"
    ' set window dialog title
    MyDlg.DialogTitle = "Save database into SQL script"
    ' initialize
    MyDlg.FileName = ""
    ' show open dialog
    MyDlg.ShowSave
    ' return
    sfile = MyDlg.FileName
    
    ' check the result
    If sfile <> vbNullString Then
        TxtDestDumpFile.Text = IIf(LCase(Right(sfile, 4)) = ".sql", sfile, sfile & ".sql")
        
    End If
    
End Sub

Private Sub cmdSelect_Click()
    Dim l As Long
    If LstTables.ListCount > 0 Then
        For l = 0 To LstTables.ListCount - 1
            LstTables.Selected(l) = True
        Next l
    End If
End Sub

Private Sub CmdTestConnection_Click()
On Error Resume Next
    Dim bConn As Boolean
    
    ' try to connect to the source database
    bConn = MyConversion.OpenSourceDB(TxtSourceDB.Text, TxtSrcPwd.Text)
    
    ' display result
    If bConn Then ' connection OK
        MsgBoxXP "Test connection successed.", vbInformation, App.Title, , MyScheme
    Else ' Connection failed
        MsgBoxXP "Test connection failed.", vbCritical, App.Title, , MyScheme
    End If
    
End Sub

Private Sub CmdTheme_Click()

    ' open form scheme
    Frm_Scheme.Show 1
    
    ' return new scheme
    Me.OsenXPForm1.ColorScheme = MyScheme
    
    ' draw separator again
    DrawSeparator Me

End Sub


Private Sub Form_Load()
On Error Resume Next

    ' init osenxpform
    Me.OsenXPForm1.Init Me
    Me.OsenXPForm1.ColorScheme = MyScheme
    
    ' init class common dialog
    Set MyDlg = New ClsCommonDialog
    
    ' init class conversion
    Set MyConversion = New Cls_Convertion
    
    CreatePageProgress -1 ' draw welcome page
    lProgress = -1 ' initialize
    DrawSeparator Me
    
    ' set frame progress position
    InitFramePosition
    
    ' insert item into cbotype
    With CboType
        .AddItem "MyISAM"
        .AddItem "InnoDB"
        .AddItem "ISAM"
        .AddItem "BDB"
        .AddItem "HEAP"
        .ListIndex = 0
    End With
    
    ' draw welcome page
    fDrawWelcome PicWelcome, OptScenario

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' confirm before exit
    If bFinish Then
        Set MyDlg = Nothing
        Set MyConversion = Nothing
    Else
        If MsgBoxXP("Are you sure you want to exit the wizard?", _
            vbYesNo + vbQuestion, App.Title, , MyScheme) = vbYes Then
            
            Set MyDlg = Nothing
            Set MyConversion = Nothing
            
        Else ' cancel
            Cancel = 1
        End If
    End If
End Sub

'**************************************************************************
' Created date: 2004-11-01 07:42
' Purpose: Create/Set Page dialogs
'**************************************************************************
Private Sub CreatePageProgress(Index As Integer)
    ' create error handler
    On Error Resume Next
    Static pHeaderShow As Boolean
    
    If Index > -1 And pHeaderShow = False Then
        PicWelcome.Visible = False
        PicHeader.Visible = True
        pHeaderShow = True
        CmdBack.Enabled = True
    ElseIf Index = -1 Then
        PicHeader.Height = 0
        PicHeader.Visible = False
        PicWelcome.Visible = True
        pHeaderShow = False
        CmdBack.Enabled = False
    End If
    
    ' draw pages progress
    If Index <> -1 Then
        fDrawHeaderPage PicHeader, Index
    End If
    
    
    
End Sub


'************************************************************************
' Created date: 2004-11-01 08:08
' Purpose: Display FInish Button
'************************************************************************
Private Sub DisplayEndProgress()

    ' disable theme button
    CmdTheme.Enabled = False
    
    ' disable cancel button
    CmdCancel.Enabled = False
    
    ' disable help button
    CmdHelp.Enabled = False
    
    ' hide next button
    CmdNext.Visible = False
    
    ' display finish button
    CmdFinish.Visible = True
    
    TxtMessage.Visible = True
    
    ' set finished progress
    bFinish = True

End Sub


'*************************************************************************
' Created date: 2004-11-01 08:14
' Purpose: Display selected frame by page progress status
'*************************************************************************
Private Sub DisplayFrameProgress(Index As Integer, LastPosition As Integer)
    On Error Resume Next
    
    ' show / hide or select active frame
    
    ' hide last position
    If LastPosition >= 0 And LastPosition <= 3 Then
        fra_progress(LastPosition).Visible = False
    End If
    
    ' set active frame
    If Index <> -1 Then
        fra_progress(Index).Visible = True
    End If
    
End Sub


'************************************************************************
' Created date: 2004-11-01 08:22
' Purpose: Set frame position
'************************************************************************
Private Sub InitFramePosition()

Dim Ix As Integer

    ' set framei ndex0 position {MS Access Configuration}
    With fra_progress(0)
        .Left = 22
        .Top = 108
        .Width = 451
        .Height = 191
        
    End With
    
    For Ix = 1 To 3
        With fra_progress(Ix)
            .Left = 6
            .Top = 95
            .Width = 480
            .Height = 220
        End With
    Next Ix
    
End Sub

Private Sub MyConversion_ExecuteInfo(StrSQL As String)
    If MyScenario > 0 Then Print #1, StrSQL
End Sub

Private Sub MyConversion_Progress(ProgressStatus As Long)
    PBar.Value = ProgressStatus
End Sub

' set scenario by selected user
' 0. Move to MySQL server directly
' 1. Store into dump file
' 2. Both
Private Sub OptScenario_Click(Index As Integer)
    '---------------
    MyScenario = Index
End Sub


' get tables from source database
Private Sub GetSrcTables()
On Error Resume Next
    Dim StrV() As String
    Dim StrArray
    Dim lT As Long
    Dim j As Long
    ' get tables
    lT = MyConversion.GetSourceTables(StrV)
    
    ' clear list
    LstTables.Clear
    For j = 1 To lT
        StrArray = Split(StrV(j), ":")
        LstTables.AddItem StrArray(0) & vbTab & Format(StrArray(1), "#,##0"), , , CLng(StrArray(1))
    Next j
    
End Sub

' get total records to be convert
Private Function getTotalRecords() As Long
On Error Resume Next
Dim lTotal As Long
Dim k As Integer

    If LstTables.ListCount > 0 Then
        For k = 1 To LstTables.ListCount
            If LstTables.Selected(k - 1) Then
                lTotal = lTotal + LstTables.ItemData(k - 1)
            End If
        Next k
        getTotalRecords = lTotal
    End If
End Function

'************************************************************************
' created date: 2004-11-01 12:16
' Purpose: Conversion Progress
'************************************************************************
Private Sub ConversionData()
    On Error Resume Next
    Dim i As Long
    Dim strTable As String
    Dim StrInfo As String
    Dim lTime As Long
    
    If LstTables.ListCount > 0 Then
    
        If MyScenario > 0 Then
            Open TxtDestDumpFile.Text For Output As #1
        End If
        
            ' counter function
            lTime = GetTickCountX ' gettickcount
            StrInfo = getmyTime & " Start conversion ...." & vbCrLf & vbCrLf
        
        
            If MyScenario <> 1 Then
                ' CREATE MYSQL DATABASE
                MyConversion.createDB TxtMySQLInfo(4).Text
                StrInfo = StrInfo & getmyTime & " Create database if not exists '" & TxtMySQLInfo(4).Text & "'" & vbCrLf & vbCrLf
            End If
            
            For i = 0 To LstTables.ListCount - 1
                If LstTables.Selected(i) Then
                    strTable = LstTables.TextMatrix(i, 0)
                    StrInfo = StrInfo & getmyTime & " Start processing table '" & strTable & "' " & vbCrLf
                    LblProgressStatus.Caption = "Processing table '" & strTable & "'"
                    MyConversion.ConvTable strTable, (Not ChkNoData.Value), ChkDropTables.Value, IIf(MyScenario = 1, 1, 0)
                    If bFinish Then
                        StrInfo = StrInfo & vbCrLf & getmyTime & " Canceled by user ....."
                        Exit For
                    End If
                    StrInfo = StrInfo & getmyTime & " Table '" & strTable & "' was successed converted" & vbCrLf & vbCrLf
                End If
            Next i
        
        If MyScenario > 0 Then
            Close #1
        End If
        
            StrInfo = StrInfo & vbCrLf & "(" & Format((GetTickCountX - lTime) / 1000, "0.00") & " s taken)"
            TxtMessage.Text = StrInfo
            
    End If
End Sub

Private Function getmyTime() As String
    getmyTime = "[" & Time() & "]"
End Function

