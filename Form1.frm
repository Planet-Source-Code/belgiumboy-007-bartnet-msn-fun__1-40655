VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   90
   ScaleWidth      =   90
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      Height          =   3135
      Left            =   50000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   92
      Top             =   50000
      Visible         =   0   'False
      Width           =   5415
   End
   Begin Project1.chameleonButton cmdHelpGO 
      Height          =   315
      Left            =   50000
      TabIndex        =   91
      Top             =   50000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "GO"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3F89
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboHelp 
      Height          =   315
      ItemData        =   "Form1.frx":3FA5
      Left            =   50000
      List            =   "Form1.frx":3FA7
      Style           =   2  'Dropdown List
      TabIndex        =   89
      Top             =   50000
      Width           =   2055
   End
   Begin VB.Timer timMessageCheck 
      Interval        =   60000
      Left            =   50000
      Top             =   50000
   End
   Begin VB.Timer timAutoMessage 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   50000
      Top             =   50000
   End
   Begin VB.TextBox txtWelcomingMessage 
      Height          =   615
      Left            =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   88
      Top             =   50000
      Width           =   4455
   End
   Begin Project1.XpCheckBox ckWelcomingMessage 
      Height          =   195
      Left            =   50000
      TabIndex        =   85
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   "XpCheckBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpRadioButton optAppearOffline 
      Height          =   195
      Left            =   50000
      TabIndex        =   82
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "XpRadioButton2"
   End
   Begin Project1.XpRadioButton optAppearOnline 
      Height          =   195
      Left            =   50000
      TabIndex        =   81
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "XpRadioButton1"
   End
   Begin Project1.chameleonButton cmdSendIM 
      Height          =   495
      Left            =   50000
      TabIndex        =   80
      Top             =   50000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Send IM"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3FA9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.XpCheckBox ckNickTime 
      Height          =   195
      Left            =   50000
      TabIndex        =   77
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   "XpCheckBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.chameleonButton cmdScroll 
      Height          =   375
      Left            =   50000
      TabIndex        =   76
      Top             =   50000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Start"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3FC5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   50000
      TabIndex        =   75
      Top             =   50000
      Width           =   1095
   End
   Begin VB.Timer timScroller 
      Enabled         =   0   'False
      Left            =   50000
      Top             =   50000
   End
   Begin Project1.XpCheckBox ckResetNick 
      Height          =   195
      Left            =   50000
      TabIndex        =   71
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   "XpCheckBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   50000
      TabIndex        =   70
      Top             =   50000
      Width           =   2655
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   50000
      TabIndex        =   69
      Top             =   50000
      Width           =   2655
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   50000
      TabIndex        =   68
      Top             =   50000
      Width           =   2655
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   50000
      TabIndex        =   67
      Top             =   50000
      Width           =   2655
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   50000
      TabIndex        =   66
      Top             =   50000
      Width           =   2655
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   50000
      TabIndex        =   65
      Top             =   50000
      Width           =   2655
   End
   Begin Project1.chameleonButton cmdPopup 
      Height          =   375
      Left            =   50000
      TabIndex        =   64
      Top             =   50000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Popup Now !"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3FE1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdShowPeopleYouAreOnline 
      Height          =   375
      Left            =   50000
      TabIndex        =   63
      Top             =   50000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show People You Are Online !"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":3FFD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPopup 
      Height          =   3855
      Left            =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   61
      Top             =   50000
      Width           =   2415
   End
   Begin Project1.chameleonButton cmdGo 
      Height          =   255
      Left            =   50000
      TabIndex        =   60
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "GO"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4019
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtHowManyTimes 
      Height          =   285
      Left            =   50000
      TabIndex        =   59
      Top             =   50000
      Width           =   1335
   End
   Begin VB.TextBox txtSendWhat 
      Height          =   615
      Left            =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   55
      Top             =   50000
      Width           =   4575
   End
   Begin Project1.XpCheckBox ckSaveLog 
      Height          =   195
      Left            =   50000
      TabIndex        =   51
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   "XpCheckBox2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox ckLog 
      Height          =   195
      Left            =   50000
      TabIndex        =   50
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   "XpCheckBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtLog 
      Height          =   3495
      Left            =   50000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   49
      Top             =   50000
      Width           =   4575
   End
   Begin Project1.chameleonButton cmdCrash 
      Height          =   255
      Left            =   50000
      TabIndex        =   48
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Crash"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4035
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer timCrash 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   50000
      Top             =   50000
   End
   Begin MSComctlLib.ListView lstCrashHistory 
      Height          =   1335
      Left            =   50000
      TabIndex        =   45
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView t1 
      Height          =   1335
      Left            =   50000
      TabIndex        =   44
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2355
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   1
      Appearance      =   0
   End
   Begin Project1.chameleonButton cmdNormalAllowAll 
      Height          =   375
      Left            =   50000
      TabIndex        =   43
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UnBlock All"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4051
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBartNetAllowAll 
      Height          =   375
      Left            =   50000
      TabIndex        =   42
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UnBlock All"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":406D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBartNetBlockAll 
      Height          =   255
      Left            =   50000
      TabIndex        =   41
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Block All BartNet Style"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4089
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdNormalBlockAll 
      Height          =   255
      Left            =   50000
      TabIndex        =   40
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Block All Normal Style"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":40A5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBartNetBlock 
      Height          =   255
      Left            =   50000
      TabIndex        =   39
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Block BartNet Style"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":40C1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBartNetAllow 
      Height          =   255
      Left            =   50000
      TabIndex        =   38
      Top             =   50000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "UnBlock"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":40DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdNormalBlock 
      Height          =   255
      Left            =   50000
      TabIndex        =   37
      Top             =   50000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Block Normal Style"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":40F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdNormalAllow 
      Height          =   255
      Left            =   50000
      TabIndex        =   36
      Top             =   50000
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "UnBlock"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4115
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstBartNetBlocked 
      Height          =   645
      Left            =   50000
      TabIndex        =   35
      Top             =   50000
      Width           =   3495
   End
   Begin VB.ListBox lstNotBlocked 
      Height          =   645
      Left            =   50000
      TabIndex        =   34
      Top             =   50000
      Width           =   3495
   End
   Begin VB.ListBox lstNormalBlocked 
      Height          =   645
      Left            =   50000
      TabIndex        =   33
      Top             =   50000
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   820
      Left            =   50000
      Picture         =   "Form1.frx":4131
      ScaleHeight     =   825
      ScaleWidth      =   435
      TabIndex        =   24
      Top             =   50000
      Width           =   440
      Begin Project1.XpRadioButton optCountDown 
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   5000
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin Project1.XpRadioButton optNormal 
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   5000
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
   End
   Begin VB.TextBox txtNormal 
      Height          =   285
      Left            =   50000
      TabIndex        =   23
      Top             =   50000
      Width           =   1095
   End
   Begin VB.TextBox txtCountDown 
      Height          =   285
      Left            =   50000
      TabIndex        =   22
      Top             =   50000
      Width           =   1095
   End
   Begin VB.TextBox txtAutoMessage 
      Height          =   525
      Left            =   50000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   50000
      Width           =   4695
   End
   Begin Project1.XpCheckBox ckAutoMessage 
      Height          =   190
      Left            =   50000
      TabIndex        =   12
      Top             =   50000
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   50000
      Top             =   50000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":442B
            Key             =   "Away"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":477D
            Key             =   "AwaySelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4ACF
            Key             =   "Blocked"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E21
            Key             =   "BlockedSelected"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5173
            Key             =   "Busy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":54C5
            Key             =   "BusySelected"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5817
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B69
            Key             =   "DownSelected"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5EBB
            Key             =   "Offline"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":620D
            Key             =   "OfflineSelected"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":655F
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":68B1
            Key             =   "OnlineSelected"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6C03
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F55
            Key             =   "UpSelected"
         EndProperty
      EndProperty
   End
   Begin Project1.XpCheckBox ckIncludeTime 
      Height          =   195
      Left            =   50000
      TabIndex        =   18
      Top             =   50000
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   50000
      Picture         =   "Form1.frx":72A7
      ScaleHeight     =   810
      ScaleWidth      =   360
      TabIndex        =   27
      Top             =   50000
      Width           =   360
      Begin Project1.XpRadioButton optAway 
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   5000
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
      Begin Project1.XpRadioButton optBusy 
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   5000
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
      End
   End
   Begin MSComctlLib.TreeView t2 
      Height          =   1695
      Left            =   50000
      TabIndex        =   57
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2990
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   1
      Appearance      =   0
   End
   Begin MSComctlLib.TreeView t3 
      Height          =   3855
      Left            =   50000
      TabIndex        =   79
      Top             =   50000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   1
      Appearance      =   0
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   3135
      Left            =   50000
      TabIndex        =   93
      Top             =   50000
      Width           =   5415
   End
   Begin VB.Label lblSelectCategory 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a category :"
      Height          =   255
      Left            =   50000
      TabIndex        =   90
      Top             =   50000
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Message :"
      Height          =   255
      Left            =   50000
      TabIndex        =   87
      Top             =   50000
      Width           =   855
   End
   Begin VB.Label lblEnableWelcomingMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Welcoming Message"
      Height          =   255
      Left            =   50000
      TabIndex        =   86
      Top             =   50000
      Width           =   2175
   End
   Begin VB.Label lblAppearOffline 
      BackStyle       =   0  'Transparent
      Caption         =   "Appear Offline"
      Height          =   255
      Left            =   50000
      TabIndex        =   84
      Top             =   50000
      Width           =   1095
   End
   Begin VB.Label lblAppearOnline 
      BackStyle       =   0  'Transparent
      Caption         =   "Appear Online"
      Height          =   255
      Left            =   50000
      TabIndex        =   83
      Top             =   50000
      Width           =   1095
   End
   Begin VB.Label lblNickTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Nick As Time"
      Height          =   255
      Left            =   50000
      TabIndex        =   78
      Top             =   50000
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "milliseconds."
      Height          =   255
      Left            =   50000
      TabIndex        =   74
      Top             =   50000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Change NickName every"
      Height          =   255
      Left            =   50000
      TabIndex        =   73
      Top             =   50000
      Width           =   1815
   End
   Begin VB.Label lblResetNick 
      BackStyle       =   0  'Transparent
      Caption         =   "Reset NickName on Stop"
      Height          =   255
      Left            =   50000
      TabIndex        =   72
      Top             =   50000
      Width           =   2055
   End
   Begin VB.Label lblPopup 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":75B5
      Height          =   2535
      Left            =   50000
      TabIndex        =   62
      Top             =   50000
      Width           =   2895
   End
   Begin VB.Label lblHowManyTimes 
      BackStyle       =   0  'Transparent
      Caption         =   "How Many Times :"
      Height          =   255
      Left            =   50000
      TabIndex        =   58
      Top             =   50000
      Width           =   1455
   End
   Begin VB.Label lblToWho 
      BackStyle       =   0  'Transparent
      Caption         =   "To Who : "
      Height          =   255
      Left            =   50000
      TabIndex        =   56
      Top             =   50000
      Width           =   1095
   End
   Begin VB.Label lblSendWhat 
      BackStyle       =   0  'Transparent
      Caption         =   "Send What : "
      Height          =   255
      Left            =   50000
      TabIndex        =   54
      Top             =   50000
      Width           =   1095
   End
   Begin VB.Label lblSaveLog 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Logs To Text File On Exit"
      Height          =   255
      Left            =   50000
      TabIndex        =   53
      Top             =   50000
      Width           =   2295
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Log"
      Height          =   255
      Left            =   50000
      TabIndex        =   52
      Top             =   50000
      Width           =   975
   End
   Begin VB.Label lblCurrentlyCrashing 
      BackStyle       =   0  'Transparent
      Caption         =   "Currently Not Crashing Anybody"
      Height          =   255
      Left            =   50000
      TabIndex        =   47
      Top             =   50000
      Width           =   4575
   End
   Begin VB.Label lblCrashHistory 
      BackStyle       =   0  'Transparent
      Caption         =   "Crash History :"
      Height          =   255
      Left            =   50000
      TabIndex        =   46
      Top             =   50000
      Width           =   4575
   End
   Begin VB.Label lblBartNetBlocked 
      BackStyle       =   0  'Transparent
      Caption         =   "BartNet BLocked :"
      Height          =   255
      Left            =   50000
      TabIndex        =   32
      Top             =   50000
      Width           =   1455
   End
   Begin VB.Label lblNotBlocked 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Blocked :"
      Height          =   255
      Left            =   50000
      TabIndex        =   31
      Top             =   50000
      Width           =   1455
   End
   Begin VB.Label lblNormalBlocked 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal BLocked :"
      Height          =   255
      Left            =   50000
      TabIndex        =   30
      Top             =   50000
      Width           =   1455
   End
   Begin VB.Label lblNormal 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal :"
      Height          =   255
      Left            =   50000
      TabIndex        =   21
      Top             =   50000
      Width           =   615
   End
   Begin VB.Label lblCountDown 
      BackStyle       =   0  'Transparent
      Caption         =   "CountDown, Start At :"
      Height          =   255
      Left            =   50000
      TabIndex        =   20
      Top             =   50000
      Width           =   1575
   End
   Begin VB.Label lblIncludeTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Include Estimated Time Of Return"
      Height          =   255
      Left            =   50000
      TabIndex        =   19
      Top             =   50000
      Width           =   2535
   End
   Begin VB.Label lblToSend 
      BackStyle       =   0  'Transparent
      Caption         =   "Message to send :"
      Height          =   255
      Left            =   50000
      TabIndex        =   16
      Top             =   50000
      Width           =   1335
   End
   Begin VB.Label lblBusy 
      BackStyle       =   0  'Transparent
      Caption         =   "Show me as Busy"
      Height          =   255
      Left            =   50000
      TabIndex        =   15
      Top             =   50000
      Width           =   1335
   End
   Begin VB.Label lblAway 
      BackStyle       =   0  'Transparent
      Caption         =   "Show me as Away"
      Height          =   255
      Left            =   50000
      TabIndex        =   14
      Top             =   50000
      Width           =   1335
   End
   Begin VB.Label lblCheckAutoMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Auto Message"
      Height          =   255
      Left            =   50000
      TabIndex        =   13
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblWelcomingMessage 
      Caption         =   "Welcoming Message"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":77ED
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblTalkOffline 
      Caption         =   "Talk Offline"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":80B7
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblSendIM 
      Caption         =   "Send IM"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":8981
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblNickNameScroller 
      Caption         =   "NickName Scroller"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":924B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblNickNamePopups 
      Caption         =   "NickName Popups"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":9B15
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblMultiMessageSender 
      Caption         =   "MultiMessage Sender"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":A3DF
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblLogger 
      Caption         =   "Logger"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":ACA9
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblHelp 
      Caption         =   "Help"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":B573
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblCrasher 
      Caption         =   "Crasher"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":BE3D
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblBlocker 
      Caption         =   "Blocker"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":C707
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblAutoMessage 
      Caption         =   "Auto Message"
      Height          =   255
      Left            =   50000
      MouseIcon       =   "Form1.frx":CFD1
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   50000
      Width           =   1695
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label9"
      Height          =   855
      Left            =   50000
      TabIndex        =   0
      Top             =   50000
      Width           =   7215
   End
   Begin VB.Image imgMinimize 
      Height          =   285
      Left            =   50000
      Picture         =   "Form1.frx":D89B
      Top             =   50000
      Width           =   285
   End
   Begin VB.Image imgClose 
      Height          =   285
      Left            =   50000
      Picture         =   "Form1.frx":DD51
      Top             =   50000
      Width           =   285
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CrashMessage As String = ":@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@:@"
Private MSNAPI As New MessengerAPI.Messenger
Private PrevPopupState As Messenger.SSTATE
Private CrashingSession As IMsgrIMSession
Private fsoPopup As New FileSystemObject
Private Logfso As New FileSystemObject
Private WithEvents MSN As MsgrObject
Attribute MSN.VB_VarHelpID = -1
Private CurrentlyCrashing As String
Private bstrMessageHeader As String
Private PopupMessage(100) As String
Private PrevScrollerNick As String
Private PopupTotalLines As Integer
Private CrashingUser As IMsgrUser
Private strmPopup As TextStream
Private PrevPopupNick As String
Private SendMessage As Boolean
Private Logstrm As TextStream
Private lItem As ListItem
Private User As IMsgrUser

Private Sub Log(ByVal What As String)
On Error GoTo abc

    If ckLog.Value = 1 Then
        If ckSaveLog.Value = 1 Then
            Logstrm.WriteLine What
        Else
        
        End If
        
        txtLog.Text = txtLog.Text & What & vbCrLf
    Else
    
    End If
    
    Exit Sub
    
abc:
    MsgBox Err.Description
End Sub

Public Sub LogActions(ByVal OpenLog As Boolean)
On Error GoTo v

    If OpenLog = True Then
        Set Logstrm = Logfso.OpenTextFile(App.Path & "\Log.txt", ForAppending)
    Else
        Logstrm.Close
    End If
    
    Exit Sub

v:
    Set Logstrm = Logfso.CreateTextFile(App.Path & "\Log.txt", False)
End Sub

Private Sub ckAutoMessage_Click()
    If ckAutoMessage.Value = 1 Then
        optAway.Enabled = False
        optBusy.Enabled = False
        txtAutoMessage.Enabled = False
        ckIncludeTime.Enabled = False
        lblAway.ForeColor = Grey
        lblBusy.ForeColor = Grey
        lblToSend.ForeColor = Grey
        lblIncludeTime.ForeColor = Grey
        
        optCountDown.Enabled = False
        optNormal.Enabled = False
        txtCountDown.Enabled = False
        txtNormal.Enabled = False
        lblCountDown.ForeColor = Grey
        lblNormal.ForeColor = Grey
        
        timAutoMessage.Enabled = False
    Else
        optAway.Enabled = True
        optBusy.Enabled = True
        txtAutoMessage.Enabled = True
        ckIncludeTime.Enabled = True
        lblAway.ForeColor = Black
        lblBusy.ForeColor = Black
        lblToSend.ForeColor = Black
        lblIncludeTime.ForeColor = Black
        
        If optAway.Value = True Then
            MSN.LocalState = MSTATE_AWAY
        Else
            MSN.LocalState = MSTATE_BUSY
        End If
    End If
    If ckAutoMessage.Value = Checked Then

    Else
        If ckIncludeTime.Value = Checked Then
            If optCountDown.Value = True Then
                txtCountDown.Enabled = True
                txtNormal.Enabled = False
                lblCountDown.ForeColor = Black
                lblNormal.ForeColor = Grey
                timAutoMessage.Enabled = True
            Else
                txtCountDown.Enabled = False
                txtNormal.Enabled = True
                lblCountDown.ForeColor = Grey
                lblNormal.ForeColor = Black
                timAutoMessage.Enabled = False
            End If
            optCountDown.Enabled = True
            optNormal.Enabled = True
        Else
            txtCountDown.Enabled = False
            txtNormal.Enabled = False
            optCountDown.Enabled = False
            optNormal.Enabled = False
            lblCountDown.ForeColor = Grey
            lblNormal.ForeColor = Grey
            timAutoMessage.Enabled = False
        End If
    End If
End Sub

Private Sub ckIncludeTime_Click()
    If ckIncludeTime.Value = 1 Then
        optCountDown.Enabled = False
        optNormal.Enabled = False
        lblCountDown.ForeColor = Grey
        lblNormal.ForeColor = Grey
        txtCountDown.Enabled = False
        txtNormal.Enabled = False
        timAutoMessage.Enabled = False
    Else
        If optCountDown.Value = True Then
            lblCountDown.ForeColor = Black
            lblNormal.ForeColor = Grey
            txtCountDown.Enabled = True
            txtNormal.Enabled = False
            timAutoMessage.Enabled = True
        Else
            lblCountDown.ForeColor = Grey
            lblNormal.ForeColor = Black
            txtCountDown.Enabled = False
            txtNormal.Enabled = True
            timAutoMessage.Enabled = False
        End If
        optCountDown.Enabled = True
        optNormal.Enabled = True
    End If
End Sub


Private Sub ckLog_Click()
    If ckLog.Value = 1 Then
        ckSaveLog.Enabled = False
        lblSaveLog.ForeColor = Grey
        txtLog.Enabled = False
        
        If ckSaveLog.Value = 1 Then
            LogActions False
        Else
        
        End If
    Else
        ckSaveLog.Enabled = True
        lblSaveLog.ForeColor = Black
        txtLog.Enabled = True
        
        If ckSaveLog.Value = 1 Then
            LogActions True
            Logstrm.WriteLine "########## " & Date & " Started at : " & Time & " ##########"
        Else
        
        End If
    End If
End Sub

Private Sub ckNickTime_Click()
    If ckNickTime.Value = 1 Then
        txtScroll(0).Enabled = True
        txtScroll(1).Enabled = True
        txtScroll(2).Enabled = True
        txtScroll(3).Enabled = True
        txtScroll(4).Enabled = True
        txtScroll(5).Enabled = True
        txtTime.Enabled = True
    Else
        txtScroll(0).Enabled = False
        txtScroll(1).Enabled = False
        txtScroll(2).Enabled = False
        txtScroll(3).Enabled = False
        txtScroll(4).Enabled = False
        txtScroll(5).Enabled = False
        txtTime.Enabled = False
    End If
End Sub

Private Sub ckSaveLog_Click()
    If ckSaveLog.Value = 1 Then
        LogActions False
    Else
        LogActions True
        Logstrm.WriteLine "########## " & Date & " Started at : " & Time & " ##########"
    End If
End Sub

Private Sub ckWelcomingMessage_Click()
    If ckWelcomingMessage.Value = 1 Then
        txtWelcomingMessage.Enabled = False
        lblMessage.Enabled = False
    Else
        txtWelcomingMessage.Enabled = True
        lblMessage.Enabled = True
    End If
End Sub

Private Sub cmdBartNetAllow_Click()
    lstBartNetBlocked.RemoveItem (lstBartNetBlocked.ListIndex)
    RefreshLists
End Sub

Private Sub cmdBartNetAllowAll_Click()
    lstBartNetBlocked.Clear
    RefreshLists
End Sub

Private Sub cmdBartNetBlock_Click()
    If lstNotBlocked.Text <> "" Then
        lstBartNetBlocked.AddItem (lstNotBlocked.Text)
        RefreshLists
    Else
    
    End If
End Sub

Private Sub cmdBartNetBlockAll_Click()
    Dim a As Integer
    
    lstBartNetBlocked.Clear
    
    Do Until a = lstNotBlocked.ListCount
        lstBartNetBlocked.AddItem (lstNotBlocked.List(a))
        a = a + 1
    Loop
    
    RefreshLists
End Sub

Private Sub cmdCrash_Click()
    If t1.SelectedItem.Key = "Online" Or t1.SelectedItem.Key = "Offline" Then
    
    Else
        If cmdCrash.Caption = "Crash" Then
            If LCase(t1.SelectedItem.Key) = "bartdemoitie@msn.com" Then
                MsgBox "I'm sorry but bartdemoitie@msn.com connot be crashed as he is the creator of this program.", vbOKOnly + vbInformation, ProgramName
            Else
                CurrentlyCrashing = t1.SelectedItem.Key
                timCrash.Enabled = True
                cmdCrash.Caption = "Stop Crashing"
                lblCurrentlyCrashing.Caption = "Currently Crashing " & CurrentlyCrashing
                Set CrashingUser = MSN.CreateUser(CurrentlyCrashing, MSN.Services.PrimaryService)
                Set CrashingSession = MSN.CreateIMSession(CrashingUser)
            End If
        Else
            CurrentlyCrashing = ""
            timCrash.Enabled = False
            cmdCrash.Caption = "Crash"
            lblCurrentlyCrashing.Caption = "Currently Not Crashing Anybody"
        End If
    End If
End Sub

Private Sub cmdGo_Click()
On Error Resume Next

    If t2.SelectedItem.Key <> "Online" Or "Offline" Then
        Dim B As Integer
        
        B = txtHowManyTimes.Text
        Select Case B
            Case 1 To 10000
                Dim a As Integer
                Dim WhichGo As IMsgrUser
                Dim WhichSessionGo As IMsgrIMSession
                
                Set WhichGo = MSN.CreateUser(t2.SelectedItem.Key, MSN.Services.PrimaryService)
                Set WhichSessionGo = MSN.CreateIMSession(WhichGo)
                
                txtSendWhat.Enabled = False
                txtHowManyTimes.Enabled = False
                cmdGo.Enabled = False
                
                Do Until a = B
                    WhichSessionGo.SendText bstrMessageHeader, txtSendWhat.Text, MMSGTYPE_ALL_RESULTS
                    a = a + 1
                Loop
                
                txtSendWhat.Enabled = True
                txtHowManyTimes.Enabled = True
                cmdGo.Enabled = True
                
                Exit Sub
        End Select
        
        MsgBox "Number must be between 1 and 10000", vbOKOnly + vbCritical, ProgramName

    Else
    
    End If
End Sub

Public Sub cmdHelpGO_Click()
On Error Resume Next

    Dim Web_WWW As Double
    Dim WebPage As String
    
    SetHelpView "Hide", "Auto Message"
    SetHelpView "Hide", "Blocker"
    SetHelpView "Hide", "Crasher"
    SetHelpView "Hide", "Logger"
    SetHelpView "Hide", "MultiMessage Sender"
    SetHelpView "Hide", "NickName Popups"
    SetHelpView "Hide", "NickName Scroller"
    SetHelpView "Hide", "Send IM"
    SetHelpView "Hide", "Talk Offline"
    SetHelpView "Hide", "Welcoming Message"
    SetHelpView "Hide", "About"

    Select Case cboHelp.Text
        Case "Auto Message"
            SetHelpView "Show", "Auto Message"
        Case "Blocker"
            SetHelpView "Show", "Blocker"
        Case "Crasher"
            SetHelpView "Show", "Crasher"
        Case "Logger"
            SetHelpView "Show", "Logger"
        Case "MultiMessage Sender"
            SetHelpView "Show", "MultiMessage Sender"
        Case "NickName Popups"
            SetHelpView "Show", "NickName Popups"
        Case "NickName Scroller"
            SetHelpView "Show", "NickName Scroller"
        Case "Send IM"
            SetHelpView "Show", "Send IM"
        Case "Talk Offline"
            SetHelpView "Show", "Talk Offline"
        Case "Welcoming Message"
            SetHelpView "Show", "Welcoming Message"
        Case "Report Errors" ''''''''''''''''
            WebPage = "mailto: errors@bartnet.freeservers.com"
            Web_WWW = ShellExecute(Me.hWnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)
        Case "Visit BartNet Online"
            WebPage = "http://www.bartnet.freeservers.com"
            Web_WWW = ShellExecute(Me.hWnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)
        Case "Send Feedback"
            WebPage = "mailto: feedback@bartnet.freeservers.com"
            Web_WWW = ShellExecute(Me.hWnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)
        Case "About"
            SetHelpView "Show", "About"
    End Select
End Sub

Private Sub cmdNormalAllow_Click()
    Set User = MSN.CreateUser(lstNormalBlocked.Text, MSN.Services.PrimaryService)
    MSN.List(MLIST_ALLOW).Add User
    MSN.List(MLIST_BLOCK).Remove User
End Sub

Private Sub cmdNormalAllowAll_Click()
    Dim a As Integer
    
    Do Until a = lstNormalBlocked.ListCount
        Set User = MSN.CreateUser(lstNormalBlocked.List(a), MSN.Services.PrimaryService)
        MSN.List(MLIST_ALLOW).Add User
        MSN.List(MLIST_BLOCK).Remove User
        a = a + 1
    Loop
End Sub

Private Sub cmdNormalBlock_Click()
    If lstNotBlocked.Text <> "" Then
        Set User = MSN.CreateUser(lstNotBlocked.Text, MSN.Services.PrimaryService)
        MSN.List(MLIST_ALLOW).Remove User
        MSN.List(MLIST_BLOCK).Add User
    Else
    
    End If
End Sub

Private Sub cmdNormalBlockAll_Click()
On Error Resume Next
    Dim a As Integer
    
    Do Until a = lstNotBlocked.ListCount
        Set User = MSN.CreateUser(lstNotBlocked.List(a), MSN.Services.PrimaryService)
        MSN.List(MLIST_ALLOW).Remove User
        MSN.List(MLIST_BLOCK).Add User
        a = a + 1
    Loop
End Sub

Public Sub LoadTheFormForReal()
    Set MSN = New MsgrObject
    
    SendMessage = True
    
    t1.ImageList = ilsIcons
    t2.ImageList = ilsIcons
    t3.ImageList = ilsIcons
    
    If MSN.LocalState <> MSTATE_OFFLINE Then
        lblStatus.Caption = "Current User : " & MSN.LocalLogonName & vbCrLf & "Current NickName : " & MSN.LocalFriendlyName & vbCrLf & "Current Status : " & GetState
    Else
        lblStatus.Caption = "Currently not connected to the .NET Messenger Service"
    End If

    RefreshLists
    
    LoadCrashHistory
    
    If ckAutoMessage.Value = 1 Then
        If optAway.Value = True Then
            MSN.LocalState = MSTATE_AWAY
        Else
            MSN.LocalState = MSTATE_BUSY
        End If
    Else
    
    End If
    
    If ckLog.Value = 1 Then
        If ckSaveLog.Value = 1 Then
            LogActions True
            Logstrm.WriteLine "########## " & Date & " Started at : " & Time & " ##########"
        Else
        
        End If
    Else
    
    End If
End Sub

Private Sub LoadCrashHistory()
On Error GoTo a

    lstCrashHistory.ListItems.Clear

    Set strm = fso.OpenTextFile(App.Path & "\Crash History.BartNet", ForReading)
    With strm
        Do Until .AtEndOfStream
            Set lItem = lstCrashHistory.ListItems.Add(, , .ReadLine)
            lItem.ListSubItems.Add , , .ReadLine
            lItem.ListSubItems.Add , , .ReadLine
        Loop
        
        .Close
    End With
    
a:

End Sub

Private Sub RefreshLists()
    lstNotBlocked.Clear
    lstNormalBlocked.Clear

    For Each User In MSN.List(MLIST_ALLOW)
        lstNotBlocked.AddItem (User.EmailAddress)
    Next
    
    For Each User In MSN.List(MLIST_BLOCK)
        lstNormalBlocked.AddItem (User.EmailAddress)
    Next

    Dim UsersOnline As Integer
    Dim UsersOffline As Integer
    
    UsersOnline = 0
    UsersOffline = 0
    
    t1.Nodes.Clear
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State = MSTATE_OFFLINE Then
            UsersOffline = UsersOffline + 1
        Else
            UsersOnline = UsersOnline + 1
        End If
    Next
    
    t1.Nodes.Add , , "Online", "Online (" & UsersOnline & ")", "Up", "UpSelected"
    With t1.Nodes(1)
        .Selected = True
        .Expanded = True
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    t1.Nodes.Add , , "Offline", "Offline (" & UsersOffline & ")", "Down", "DownSelected"
    With t1.Nodes(2)
        .Expanded = False
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    For Each User In MSN.List(MLIST_ALLOW)
        Select Case User.State
            Case MSTATE_AWAY
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Away)", "Away", "AwaySelected"
            Case MSTATE_BE_RIGHT_BACK
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Be Right Back)", "Away", "AwaySelected"
            Case MSTATE_BUSY
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Busy)", "Busy", "BusySelected"
            Case MSTATE_OFFLINE

            Case MSTATE_ON_THE_PHONE
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (On The Phone)", "Busy", "BusySelected"
            Case MSTATE_ONLINE
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName, "Online", "OnlineSelected"
            Case MSTATE_OUT_TO_LUNCH
                t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Out To Lunch)", "Away", "AwaySelected"
        End Select
    Next
    
    For Each User In MSN.List(MLIST_BLOCK)
        If User.State <> MSTATE_OFFLINE Then
            t1.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Blocked)", "Blocked", "BlockedSelected"
        Else

        End If
    Next
      
    UsersOnline = 0
    UsersOffline = 0
    
    t2.Nodes.Clear
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State = MSTATE_OFFLINE Then
            UsersOffline = UsersOffline + 1
        Else
            UsersOnline = UsersOnline + 1
        End If
    Next
    
    t2.Nodes.Add , , "Online", "Online (" & UsersOnline & ")", "Up", "UpSelected"
    With t2.Nodes(1)
        .Selected = True
        .Expanded = True
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    t2.Nodes.Add , , "Offline", "Offline (" & UsersOffline & ")", "Down", "DownSelected"
    With t2.Nodes(2)
        .Expanded = False
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    For Each User In MSN.List(MLIST_ALLOW)
        Select Case User.State
            Case MSTATE_AWAY
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Away)", "Away", "AwaySelected"
            Case MSTATE_BE_RIGHT_BACK
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Be Right Back)", "Away", "AwaySelected"
            Case MSTATE_BUSY
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Busy)", "Busy", "BusySelected"
            Case MSTATE_OFFLINE

            Case MSTATE_ON_THE_PHONE
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (On The Phone)", "Busy", "BusySelected"
            Case MSTATE_ONLINE
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName, "Online", "OnlineSelected"
            Case MSTATE_OUT_TO_LUNCH
                t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Out To Lunch)", "Away", "AwaySelected"
        End Select
    Next
    
    For Each User In MSN.List(MLIST_BLOCK)
        If User.State <> MSTATE_OFFLINE Then
            t2.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Blocked)", "Blocked", "BlockedSelected"
        Else

        End If
    Next
    
    
    UsersOnline = 0
    UsersOffline = 0
    
    t3.Nodes.Clear
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State = MSTATE_OFFLINE Then
            UsersOffline = UsersOffline + 1
        Else
            UsersOnline = UsersOnline + 1
        End If
    Next
    
    t3.Nodes.Add , , "Online", "Online (" & UsersOnline & ")", "Up", "UpSelected"
    With t3.Nodes(1)
        .Selected = True
        .Expanded = True
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    t3.Nodes.Add , , "Offline", "Offline (" & UsersOffline & ")", "Down", "DownSelected"
    With t3.Nodes(2)
        .Expanded = False
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    For Each User In MSN.List(MLIST_ALLOW)
        Select Case User.State
            Case MSTATE_AWAY
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Away)", "Away", "AwaySelected"
            Case MSTATE_BE_RIGHT_BACK
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Be Right Back)", "Away", "AwaySelected"
            Case MSTATE_BUSY
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Busy)", "Busy", "BusySelected"
            Case MSTATE_OFFLINE

            Case MSTATE_ON_THE_PHONE
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (On The Phone)", "Busy", "BusySelected"
            Case MSTATE_ONLINE
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName, "Online", "OnlineSelected"
            Case MSTATE_OUT_TO_LUNCH
                t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Out To Lunch)", "Away", "AwaySelected"
        End Select
    Next
    
    For Each User In MSN.List(MLIST_BLOCK)
        If User.State <> MSTATE_OFFLINE Then
            t3.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Blocked)", "Blocked", "BlockedSelected"
        Else

        End If
    Next
    
    If lstNormalBlocked.ListCount = 0 Then
        cmdNormalAllow.Enabled = False
        cmdNormalAllowAll.Enabled = False
    Else
        cmdNormalAllow.Enabled = True
        cmdNormalAllowAll.Enabled = True
    End If
    
    If lstBartNetBlocked.ListCount = 0 Then
        cmdBartNetAllow.Enabled = False
        cmdBartNetAllowAll.Enabled = False
    Else
        cmdBartNetAllow.Enabled = True
        cmdBartNetAllowAll.Enabled = True
    End If
    
    If lstNotBlocked.ListCount = 0 Then
        cmdNormalBlock.Enabled = False
        cmdBartNetBlock.Enabled = False
        cmdNormalBlockAll.Enabled = False
        cmdBartNetBlockAll.Enabled = False
    Else
        cmdNormalBlock.Enabled = True
        cmdBartNetBlock.Enabled = True
        cmdNormalBlockAll.Enabled = True
        cmdBartNetBlockAll.Enabled = True
    End If
End Sub

Public Function GetState()
    Select Case MSN.LocalState
        Case MSTATE_AWAY
            GetState = "Away"
        Case MSTATE_BE_RIGHT_BACK
            GetState = "Be Right Back"
        Case MSTATE_BUSY
            GetState = "Busy"
        Case MSTATE_INVISIBLE
            GetState = "Appear Offline"
        Case MSTATE_ON_THE_PHONE
            GetState = "On The Phone"
        Case MSTATE_ONLINE
            GetState = "Online"
        Case MSTATE_OUT_TO_LUNCH
            GetState = "Out To Lunch"
    End Select
End Function

Public Function GetUserState(ByVal WhichUser As IMsgrUser)
    Select Case WhichUser.State
        Case MSTATE_AWAY
            GetUserState = "Away"
        Case MSTATE_BE_RIGHT_BACK
            GetUserState = "Be Right Back"
        Case MSTATE_BUSY
            GetUserState = "Busy"
        Case MSTATE_INVISIBLE
            GetUserState = "Appear Offline"
        Case MSTATE_ON_THE_PHONE
            GetUserState = "On The Phone"
        Case MSTATE_ONLINE
            GetUserState = "Online"
        Case MSTATE_OUT_TO_LUNCH
            GetUserState = "Out To Lunch"
    End Select
End Function

Private Sub cmdPopup_Click()
    Dim ScrollerOnOff As Boolean
    
    cmdShowPeopleYouAreOnline.Enabled = False
    cmdPopup.Enabled = False
    If cmdScroll.Caption = "Start" Then
        ScrollerOnOff = False
    Else
        cmdScroll_Click
        ScrollerOnOff = True
    End If

    PopupTotalLines = 0
    PrevPopupNick = MSN.LocalFriendlyName
    PrevPopupState = MSN.LocalState
    
    Set strmPopup = fsoPopup.CreateTextFile(fsoPopup.GetSpecialFolder(TemporaryFolder) & "\Temp.BartNet", True)
    With strmPopup
        .Write txtPopup.Text
        .Close
    End With
    
    Set strmPopup = fsoPopup.OpenTextFile(fsoPopup.GetSpecialFolder(TemporaryFolder) & "\Temp.BartNet", ForReading)
    With strmPopup
        Do Until .AtEndOfStream
            .ReadLine
            PopupTotalLines = PopupTotalLines + 1
        Loop
        .Close
    End With
    
    Dim a As Integer
    a = 0
    
    Set strmPopup = fsoPopup.OpenTextFile(fsoPopup.GetSpecialFolder(TemporaryFolder) & "\Temp.BartNet", ForReading)
    With strmPopup
        Do Until a = PopupTotalLines
            PopupMessage(a) = .ReadLine
            a = a + 1
        Loop
        .Close
    End With
    
    Do Until PopupTotalLines = 0
        MSN.LocalState = MSTATE_INVISIBLE
        MSN.Services.PrimaryService.FriendlyName = PopupMessage(PopupTotalLines - 1)
        MSN.LocalState = MSTATE_ONLINE
        
        PopupTotalLines = PopupTotalLines - 1
    Loop
    
    MSN.LocalState = PrevPopupState
    MSN.Services.PrimaryService.FriendlyName = PrevPopupNick
    
    cmdShowPeopleYouAreOnline.Enabled = True
    cmdPopup.Enabled = True
    If ScrollerOnOff = True Then
        cmdScroll_Click
    Else
    
    End If
End Sub

Private Sub cmdScroll_Click()
    If cmdScroll.Caption = "Start" Then
        PrevScrollerNick = MSN.LocalFriendlyName
    
        cmdScroll.Caption = "Stop"
        txtScroll(0).Enabled = False
        txtScroll(1).Enabled = False
        txtScroll(2).Enabled = False
        txtScroll(3).Enabled = False
        txtScroll(4).Enabled = False
        txtScroll(5).Enabled = False
        ckNickTime.Enabled = False
        lblNickTime.ForeColor = Grey
        ckResetNick.Enabled = False
        lblResetNick.ForeColor = Grey
        Label1.ForeColor = Grey
        Label2.ForeColor = Grey
        txtTime.Enabled = False
        
        Dim a As Integer
        
        If ckNickTime.Value = 1 Then
            a = 1000
            MSN.Services.PrimaryService.FriendlyName = Mid(Time, 1, 5)
        Else
            a = txtTime.Text
        End If
        
        timScroller.Interval = a
        timScroller.Enabled = True
    Else
        cmdScroll.Caption = "Start"
        
        If ckNickTime.Value = 1 Then
            ckNickTime.Enabled = True
            ckResetNick.Enabled = True
            lblNickTime.ForeColor = Black
            lblResetNick.ForeColor = Black
            Label1.ForeColor = Black
            Label2.ForeColor = Black
        Else
            txtScroll(0).Enabled = True
            txtScroll(1).Enabled = True
            txtScroll(2).Enabled = True
            txtScroll(3).Enabled = True
            txtScroll(4).Enabled = True
            txtScroll(5).Enabled = True
            ckNickTime.Enabled = True
            lblNickTime.ForeColor = Black
            ckResetNick.Enabled = True
            lblResetNick.ForeColor = Black
            Label1.ForeColor = Black
            Label2.ForeColor = Black
            txtTime.Enabled = True
        End If
        
        If ckResetNick.Value = 1 Then
            MSN.Services.PrimaryService.FriendlyName = PrevScrollerNick
        Else
        
        End If
        
        timScroller.Enabled = False
    End If
End Sub


Private Sub cmdSendIM_Click()
    If t3.Nodes.Item(1).Selected = True Or t3.Nodes.Item(2).Selected = True Then
   
    Else
        MSNAPI.InstantMessage (t3.SelectedItem.Key)
    End If
End Sub

Private Sub cmdShowPeopleYouAreOnline_Click()
On Error Resume Next
    Dim ScrollerOnOff As Boolean

    cmdShowPeopleYouAreOnline.Enabled = False
    cmdPopup.Enabled = False
    If cmdScroll.Caption = "Start" Then
        ScrollerOnOff = False
    Else
        cmdScroll_Click
        ScrollerOnOff = True
    End If

    Dim a As Integer
    Dim PrevState As Messenger.MSTATE
    
    PrevState = MSN.LocalState
    
    Do Until a = 8
        MSN.LocalState = MSTATE_INVISIBLE
        MSN.LocalState = MSTATE_ONLINE
        a = a + 1
    Loop
    
    MSN.LocalState = PrevState
    
    cmdShowPeopleYouAreOnline.Enabled = True
    cmdPopup.Enabled = True
    If ScrollerOnOff = True Then
        cmdScroll_Click
    Else
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    SaveValues
    
    Logstrm.Close
    
    If optAppearOffline.Value = True Then
        optAppearOnline_MouseUp 1, 1, 1, 1
    Else
    
    End If
End Sub

Private Sub imgClose_Click()
    Unload Me
    End
End Sub

Private Sub imgMinimize_Click()
    Set imgMinimize.Picture = LoadPicture(App.Path & "\Minimize 1.bmp")
    Me.WindowState = vbMinimized
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set imgClose.Picture = LoadPicture(App.Path & "\Close 3.bmp")
End Sub

Private Sub imgMinimize_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set imgMinimize.Picture = LoadPicture(App.Path & "\Minimize 3.bmp")
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set imgClose.Picture = LoadPicture(App.Path & "\Close 2.bmp")
End Sub

Private Sub imgMinimize_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set imgMinimize.Picture = LoadPicture(App.Path & "\Minimize 2.bmp")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set imgClose.Picture = LoadPicture(App.Path & "\Close 1.bmp")
    Set imgMinimize.Picture = LoadPicture(App.Path & "\Minimize 1.bmp")
    
    If lblAutoMessage.ForeColor = Blue Then
        lblAutoMessage.ForeColor = Black
    Else
    
    End If
    
    If lblBlocker.ForeColor = Blue Then
        lblBlocker.ForeColor = Black
    Else
    
    End If
    
    If lblCrasher.ForeColor = Blue Then
        lblCrasher.ForeColor = Black
    Else
    
    End If
    
    If lblHelp.ForeColor = Blue Then
        lblHelp.ForeColor = Black
    Else
    
    End If
    
    If lblLogger.ForeColor = Blue Then
        lblLogger.ForeColor = Black
    Else
    
    End If
    
    If lblMultiMessageSender.ForeColor = Blue Then
        lblMultiMessageSender.ForeColor = Black
    Else
    
    End If
    
    If lblNickNamePopups.ForeColor = Blue Then
        lblNickNamePopups.ForeColor = Black
    Else
    
    End If
    
    If lblNickNameScroller.ForeColor = Blue Then
        lblNickNameScroller.ForeColor = Black
    Else
    
    End If
    
    If lblSendIM.ForeColor = Blue Then
        lblSendIM.ForeColor = Black
    Else
    
    End If
    
    If lblTalkOffline.ForeColor = Blue Then
        lblTalkOffline.ForeColor = Black
    Else
    
    End If
    
    If lblWelcomingMessage.ForeColor = Blue Then
        lblWelcomingMessage.ForeColor = Black
    Else
    
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    formdrag Me
End Sub

Private Sub lblAutoMessage_Click()
    If lblAutoMessage.ForeColor = Blue Then
        lblAutoMessage.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Show", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblAutoMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
    End If
End Sub

Private Sub lblAutoMessage_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblAutoMessage.ForeColor <> Grey Then
        lblAutoMessage.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblBlocker_Click()
    If lblBlocker.ForeColor = Blue Then
        lblBlocker.ForeColor = Grey
        
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Show", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblBlocker.ForeColor = Black
        
        SetView "Hide", "Blocker"
    End If
End Sub
Private Sub lblBlocker_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblBlocker.ForeColor <> Grey Then
        lblBlocker.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblCrasher_Click()
    If lblCrasher.ForeColor = Blue Then
        lblCrasher.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Show", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblCrasher.ForeColor = Black
        
        SetView "Hide", "Crasher"
    End If
End Sub

Private Sub lblCrasher_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblCrasher.ForeColor <> Grey Then
        lblCrasher.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblHelp_Click()
    If lblHelp.ForeColor = Blue Then
        lblHelp.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Show", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblHelp.ForeColor = Black
        
        SetView "Hide", "Help"
    End If
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblHelp.ForeColor <> Grey Then
        lblHelp.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblLogger_Click()
    If lblLogger.ForeColor = Blue Then
        lblLogger.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Show", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblLogger.ForeColor = Black
        
        SetView "Hide", "Logger"
    End If
End Sub

Private Sub lblLogger_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblLogger.ForeColor <> Grey Then
        lblLogger.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblMultiMessageSender_Click()
    If lblMultiMessageSender.ForeColor = Blue Then
        lblMultiMessageSender.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Show", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblMultiMessageSender.ForeColor = Black
        
        SetView "Hide", "MultiMessageSender"
    End If
End Sub

Private Sub lblMultiMessageSender_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblMultiMessageSender.ForeColor <> Grey Then
        lblMultiMessageSender.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblNickNamePopups_Click()
    If lblNickNamePopups.ForeColor = Blue Then
        lblNickNamePopups.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Show", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblNickNamePopups.ForeColor = Black
        
        SetView "Hide", "NickNamePopups"
    End If
End Sub

Private Sub lblNickNamePopups_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblNickNamePopups.ForeColor <> Grey Then
        lblNickNamePopups.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblNickNameScroller_Click()
    If lblNickNameScroller.ForeColor = Blue Then
        lblNickNameScroller.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Show", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblNickNameScroller.ForeColor = Black
        
        SetView "Hide", "NickNameScroller"
    End If
End Sub

Private Sub lblNickNameScroller_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblNickNameScroller.ForeColor <> Grey Then
        lblNickNameScroller.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblSendIM_Click()
    If lblSendIM.ForeColor = Blue Then
        lblSendIM.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Show", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblSendIM.ForeColor = Black
        
        SetView "Hide", "SendIM"
    End If
End Sub

Private Sub lblSendIM_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblSendIM.ForeColor <> Grey Then
        lblSendIM.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblTalkOffline_Click()
    If lblTalkOffline.ForeColor = Blue Then
        lblTalkOffline.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Show", "TalkOffline"
        SetView "Hide", "WelcomingMessage"
    Else
        lblTalkOffline.ForeColor = Black
        
        SetView "Hide", "TalkOffline"
    End If
End Sub

Private Sub lblTalkOffline_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblTalkOffline.ForeColor <> Grey Then
        lblTalkOffline.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub lblWelcomingMessage_Click()
    If lblWelcomingMessage.ForeColor = Blue Then
        lblWelcomingMessage.ForeColor = Grey
        
        lblBlocker.ForeColor = Black
        lblAutoMessage.ForeColor = Black
        lblCrasher.ForeColor = Black
        lblHelp.ForeColor = Black
        lblLogger.ForeColor = Black
        lblMultiMessageSender.ForeColor = Black
        lblNickNamePopups.ForeColor = Black
        lblNickNameScroller.ForeColor = Black
        lblSendIM.ForeColor = Black
        lblTalkOffline.ForeColor = Black
        
        SetView "Hide", "AutoMessage"
        SetView "Hide", "Blocker"
        SetView "Hide", "Crasher"
        SetView "Hide", "Help"
        SetView "Hide", "Logger"
        SetView "Hide", "MultiMessageSender"
        SetView "Hide", "NickNamePopups"
        SetView "Hide", "NickNameScroller"
        SetView "Hide", "SendIM"
        SetView "Hide", "TalkOffline"
        SetView "Show", "WelcomingMessage"
    Else
        lblWelcomingMessage.ForeColor = Black
        
        SetView "Hide", "WelcomingMessage"
    End If
End Sub

Private Sub lblWelcomingMessage_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If lblWelcomingMessage.ForeColor <> Grey Then
        lblWelcomingMessage.ForeColor = Blue
    Else
    
    End If
End Sub


Private Sub MSN_OnFileTransferInviteAccepted(ByVal pUser As Messenger.IMsgrUser, ByVal lCookie As Long, pfEnableDefault As Boolean)
    Log "[" & Time & "] - You have accepted the file transfer from " & pUser.EmailAddress
End Sub

Private Sub MSN_OnFileTransferInviteCancelled(ByVal pUser As Messenger.IMsgrUser, ByVal lCookie As Long, ByVal hrReason As Long, pfEnableDefault As Boolean)
    Log "[" & Time & "] - You have not accepted the file transfer from " & pUser.EmailAddress
End Sub

Private Sub MSN_OnFileTransferInviteReceived(ByVal pUser As Messenger.IMsgrUser, ByVal lCookie As Long, ByVal bstrFileName As String, ByVal lFileSize As Long, pfEnableDefault As Boolean)
    Log "[" & Time & "] - " & pUser.EmailAddress & " has send you a file transfer for '" & bstrFileName & "' (" & lFileSize & " KB)"
End Sub

Private Sub MSN_OnListAddResult(ByVal hr As Long, ByVal MLIST As Messenger.MLIST, ByVal pUser As Messenger.IMsgrUser)
    RefreshLists
    
    If MLIST = MLIST_BLOCK Then
        Log "[" & Time & "] - You have blocked " & pUser.EmailAddress
    Else
        If MLIST = MLIST_ALLOW Then
            Log "[" & Time & "] - You have unblocked " & pUser.EmailAddress
        Else
        
        End If
    End If
End Sub

Private Sub MSN_OnLocalFriendlyNameChangeResult(ByVal hr As Long, ByVal pService As Messenger.IMsgrService, ByVal bstrPrevFriendlyName As String)
    lblStatus.Caption = "Current User : " & MSN.LocalLogonName & vbCrLf & "Current NickName : " & MSN.LocalFriendlyName & vbCrLf & "Current Status : " & GetState
    
    Log "[" & Time & "] - You have changed your nickname from '" & bstrPrevFriendlyName & "' to '" & MSN.LocalFriendlyName & "'"
End Sub

Private Sub MSN_OnLocalStateChangeResult(ByVal hr As Long, ByVal mLocalState As Messenger.MSTATE, ByVal pService As Messenger.IMsgrService)
    lblStatus.Caption = "Current User : " & MSN.LocalLogonName & vbCrLf & "Current NickName : " & MSN.LocalFriendlyName & vbCrLf & "Current Status : " & GetState
    
    Log "[" & Time & "] - You have changed your state to " & GetState
End Sub

Private Sub MSN_OnLogoff()
    lblStatus.Caption = "Currently not connected to the .NET Messenger Service"
    Log "[" & Time & "] - You have logged out of the .NET Messenger Service"
End Sub

Private Sub MSN_OnLogonResult(ByVal hr As Long, ByVal pService As Messenger.IMsgrService)
    lblStatus.Caption = "Current User : " & MSN.LocalLogonName & vbCrLf & "Current NickName : " & MSN.LocalFriendlyName & vbCrLf & "Current Status : " & GetState
    Log "[" & Time & "] - You have logged out of the .NET Messenger Service with " & MSN.LocalLogonName
End Sub

Private Sub MSN_OnTextReceived(ByVal pIMSession As Messenger.IMsgrIMSession, ByVal pSourceUser As Messenger.IMsgrUser, ByVal bstrMsgHeader As String, ByVal bstrMsgText As String, pfEnableDefault As Boolean)
    Dim MessageHeader As String
    
    'AutoMessage
    If SendMessage = True Then
        If ckAutoMessage.Value = 1 Then
            If ckIncludeTime.Value = 1 Then
                If optCountDown.Value = True Then
                    pSourceUser.SendText MessageHeader, txtAutoMessage.Text & vbCrLf & vbCrLf & "Expected remaining time until return : " & txtCountDown.Text, MMSGTYPE_ALL_RESULTS
                Else
                    pSourceUser.SendText MessageHeader, txtAutoMessage.Text & vbCrLf & vbCrLf & "Expected time of return : " & txtNormal.Text, MMSGTYPE_ALL_RESULTS
                End If
            Else
                pSourceUser.SendText MessageHeader, txtAutoMessage.Text, MMSGTYPE_ALL_RESULTS
            End If
        Else
        
        End If
        
        SendMessage = False
        timMessageCheck.Enabled = True
    Else
    
    End If
    
    'Blocker
    Dim a As Integer
    
    Do Until a = lstBartNetBlocked.ListCount
        If LCase(pSourceUser.EmailAddress) = LCase(lstBartNetBlocked.List(a)) Then
            pfEnableDefault = False
        Else
            pfEnableDefault = True
        End If
        a = a + 1
    Loop
End Sub

Private Sub MSN_OnUnreadEmailChanged(ByVal MFOLDER As Messenger.MFOLDER, ByVal cUnreadEmail As Long, pfEnableDefault As Boolean)
    Log "[" & Time & "] - You have recieved an email message"
End Sub

Private Sub MSN_OnUserFriendlyNameChangeResult(ByVal hr As Long, ByVal pUser As Messenger.IMsgrUser, ByVal bstrPrevFriendlyName As String)
    If LCase(pUser.EmailAddress) = LCase(MSN.LocalLogonName) Then
    
    Else
        Log "[" & Time & "] - " & pUser.EmailAddress & " has changed his / her nickname from '" & bstrPrevFriendlyName & "' to '" & pUser.FriendlyName & "'"
    End If
End Sub

Private Sub MSN_OnUserStateChanged(ByVal pUser As Messenger.IMsgrUser, ByVal mPrevState As Messenger.MSTATE, pfEnableDefault As Boolean)
    RefreshLists
    
    'Crasher
    If pUser.State = MSTATE_OFFLINE Then
        If LCase(pUser.EmailAddress) = LCase(CurrentlyCrashing) Then
            timCrash.Enabled = False
            cmdCrash.Caption = "Crash"
            lblCurrentlyCrashing.Caption = "Currently Not Crashing Anybody"
            Set strm = fso.OpenTextFile(App.Path & "\Crash History.BartNet", ForAppending)
            With strm
                .WriteLine CurrentlyCrashing
                .WriteLine Date
                .WriteLine Time
                
                .Close
            End With
            
            LoadCrashHistory
            
            CurrentlyCrashing = ""
        Else
        
        End If
    Else
    
    End If
    
    If mPrevState = MSTATE_OFFLINE Then
        Log "[" & Time & "] - " & pUser.EmailAddress & " has signed in to the .NET Messenger Service"
        If ckWelcomingMessage.Value = 1 Then
            pUser.SendText bstrMessageHeader, txtWelcomingMessage.Text, MMSGTYPE_ALL_RESULTS
        Else
        
        End If
        Exit Sub
    Else
    
    End If
    
    If pUser.State = MSTATE_OFFLINE Then
        Log "[" & Time & "] - " & pUser.EmailAddress & " has signed out of the .NET Messenger Service"
        Exit Sub
    Else
    
    End If
    
    Log "[" & Time & "] - " & pUser.EmailAddress & " has changed his / her status to " & GetUserState(pUser)
End Sub

Private Sub optAppearOffline_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    Dim usr As IMsgrUser
    
    For Each usr In MSN.List(MLIST_ALLOW)
        MSN.List(MLIST_ALLOW).Remove usr
    Next
    
    optAppearOffline.Value = True
    optAppearOnline.Value = False
End Sub

Private Sub optAppearOnline_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    Dim usr As IMsgrUser
    
    For Each usr In MSN.List(MLIST_CONTACT)
        MSN.List(MLIST_ALLOW).Add usr
    Next
    
    optAppearOffline.Value = False
    optAppearOnline.Value = True
End Sub

Private Sub optAway_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MSN.LocalState = MSTATE_AWAY
    optAway.Value = True
    optBusy.Value = False
End Sub

Private Sub optBusy_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MSN.LocalState = MSTATE_BUSY
    optAway.Value = False
    optBusy.Value = True
End Sub

Private Sub optCountDown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCountDown.ForeColor = Black
    txtCountDown.Enabled = True
    lblNormal.ForeColor = Grey
    txtNormal.Enabled = False
    optCountDown.Value = True
    optNormal.Value = False
    timAutoMessage.Enabled = True
End Sub

Private Sub optNormal_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCountDown.ForeColor = Grey
    txtCountDown.Enabled = False
    lblNormal.ForeColor = Black
    txtNormal.Enabled = True
    optCountDown.Value = False
    optNormal.Value = True
    timAutoMessage.Enabled = False
End Sub

Private Sub t1_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Online" Then
        If Node.Expanded = True Then
            Node.Expanded = False
            Node.Image = "Down"
            Node.SelectedImage = "DownSelected"
        Else
            Node.Expanded = True
            Node.Image = "Up"
            Node.SelectedImage = "UpSelected"
        End If
    Else
    
    End If
End Sub

Private Sub t2_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Online" Then
        If Node.Expanded = True Then
            Node.Expanded = False
            Node.Image = "Down"
            Node.SelectedImage = "DownSelected"
        Else
            Node.Expanded = True
            Node.Image = "Up"
            Node.SelectedImage = "UpSelected"
        End If
    Else
    
    End If
End Sub

Private Sub t3_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Online" Then
        If Node.Expanded = True Then
            Node.Expanded = False
            Node.Image = "Down"
            Node.SelectedImage = "DownSelected"
        Else
            Node.Expanded = True
            Node.Image = "Up"
            Node.SelectedImage = "UpSelected"
        End If
    Else
    
    End If
End Sub

Private Sub timAutoMessage_Timer()
    Dim a As Integer
    Dim B As Integer
    
    a = Mid(txtCountDown.Text, 1, 2)
    B = Mid(txtCountDown.Text, 4, 2)
    
    If B = 0 Then
        If a = 0 Then
        
        Else
            a = a - 1
            B = 59
        End If
    Else
        B = B - 1
    End If
    
    txtCountDown.Text = a & ":" & B
End Sub

Private Sub timCrash_Timer()
On Error Resume Next
    CrashingSession.SendText bstrMessageHeader, CrashMessage, MMSGTYPE_ERRORS_ONLY
End Sub

Private Sub timMessageCheck_Timer()
    SendMessage = True
    timMessageCheck.Enabled = False
End Sub

Private Sub timScroller_Timer()
    If ckNickTime.Value = 1 Then
        If MSN.LocalFriendlyName = Mid(Time, 1, 5) Then
        
        Else
            MSN.Services.PrimaryService.FriendlyName = Mid(Time, 1, 5)
        End If
    Else
        Select Case MSN.LocalFriendlyName
            Case txtScroll(0).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(1).Text
                Exit Sub
            Case txtScroll(1).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(2).Text
                Exit Sub
            Case txtScroll(2).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(3).Text
                Exit Sub
            Case txtScroll(3).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(4).Text
                Exit Sub
            Case txtScroll(4).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(5).Text
                Exit Sub
            Case txtScroll(5).Text
                MSN.Services.PrimaryService.FriendlyName = txtScroll(0).Text
                Exit Sub
        End Select
        
        MSN.Services.PrimaryService.FriendlyName = txtScroll(0).Text
    End If
End Sub

Private Sub txtCountDown_GotFocus()
    timAutoMessage.Enabled = False
End Sub

Private Sub txtCountDown_LostFocus()
    timAutoMessage.Enabled = True
End Sub

