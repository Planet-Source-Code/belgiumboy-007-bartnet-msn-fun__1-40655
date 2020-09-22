VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XpCheckBox"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Mixed"
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   1630
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Disable"
      Height          =   255
      Left            =   5280
      TabIndex        =   24
      Top             =   1000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enable"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   1000
      Width           =   640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Uncheck"
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   1630
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   1630
      Width           =   615
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   0
      Left            =   4680
      TabIndex        =   20
      Top             =   1320
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Check1"
      Height          =   195
      Left            =   1680
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   1
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   2
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   7
      Top             =   2760
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   3
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   8
      Top             =   3000
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   4
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   6
      Left            =   1800
      TabIndex        =   9
      Top             =   3240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   5
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   7
      Left            =   2040
      TabIndex        =   10
      Top             =   3480
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   6
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   8
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   7
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   9
      Left            =   5520
      TabIndex        =   12
      Top             =   3720
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   8
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   10
      Left            =   5280
      TabIndex        =   13
      Top             =   3480
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   9
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   11
      Left            =   5040
      TabIndex        =   14
      Top             =   3240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   10
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   12
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   11
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   13
      Left            =   4560
      TabIndex        =   16
      Top             =   2760
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   12
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   14
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   13
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   15
      Left            =   4080
      TabIndex        =   18
      Top             =   2280
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      CheckBoxLook    =   6
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.XpCheckBox XpCheckBox1 
      Height          =   195
      Index           =   16
      Left            =   3840
      TabIndex        =   19
      Top             =   2040
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   344
      Caption         =   "CheckBox1"
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You can use ""Tab"" to jump from one to another.            To check/uncheck use ""Space""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   4080
      Width           =   6735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   480
      X2              =   6720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   2
      X1              =   6720
      X2              =   6720
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   480
      X2              =   480
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   3720
      X2              =   3720
      Y1              =   960
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   480
      X2              =   6720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Can you see the difference ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
XpCheckBox1(0).Value = Checked
End Sub

Private Sub Command2_Click()
XpCheckBox1(0).Value = Unchecked
End Sub

Private Sub Command3_Click()
XpCheckBox1(0).Enabled = True
End Sub

Private Sub Command4_Click()
XpCheckBox1(0).Enabled = False
End Sub

Private Sub Command5_Click()
XpCheckBox1(0).Value = Mixed
End Sub
