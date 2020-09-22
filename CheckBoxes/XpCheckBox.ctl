VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl XpCheckBox 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   FillStyle       =   0  'Solid
   ScaleHeight     =   7665
   ScaleWidth      =   9120
   ToolboxBitmap   =   "XpCheckBox.ctx":0000
   Begin PicClip.PictureClip pc 
      Left            =   0
      Top             =   480
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":0312
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   11
      Left            =   0
      Top             =   3360
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":1B28
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   10
      Left            =   0
      Top             =   3120
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":333E
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   2
      Left            =   0
      Top             =   1200
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":4B54
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   3
      Left            =   0
      Top             =   1440
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":636A
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   7
      Left            =   0
      Top             =   2400
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":7B80
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   0
      Left            =   0
      Top             =   720
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":9396
   End
   Begin VB.PictureBox p 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   0
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   6
      Left            =   0
      Top             =   2160
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":ABAC
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   1
      Left            =   0
      Top             =   960
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":C3C2
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   4
      Left            =   0
      Top             =   1680
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":DBD8
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   8
      Left            =   0
      Top             =   2640
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":F3EE
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   12
      Left            =   0
      Top             =   3600
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":10C04
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   5
      Left            =   0
      Top             =   1920
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":1241A
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   13
      Left            =   0
      Top             =   3840
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":13C30
   End
   Begin PicClip.PictureClip pcChoice 
      Index           =   9
      Left            =   0
      Top             =   2880
      _ExtentX        =   4128
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   12
      Picture         =   "XpCheckBox.ctx":15446
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   375
      TabIndex        =   1
      Top             =   0
      Width           =   585
   End
End
Attribute VB_Name = "XpCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINT_API) As Long                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          'Aki
Private Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'  *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
'*   *  This code is originaly writen by Aleks Kimhi - Aki.         *   *   *                             *
'  *    As you might have read that this ActiveX control is free,   *   *         O                       *
'*   *  and therefore can be distributed for free as long as you      *                                       *
'  *    do not sell it for profit and there's still my name on it.         *                     ____|            *
'*   *  I've spend some time on this project so please                *  *                                        *
'  *    don't just take it, use it and say you wrote it.                 *  *   *                                     *
'*   *  Thank you for your co-operation. Any comments,        *   *  *   *      P                          *
'  *    good or bad, would be greatly appreciated.                       *                                      *
'*   *  E -mail: aniram@ zahav.net.il                                       *    *                                 *
'  *    Tel-Aviv, Israel    A & M © Copyright 2002                   *    *     *                          *
'*   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
' *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *   *
Private Type POINT_API
    x As Long
    Y As Long
End Type

Public Enum State
    Unchecked = 0
    Checked = 1
    Mixed = 2
End Enum

Public Enum Pict
    XP_Default = 0
    XP_AccentedEdges = 1
    XP_BlackWhite = 2
    XP_Blue = 3
    XP_Disco = 4
    XP_Green = 5
    XP_HighPass = 6
    XP_Lily = 7
    XP_MidlleAges = 8
    XP_Orange = 9
    XP_Red = 10
    XP_Solarize = 11
    XP_Spectrum = 12
    XP_Yellow = 13
End Enum

Dim mPic As Pict
Const defPic = Pict.XP_Default

Dim mFont As Font
Dim mValue As State
Dim mBackColor As OLE_COLOR
Dim mForeColor As OLE_COLOR

Const defValue = State.Unchecked
Const defBackColor = vbButtonFace
Const defForeColor = vbBlack

Dim chVal, btnDown As Integer

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
        btnDown = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Enabled = False Then Exit Sub
        If mValue = Checked Then
            p.Picture = pc.GraphicCell(6)
                ElseIf mValue = Mixed Then
                    p.Picture = pc.GraphicCell(10)
                        ElseIf mValue = Unchecked Then
                    p.Picture = pc.GraphicCell(2)
                End If
            btnDown = 1
        RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Enabled = False Then Exit Sub
        If p.Picture = pc.GraphicCell(chVal) Then Exit Sub 'No reason to came in all the time
           If btnDown = 1 Then Exit Sub
            Timer1.Enabled = True
                If mValue = Checked Then
                    p.Picture = pc.GraphicCell(5)
                        chVal = 5
                            ElseIf mValue = Mixed Then
                                p.Picture = pc.GraphicCell(9)
                                    chVal = 9
                                ElseIf mValue = Unchecked Then
                            p.Picture = pc.GraphicCell(1)
                        chVal = 1
                End If
        RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub p_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub p_KeyPress(KeyAscii As Integer) 'Like Sub MouseDown(just using KeyPress)
    If KeyAscii <> vbKeySpace Then Exit Sub 'only "space" can come in
           RaiseEvent KeyPress(KeyAscii)
              RaiseEvent Click
                   Call UserControl_MouseDown(1, 0, 0, 0)
    End Sub

Private Sub p_KeyUp(KeyCode As Integer, Shift As Integer) 'Like MouseUp
    If KeyCode <> vbKeySpace Then Exit Sub ' and come out
       RaiseEvent KeyUp(KeyCode, Shift)
           Call UserControl_Click 'we didn't call MouseUp 'cause he will not change the picture
               btnDown = 0 'this is also in sub MouseUp
End Sub
Private Sub p_AccessKeyPress(KeyAscii As Integer)
  RaiseEvent Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, x, Y)
End Sub

Private Sub p_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, x, Y)
End Sub

Private Sub p_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call UserControl_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub p_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, x, Y)
End Sub
Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    Call UserControl_Click
End Sub

Private Sub p_Click()
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
        If mValue = Checked Then
            Value = Unchecked
                ElseIf mValue = Unchecked Then
                    Value = Checked
                ElseIf mValue = Mixed Then
            Value = Unchecked
        End If
    DisablePc
End Sub

Private Sub UserControl_Initialize()
    pc.Picture = pcChoice(3).Picture
    DisablePc
    UserControl_Resize
    UserControl.BackColor = mBackColor
    chVal = 1
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
    BackColor = defBackColor
    CheckBoxLook = XP_Default
    Value = Unchecked
    Caption = Ambient.DisplayName
    Set Font = UserControl.Ambient.Font
    ForeColor = defForeColor
End Sub

Private Sub UserControl_Resize()
    UserControl.ScaleMode = 1
    p.Height = 195
    p.Width = 195
    p.Left = 0
    p.Top = (UserControl.Height - p.Height) \ 2
    lbl.Top = (UserControl.Height - lbl.Height) \ 2
    lbl.Left = 240
End Sub

Private Function DisablePc()
    If Enabled = True Then
        If mValue = Checked Then
            p.Picture = pc.GraphicCell(4)
                ElseIf mValue = Mixed Then
                    p.Picture = pc.GraphicCell(8)
                ElseIf mValue = Unchecked Then
            p.Picture = pc.GraphicCell(0)
        End If
            Else: EnablePc
    End If
End Function

Private Function EnablePc()
    If mValue = Checked Then
        p.Picture = pc.GraphicCell(7)
            ElseIf mValue = Mixed Then
                p.Picture = pc.GraphicCell(11)
            ElseIf mValue = Unchecked Then
        p.Picture = pc.GraphicCell(3)
    End If
End Function

Private Sub DoIt(z As Integer)
    pc.Picture = pcChoice(z).Picture
End Sub

Private Sub CheckEnabled()
    If Enabled = False Then
        EnablePc
            lbl.ForeColor = &H80000011
                Timer1.Enabled = False
            Else: DisablePc
        lbl.ForeColor = mForeColor
    End If
End Sub

Private Sub p_GotFocus() 'in case that you move with key "Tab" or mouse click, picure p get focus
    Call UserControl_MouseMove(0, 0, 0, 0)
        Timer1.Enabled = False 'timer must be disabled 'cause we will not see the change
End Sub

Private Sub p_LostFocus() 'here p losts focus and must change picture
    chVal = 11 'must be done 'cause else will not change the picture
        Call UserControl_MouseMove(0, 0, 0, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    CheckBoxLook = PropBag.ReadProperty("CheckBoxLook", mPicDefault)
    Value = PropBag.ReadProperty("Value", defValue)
    Caption = PropBag.ReadProperty("Caption", "CheckBox1")
    BackColor = PropBag.ReadProperty("BackColor", defBackColor)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    ForeColor = PropBag.ReadProperty("ForeColor", defForeColor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("CheckBoxLook", mPic, defPic)
    Call PropBag.WriteProperty("Value", mValue, defValue)
    Call PropBag.WriteProperty("Caption", lbl.Caption, "CheckBox")
    Call PropBag.WriteProperty("BackColor", mBackColor, defBackColor)
    Call PropBag.WriteProperty("Font", mFont, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mForeColor, defForeColor)
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)
    UserControl.Enabled() = NewEnabled
    CheckEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get CheckBoxLook() As Pict
    CheckBoxLook = mPic
End Property

Public Property Let CheckBoxLook(ByVal NewCheckBoxLook As Pict)
    mPic = NewCheckBoxLook
    PropertyChanged "CheckBoxLook"
    DoIt (mPic)
    CheckEnabled
End Property

Public Property Get Value() As State
    Value = mValue
End Property

Public Property Let Value(ByVal NewValue As State)
    mValue = NewValue
    DisablePc
    PropertyChanged "Value"
End Property

Public Property Get Caption() As String
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    lbl.Caption() = NewCaption
    Call UserControl_Resize
    PropertyChanged "Caption"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    mBackColor = NewBackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = mBackColor
End Property

Public Property Get Font() As Font
    Set Font = mFont
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set mFont = NewFont
    Set UserControl.Font = NewFont
    Set lbl.Font = mFont
    Call UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    mForeColor = NewForeColor
    CheckEnabled
    PropertyChanged "ForeColor"
End Property

Private Sub Timer1_Timer()
    Dim dot As POINT_API
    UserControl.ScaleMode = 3 'must have this 'cause of x and y, to know how to calc
    Call GetCursorPos(dot) 'get mouse position
        ScreenToClient UserControl.hWnd, dot 'must have
  
  'checking if mouse is on our control by x and y
            If dot.x < UserControl.ScaleLeft Or _
                dot.Y < UserControl.ScaleTop Or _
                    dot.x > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
                        dot.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
                        
                            If btnDown = 1 Then Exit Sub 'in case that user clicked and did not
                            DisablePc                            'left the button, this will prevent from calling
                        Timer1.Enabled = False            ' DisablePc with no end
                RaiseEvent MouseOut
            End If
End Sub
