VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   ScaleHeight     =   390
   ScaleWidth      =   1800
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1470
      Picture         =   "UserControl1.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   225
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1425
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -30
         TabIndex        =   1
         Top             =   -30
         Width           =   1365
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************************************
'*  Copyright (C) Kobi Vazana 2001 - All Rights Reserved        *
'*                                                              *
'*  FILE:  KDCFlatCombo.ctl                                        *
'*                                                              *
'*  DESCRIPTION:                                                *
'*      Gradient button with color sets that can be modified    *
'*      At Design time ,Centered icon ,and all min Properties   *
'****************************************************************

Private MyText As String
Private MyFont As Font
Private MyForeColor As OLE_COLOR
Private MyBackColor As OLE_COLOR
Private NewButtonIcon As Picture

Private MyLocked As Boolean
Private MyEnabled As Boolean
Private MyHasFocus As Boolean
Private MyLeftFocus As Boolean
Private MyRightToLeft As Boolean
Private MySorted As Boolean

Private Const DefText = "KDC"
Private Const MyDefEnabled = True
Private Const DefForeColor = vbBlack
Private Const DefRightToLeft = False
Private Const DefBackColor = &HFFFFFF
Private Const DefLocked = False
Private Const DefSorted = False

Public Event Click()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event Resize()

Private Sub Pic2_Click()
    Combo1.SetFocus
    SendKeys "%{Down}"
End Sub
Private Sub UserControl_Initialize()
    Call UserControl_Resize
End Sub
Private Sub UserControl_Resize()
    Pic1.Width = UserControl.Width
    Combo1.Width = Pic1.Width + 22
    UserControl.Height = Pic1.Height
    Pic2.Width = 250
    Pic2.Height = 285
    If Combo1.RightToLeft = True Then
        Pic2.Left = 0
    Else
        Pic2.Left = UserControl.Width - 250
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", MyText, DefText)
    Call PropBag.WriteProperty("ForeColor", MyForeColor, DefForeColor)
    Call PropBag.WriteProperty("Font", MyFont, Ambient.Font)
    Call PropBag.WriteProperty("ButtonIcon", Me.ButtonIcon, Nothing)
    Call PropBag.WriteProperty("Enabled", MyEnabled, MyDefEnabled)
    Call PropBag.WriteProperty("RightToLeft", MyRightToLeft, DefRightToLeft)
    Call PropBag.WriteProperty("BackColor", MyBackColor, DefBackColor)
    Call PropBag.WriteProperty("Locked", MyLocked, DefLocked)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text = PropBag.ReadProperty("Text", DefText)
    ForeColor = PropBag.ReadProperty("ForeColor", DefForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set ButtonIcon = PropBag.ReadProperty("ButtonIcon", Nothing)
    Enabled = PropBag.ReadProperty("Enabled", MyDefEnabled)
    RightToLeft = PropBag.ReadProperty("RightToLeft", DefRightToLeft)
    BackColor = PropBag.ReadProperty("BackColor", DefBackColor)
    Locked = PropBag.ReadProperty("Locked", DefLocked)
End Sub
Private Sub UserControl_InitProperties()
    Text = DefText
    ForeColor = DefForeColor
    BackColor = DefBackColor
    Set Font = Ambient.Font
    Enabled = MyDefEnabled
    RightToLeft = DefRightToLeft
    Locked = DefLocked
End Sub
Public Property Get ButtonIcon() As Picture
    Set ButtonIcon = Pic2.Picture
End Property
Public Property Set ButtonIcon(ByVal NewButtonIcon As Picture)
    Set Pic2.Picture = NewButtonIcon
    Set Pic2.Picture = NewButtonIcon
    Call UserControl_Resize
PropertyChanged "ButtonIcon"
End Property
Public Property Get Enabled() As Boolean
    Enabled = MyEnabled
End Property
Public Property Let Enabled(ByVal vData As Boolean)
    MyEnabled = vData
    UserControl.Enabled = MyEnabled
    Call UserControl_Resize
PropertyChanged "Enabled"
End Property
Public Property Get Locked() As Boolean
    Locked = MyLocked
End Property
Public Property Let Locked(ByVal vData As Boolean)
    MyLocked = vData
    Combo1.Locked = MyLocked
    Call UserControl_Resize
PropertyChanged "Locked"
End Property
Public Property Get Font() As Font
    Set Font = MyFont
End Property
Public Property Set Font(ByVal vData As Font)
    Set MyFont = vData
    Set Combo1.Font = vData
    Call UserControl_Resize
PropertyChanged "Font"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MyForeColor
End Property
Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    MyForeColor = vData
    Combo1.ForeColor = MyForeColor
PropertyChanged "ForeColor"
End Property
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "34"
    Text = MyText
End Property
Public Property Let Text(ByVal vData As String)
    MyText = vData
    Combo1.Text = vData
PropertyChanged "Text"
End Property
Public Property Get RightToLeft() As Boolean
    RightToLeft = MyRightToLeft
End Property
Public Property Let RightToLeft(ByVal vData As Boolean)
    MyRightToLeft = vData
    Combo1.RightToLeft = vData
    Call UserControl_Resize
PropertyChanged "RightToLeft"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = MyBackColor
End Property
Public Property Let BackColor(ByVal vData As OLE_COLOR)
    MyBackColor = vData
    Combo1.BackColor = vData
    Call UserControl_Resize
PropertyChanged "BackColor"
End Property
''------------------------------------------------------------------
Public Sub AddItem(Item As Variant)
    Combo1.AddItem CStr(Item)
End Sub
Public Sub Clear()
    Combo1.Clear
End Sub
Public Sub Refresh()
    Combo1.Refresh
End Sub
Public Sub RemoveItem(Index As Integer)
    Combo1.RemoveItem Index
End Sub
Public Property Get ListIndex() As String
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = Combo1.ListIndex
End Property
Public Property Let ListIndex(ByVal vData As String)
    Combo1.ListIndex = vData
PropertyChanged "ListIndex"
End Property
Public Property Get ListCount() As String
Attribute ListCount.VB_MemberFlags = "400"
    ListCount = Combo1.ListCount
End Property
Public Property Let ListCount(ByVal vData As String)
    Combo1.ListCount = vData
PropertyChanged "ListCount"
End Property


