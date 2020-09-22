VERSION 5.00
Begin VB.UserControl ctlMacTitle 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   Begin VB.Image imgHelp 
      Height          =   195
      Index           =   0
      Left            =   1440
      Picture         =   "ctlMacTitle.ctx":0000
      Top             =   30
      Width           =   195
   End
   Begin VB.Image imgHelp 
      Height          =   195
      Index           =   1
      Left            =   1680
      Picture         =   "ctlMacTitle.ctx":024A
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHelp 
      Height          =   195
      Index           =   2
      Left            =   1680
      Picture         =   "ctlMacTitle.ctx":0494
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgHelp 
      Height          =   195
      Index           =   3
      Left            =   1680
      Picture         =   "ctlMacTitle.ctx":06DE
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   3
      Left            =   240
      Picture         =   "ctlMacTitle.ctx":0928
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   2
      Left            =   240
      Picture         =   "ctlMacTitle.ctx":0B72
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   1
      Left            =   240
      Picture         =   "ctlMacTitle.ctx":0DBC
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Index           =   3
      Left            =   720
      Picture         =   "ctlMacTitle.ctx":1006
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Index           =   2
      Left            =   720
      Picture         =   "ctlMacTitle.ctx":1250
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Index           =   1
      Left            =   720
      Picture         =   "ctlMacTitle.ctx":149A
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMaximize 
      Height          =   195
      Index           =   3
      Left            =   1200
      Picture         =   "ctlMacTitle.ctx":16E4
      Top             =   840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMaximize 
      Height          =   195
      Index           =   2
      Left            =   1200
      Picture         =   "ctlMacTitle.ctx":192E
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMaximize 
      Height          =   195
      Index           =   1
      Left            =   1200
      Picture         =   "ctlMacTitle.ctx":1B78
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgMaximize 
      Height          =   195
      Index           =   0
      Left            =   960
      Picture         =   "ctlMacTitle.ctx":1DC2
      Top             =   30
      Width           =   195
   End
   Begin VB.Image imgMinimize 
      Height          =   195
      Index           =   0
      Left            =   480
      Picture         =   "ctlMacTitle.ctx":200C
      Top             =   30
      Width           =   195
   End
   Begin VB.Image imgClose 
      Height          =   195
      Index           =   0
      Left            =   30
      Picture         =   "ctlMacTitle.ctx":2256
      Top             =   30
      Width           =   195
   End
End
Attribute VB_Name = "ctlMacTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lhs_left As Integer
Dim lhs_right As Integer
Dim rhs_left As Integer
Dim rhs_right As Integer
Dim l_top As Integer
Dim sc As Integer
Dim dorhs As Boolean
Dim dolhs As Boolean
Dim X As Integer
Dim maclefttext As Integer
Dim mactoptext As Integer

Dim CurrentIcon As PictureBox

Dim CloseVal As Boolean
Dim MinimizeVal As Boolean
Dim MaximizeVal As Boolean
Dim HelpVal As Boolean

Dim CloseVis As Boolean
Dim MinVis As Boolean
Dim MaxVis As Boolean
Dim HelpVis As Boolean

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Dim title As String

  Public Event CloseMe()
  Public Event MaximizeMe()
  Public Event MinimizeMe()
  Public Event HelpMe()
'Default Property Values:
Const m_def_Caption = "MacControl"
Const m_def_HelpV = True
Const m_def_CloseV = True
Const m_def_MaxE = True
Const m_def_MinE = True
Const m_def_CloseE = True
Const m_def_HelpE = True
Const m_def_MaxV = True
Const m_def_MinV = True
'Property Variables:
Dim m_Caption As String
Dim m_HelpV As Boolean
Dim m_CloseV As Boolean
Dim m_MaxE As Boolean
Dim m_MinE As Boolean
Dim m_CloseE As Boolean
Dim m_HelpE As Boolean
Dim m_MaxV As Boolean
Dim m_MinV As Boolean




Public Sub FormDrag(FrmHwnd As Long)
   ReleaseCapture
   Call SendMessage(FrmHwnd, &HA1, 2, 0&)
End Sub


Private Sub imgClose_Click(Index As Integer)
  RaiseEvent CloseMe
End Sub

Private Sub imgClose_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgClose(0).Picture = imgClose(2).Picture
End Sub

Private Sub imgClose_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgClose(0).Picture = imgClose(1).Picture
End Sub

Private Sub imgHelp_Click(Index As Integer)
  RaiseEvent HelpMe
End Sub

Private Sub imgHelp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgHelp(0).Picture = imgHelp(2).Picture
End Sub

Private Sub imgHelp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgHelp(0).Picture = imgHelp(1).Picture
End Sub

Private Sub imgMaximize_Click(Index As Integer)
  RaiseEvent MaximizeMe
End Sub

Private Sub imgMaximize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMaximize(0).Picture = imgMaximize(2).Picture
End Sub

Private Sub imgMaximize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMaximize(0).Picture = imgMaximize(1).Picture
End Sub

Private Sub imgMinimize_Click(Index As Integer)
  RaiseEvent MinimizeMe
End Sub

Private Sub imgMinimize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMinimize(0).Picture = imgMinimize(2).Picture
End Sub

Private Sub imgMinimize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMinimize(0).Picture = imgMinimize(1).Picture
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag UserControl.Parent.hwnd
End Sub

Private Sub UserControl_Resize()
  UserControl.Cls
  title = m_Caption
  CloseVis = m_CloseV
  HelpVis = m_HelpV
  MinVis = m_MinV
  MaxVis = m_MaxV

  CloseVal = m_CloseE
  HelpVal = m_HelpE
  MinimizeVal = m_MinE
  MaximizeVal = m_MaxE


  Call CloseBut
  Call MinBut
  Call MaxBut
  Call HelpBut
  
With UserControl
    .FontTransparent = False
    .AutoRedraw = True
    .ScaleMode = 3
    .BackColor = &HCCCCCC
    .BorderStyle = 0
    .ForeColor = QBColor(0)
    .Font = "Chicago"
    .FontBold = False
    .FontSize = 10
End With

    If (UserControl.ScaleWidth / 2) - (UserControl.TextWidth(title) / 2) <= 8 Then title = ""
        



    l_top = UserControl.ScaleHeight / 2 - 6
    dolhs = True
    dorhs = True
    
    If title = "" Then
        lhs_right = imgHelp(0).Left - 8
        dorhs = False
        GoTo drawit
    End If
            sc = UserControl.ScaleWidth
            lhs_right = ((sc / 2) - (UserControl.TextWidth(title) / 2)) - 8
            lhs_right = Int(lhs_right)
            rhs_left = ((sc / 2) + (UserControl.TextWidth(title) / 2)) + 8
            rhs_left = Int(rhs_left)
            rhs_right = imgHelp(0).Left - 8
            
 
drawit:
                    If dolhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            UserControl.Line (lhs_left - 1, X)-(lhs_right, X), &HFFFFFF
                            UserControl.Line (lhs_left, X + 1.5)-(lhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    If dorhs = True Then
                        For X = l_top To l_top + 10 Step 2
                            UserControl.Line (rhs_left - 1, X)-(rhs_right, X), &HFFFFFF
                            UserControl.Line (rhs_left, X + 1.5)-(rhs_right + 1, X + 1.5), &H666666
                        Next X
                    End If
                    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - 1), &H666666
                    
                    maclefttext = rhs_left - (UserControl.TextWidth(title)) - 8
                    
                    UserControl.CurrentX = maclefttext
                    mactoptext = (UserControl.ScaleHeight / 2) - (UserControl.TextHeight(title) / 2)
                    UserControl.CurrentY = mactoptext
                    UserControl.Print title
End Sub
'
'Public Function ChangeTitle(Caption As String, Icon As String)
'  title = Caption
'  imgIcon.Width = 0
'  If Icon <> "" Then CurrentIcon.Picture = LoadPicture(Icon)
'  Call UserControl_Resize
'End Function
'
'Public Function MinimizeV(Enabled As Boolean, Visible As Boolean)
'  MinimizeVal = Enabled
'  MinVis = Visible
'  Call UserControl_Resize
'End Function
'
'Public Function HelpV(Enabled As Boolean, Visible As Boolean)
'  HelpVal = Enabled
'  HelpVis = Visible
'  Call UserControl_Resize
'End Function
'
'Public Function MaximizeV(Enabled As Boolean, Visible As Boolean)
'  MaximizeVal = Enabled
'  MaxVis = Visible
'  Call UserControl_Resize
'End Function
'
'
'Public Function CloseV(Enabled As Boolean, Visible As Boolean)
'  CloseVal = Enabled
'  CloseVis = Visible
'  Call UserControl_Resize
'End Function

Public Sub CloseBut()
Select Case CloseVis
  Case True:
    imgClose(0).Visible = True
    lhs_left = 23
  Case False
    imgClose(0).Visible = False
    lhs_left = 8
End Select

Select Case CloseVal
    Case True:
      imgClose(0).Picture = imgClose(1).Picture
      imgClose(0).Enabled = True
    Case False:
      imgClose(0).Enabled = False
      imgClose(0).Picture = imgClose(3).Picture
End Select
End Sub

Public Sub MinBut()

Select Case MinVis
  Case True:
    imgMinimize(0).Visible = True
    imgMinimize(0).Left = UserControl.ScaleWidth - 15
  Case False:
    imgMinimize(0).Visible = False
    imgMinimize(0).Left = UserControl.ScaleWidth
End Select

  Select Case MinimizeVal
    Case True:
      imgMinimize(0).Picture = imgMinimize(1).Picture
      imgMinimize(0).Enabled = True
    Case False:
      imgMinimize(0).Enabled = False
      imgMinimize(0).Picture = imgMinimize(3).Picture
  End Select
End Sub

Public Sub MaxBut()
  Select Case MaxVis
  Case True:
    imgMaximize(0).Visible = True
    imgMaximize(0).Left = imgMinimize(0).Left - 15
  Case False:
    imgMaximize(0).Visible = False
    imgMaximize(0).Left = imgMinimize(0).Left
  End Select
  
  Select Case MaximizeVal
    Case True:
      imgMaximize(0).Picture = imgMaximize(1).Picture
      imgMaximize(0).Enabled = True
    Case False:
      imgMaximize(0).Enabled = False
      imgMaximize(0).Picture = imgMaximize(3).Picture
  End Select
End Sub

Public Sub HelpBut()
  Select Case HelpVis
  Case True:
    imgHelp(0).Visible = True
    imgHelp(0).Left = imgMaximize(0).Left - 16
  Case False:
    imgHelp(0).Visible = False
    imgHelp(0).Left = imgMaximize(0).Left
  End Select
  
  Select Case HelpVal
    Case True:
      imgHelp(0).Picture = imgHelp(1).Picture
      imgHelp(0).Enabled = True
    Case False:
      imgHelp(0).Enabled = False
      imgHelp(0).Picture = imgHelp(3).Picture
  End Select
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,MacControl
Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  m_Caption = New_Caption
  PropertyChanged "Caption"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HelpV() As Boolean
  HelpV = m_HelpV
End Property

Public Property Let HelpV(ByVal New_HelpV As Boolean)
  m_HelpV = New_HelpV
  PropertyChanged "HelpV"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CloseV() As Boolean
  CloseV = m_CloseV
End Property

Public Property Let CloseV(ByVal New_CloseV As Boolean)
  m_CloseV = New_CloseV
  PropertyChanged "CloseV"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MaxE() As Boolean
  MaxE = m_MaxE
End Property

Public Property Let MaxE(ByVal New_MaxE As Boolean)
  m_MaxE = New_MaxE
  PropertyChanged "MaxE"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MinE() As Boolean
  MinE = m_MinE
End Property

Public Property Let MinE(ByVal New_MinE As Boolean)
  m_MinE = New_MinE
  PropertyChanged "MinE"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get CloseE() As Boolean
  CloseE = m_CloseE
End Property

Public Property Let CloseE(ByVal New_CloseE As Boolean)
  m_CloseE = New_CloseE
  PropertyChanged "CloseE"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get HelpE() As Boolean
  HelpE = m_HelpE
End Property

Public Property Let HelpE(ByVal New_HelpE As Boolean)
  m_HelpE = New_HelpE
  PropertyChanged "HelpE"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MaxV() As Boolean
  MaxV = m_MaxV
End Property

Public Property Let MaxV(ByVal New_MaxV As Boolean)
  m_MaxV = New_MaxV
  PropertyChanged "MaxV"
  Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get MinV() As Boolean
  MinV = m_MinV
End Property

Public Property Let MinV(ByVal New_MinV As Boolean)
  m_MinV = New_MinV
  PropertyChanged "MinV"
  Call UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_Caption = m_def_Caption
  m_HelpV = m_def_HelpV
  m_CloseV = m_def_CloseV
  m_MaxE = m_def_MaxE
  m_MinE = m_def_MinE
  m_CloseE = m_def_CloseE
  m_HelpE = m_def_HelpE
  m_MaxV = m_def_MaxV
  m_MinV = m_def_MinV
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  m_HelpV = PropBag.ReadProperty("HelpV", m_def_HelpV)
  m_CloseV = PropBag.ReadProperty("CloseV", m_def_CloseV)
  m_MaxE = PropBag.ReadProperty("MaxE", m_def_MaxE)
  m_MinE = PropBag.ReadProperty("MinE", m_def_MinE)
  m_CloseE = PropBag.ReadProperty("CloseE", m_def_CloseE)
  m_HelpE = PropBag.ReadProperty("HelpE", m_def_HelpE)
  m_MaxV = PropBag.ReadProperty("MaxV", m_def_MaxV)
  m_MinV = PropBag.ReadProperty("MinV", m_def_MinV)
  Call UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
  Call PropBag.WriteProperty("HelpV", m_HelpV, m_def_HelpV)
  Call PropBag.WriteProperty("CloseV", m_CloseV, m_def_CloseV)
  Call PropBag.WriteProperty("MaxE", m_MaxE, m_def_MaxE)
  Call PropBag.WriteProperty("MinE", m_MinE, m_def_MinE)
  Call PropBag.WriteProperty("CloseE", m_CloseE, m_def_CloseE)
  Call PropBag.WriteProperty("HelpE", m_HelpE, m_def_HelpE)
  Call PropBag.WriteProperty("MaxV", m_MaxV, m_def_MaxV)
  Call PropBag.WriteProperty("MinV", m_MinV, m_def_MinV)
  Call UserControl_Resize
End Sub


Public Sub Refresh()
  Call UserControl_Resize
End Sub
