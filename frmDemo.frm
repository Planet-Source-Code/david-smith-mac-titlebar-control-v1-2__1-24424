VERSION 5.00
Object = "{F2316E7B-7C5B-4935-B1AF-B11414894D6F}#2.0#0"; "prjMacTitle.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "David's Mac Control"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Control Options"
      Height          =   3375
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   2295
      Begin VB.CheckBox chkEHelp 
         Caption         =   "Enable Help"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkEClose 
         Caption         =   "Enable Close"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkEMin 
         Caption         =   "Enable Minimize"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkEMax 
         Caption         =   "Enable Maximize"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Show Help"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMaximize 
         Caption         =   "Show Maximize"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMinimize 
         Caption         =   "Show Minimize"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "Show Close"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin VB.Frame frmCaption 
      Caption         =   "Control Caption"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2535
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
   End
   Begin prjMacTitle.ctlMacTitle ctlMacTitle1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      Caption         =   "David's Mac Control"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CloseE As Boolean
Dim CloseV As Boolean

Dim MinE As Boolean
Dim MinV As Boolean

Dim MaxE As Boolean
Dim MaxV As Boolean

Dim HelpE As Boolean
Dim HelpV As Boolean





Private Sub cmdUpdate_Click()
  With ctlMacTitle1
    .Caption = txtCaption.Text
    
    'B - Close Button
    Select Case chkClose.Value
    Case 1:
      .CloseV = True
      
    Case Else:
      .CloseV = False
    End Select
    
    Select Case chkEClose.Value
    Case 1:
      .CloseE = True
    Case Else:
      .CloseE = False
    End Select
    'E - Close Button
    
    'B - Max Button
    Select Case chkMaximize.Value
    Case 1:
      .MaxV = True
    Case Else:
      .MaxV = False
    End Select
    
    Select Case chkEMax.Value
    Case 1:
      .MaxE = True
    Case Else:
      .MaxE = False
    End Select
    
    'E - Max Button
    
    'B - Minimize
    Select Case chkMinimize.Value
    Case 1:
      .MinV = True
    Case Else:
      .MinV = False
    End Select
    
    Select Case chkEMin.Value
    Case 1:
      .MinE = True
    Case Else:
      .MinE = False
    End Select
    'E - Minimize
    
    'B - Help
    Select Case chkHelp.Value
    Case 1:
      .HelpV = True
    Case Else:
      .HelpV = False
    End Select
    
    Select Case chkEHelp.Value
    Case 1:
      .HelpE = True
    Case Else:
      .HelpE = False
    End Select
    .Refresh
  End With
End Sub

Private Sub ctlMacTitle1_CloseMe()
  End
End Sub

Private Sub ctlMacTitle1_HelpMe()
  MsgBox "macControl V 1.1" & vbCrLf & "Created By: David Smith" & vbCrLf & "eMail: daveismith@hotmail.com", vbInformation
End Sub

Private Sub ctlMacTitle1_MaximizeMe()
  If Me.WindowState = vbMaximized Then
    Me.WindowState = vbNormal
'    ctlMacTitle1.Width = Me.ScaleWidth
  Else
    Me.WindowState = vbMaximized
  End If
End Sub

Private Sub ctlMacTitle1_MinimizeMe()
  Me.WindowState = vbMinimized
End Sub





Private Sub Form_Load()
txtCaption = Me.Caption
With ctlMacTitle1
  txtCaption = .Caption
  .MaxE = True
  .MaxV = True
End With

End Sub

Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then ctlMacTitle1.Width = Me.ScaleWidth
End Sub
