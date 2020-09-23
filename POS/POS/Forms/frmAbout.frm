VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":058A
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin osenxpsuite.OsenXPButton btnSystem 
      Height          =   375
      Left            =   5295
      TabIndex        =   2
      Top             =   3660
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "System Info.."
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
      MICON           =   "frmAbout.frx":AD6E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   3135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmAbout.frx":AD8A
      Top             =   1125
      Width           =   3780
   End
   Begin osenxpsuite.OsenXPButton btnClose 
      Height          =   375
      Left            =   5295
      TabIndex        =   3
      Top             =   4065
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "&Close"
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
      MICON           =   "frmAbout.frx":AEE5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      OffsetLeft      =   0
      OffsetTop       =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2880
      Picture         =   "frmAbout.frx":AF01
      Top             =   195
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning: Protected with copyright law. Any illegal actions is punishable by law."
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   3180
      TabIndex        =   1
      Top             =   195
      Width           =   3225
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnSystem_Click()
On Error GoTo err:

frmMain.StartSysInfo
Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description

End Sub
