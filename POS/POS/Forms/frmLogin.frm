VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5115
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1470
      TabIndex        =   0
      Top             =   1170
      Width           =   2985
   End
   Begin VB.TextBox txtUser 
      Height          =   330
      Left            =   1755
      TabIndex        =   22
      Text            =   "localhost"
      Top             =   4125
      Width           =   2070
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Apply Changes"
      Height          =   375
      Left            =   3675
      TabIndex        =   21
      Top             =   5265
      Width           =   1515
   End
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   30
      TabIndex        =   13
      Top             =   3375
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPFrame OsenXPFrame1 
      Height          =   1800
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   3175
      Caption         =   "My SQL Administrator Port"
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
      Begin VB.TextBox txtDatabase 
         Height          =   330
         Left            =   1635
         TabIndex        =   19
         Text            =   "pos"
         Top             =   1395
         Width           =   2070
      End
      Begin VB.TextBox txtPass 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1635
         PasswordChar    =   "*"
         TabIndex        =   17
         Text            =   "root"
         Top             =   1020
         Width           =   2070
      End
      Begin VB.TextBox txtHost 
         Height          =   330
         Left            =   1635
         TabIndex        =   14
         Text            =   "localhost"
         Top             =   270
         Width           =   2070
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Database:"
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   1410
         Width           =   1140
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   1020
         Width           =   1140
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "UserName:"
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Host:"
         Height          =   285
         Left            =   195
         TabIndex        =   15
         Top             =   300
         Width           =   1140
      End
   End
   Begin VB.TextBox txtPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1470
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1545
      Width           =   2985
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   1020
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   3255
      TabIndex        =   6
      Top             =   2010
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "&Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmLogin.frx":058A
      PICN            =   "frmLogin.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnLogin 
      Height          =   375
      Left            =   1935
      TabIndex        =   7
      Top             =   1995
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmLogin.frx":0B40
      PICN            =   "frmLogin.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnOption 
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   2850
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Option>>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "frmLogin.frx":10F6
      PICN            =   "frmLogin.frx":1112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   705
      TabIndex        =   10
      Top             =   150
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   30
      Picture         =   "frmLogin.frx":16AC
      Top             =   2490
      Width           =   240
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Restricted Area. Only Authorized person is allowed to access in this area."
      Height          =   390
      Left            =   315
      TabIndex        =   9
      Top             =   2505
      Width           =   3720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your username and password."
      Height          =   300
      Left            =   735
      TabIndex        =   8
      Top             =   495
      Width           =   3555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   390
      Left            =   195
      TabIndex        =   5
      Top             =   1530
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name :"
      Height          =   390
      Left            =   180
      TabIndex        =   4
      Top             =   1155
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmLogin.frx":1C36
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opt As Boolean
Public active As Boolean
Public conn As Boolean



Private Sub btnCancel_Click()
End
End Sub

Private Sub btnLogin_Click()
On Error GoTo err:
 
 MySql txtHost.Text, txtUser.Text, txtPass.Text, txtDatabase.Text
 
execSQL "select * from users where username='" & Trim(txtUserName.Text) & "' and password='" & Trim(Encode(txtPassword.Text)) & "'"

If RS.RecordCount < 1 Then
   MsgBox "Incorrect Username Or Password!" & ".Please enter the correct username and password.", vbExclamation
   Set CN = Nothing
   Else
   If active = True Then
    Unload Me
    Set RS = Nothing
    Exit Sub
    End If

   active = True
   frmMain.Show
   frmMain.StatusBar1.Panels(4).Text = txtUserName.Text
   frmFlash.Show vbModal
    Unload Me
End If

 Set RS = Nothing
   Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description
  Set RS = Nothing
End Sub

Private Sub btnOption_Click()

If opt = True Then
btnOption.Caption = "Option<<"
frmLogin.Height = 6075
opt = False
Else

btnOption.Caption = "Option>>"
frmLogin.Height = 3840
opt = True
End If
End Sub


Private Sub Form_Load()
opt = True
CenterForm Me
loadPort
End Sub



Private Sub FillPort()
 Dim s_host           As String
 Dim s_username       As String
 Dim s_password       As String
 Dim s_database       As String
 
 If chkSave.Value = 1 Then
 
 Open App.Path & "\host.dat" For Output As #1
 Open App.Path & "\username.dat" For Output As #2
 Open App.Path & "\password.dat" For Output As #3
 Open App.Path & "\database.dat" For Output As #4
 
 s_host = txtHost.Text
 Write #1, s_host
 
      s_username = txtUser.Text
      Write #2, s_username
         
         s_password = Encode(txtPass.Text)
         Write #3, s_password
         
            s_database = txtDatabase.Text
            Write #4, s_database
            
                Close #1
          Close #2
    Close #3
Close #4
 loadPort
 chkSave.Value = 0
 MsgBox "Connection changes has been saved successfully.", vbInformation
 Else
 
  Exit Sub
  
  End If
         
End Sub

Private Sub loadPort()
   Dim s_host       As String
   Dim s_username   As String
   Dim s_password   As String
   Dim s_database   As String
   
 Open App.Path & "\host.dat" For Input As #1
 Open App.Path & "\username.dat" For Input As #2
 Open App.Path & "\password.dat" For Input As #3
 Open App.Path & "\database.dat" For Input As #4
 
 Do Until EOF(1)
 Input #1, s_host
   txtHost.Text = s_host
 Loop
 
 Do Until EOF(2)
 Input #2, s_username
   txtUser.Text = s_username
 Loop
 
 Do Until EOF(3)
 Input #3, s_password
   txtPass.Text = DeCode(s_password)
 Loop
 
 Do Until EOF(4)
 Input #4, s_database
   txtDatabase.Text = s_database
 Loop
 
 Close #1
 Close #2
 Close #3
 Close #4
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
btnLogin_Click
End If

End Sub
