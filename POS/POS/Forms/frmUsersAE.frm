VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmUsersAE 
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsersAE.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbUserType 
      Height          =   315
      ItemData        =   "frmUsersAE.frx":058A
      Left            =   1545
      List            =   "frmUsersAE.frx":058C
      TabIndex        =   2
      Top             =   2085
      Width           =   2310
   End
   Begin VB.TextBox txtEmployees 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   2445
      Width           =   3000
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1545
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1725
      Width           =   2550
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   1365
      Width           =   2550
   End
   Begin VB.TextBox txtUserNo 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1005
      Width           =   1620
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   810
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPButton btnSearch 
      Height          =   315
      Left            =   4545
      TabIndex        =   13
      ToolTipText     =   "Select Category"
      Top             =   2445
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   ""
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
      MICON           =   "frmUsersAE.frx":058E
      PICN            =   "frmUsersAE.frx":05AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   4050
      TabIndex        =   14
      Top             =   3015
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Cancel"
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
      MICON           =   "frmUsersAE.frx":0B44
      PICN            =   "frmUsersAE.frx":0B60
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   2715
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Save"
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
      MICON           =   "frmUsersAE.frx":10FA
      PICN            =   "frmUsersAE.frx":1116
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label lblEmployees 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4980
      TabIndex        =   16
      Top             =   2430
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2445
      Width           =   1140
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User Type"
      Height          =   285
      Left            =   135
      TabIndex        =   11
      Top             =   2085
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   285
      Left            =   135
      TabIndex        =   10
      Top             =   1725
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   1365
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User No."
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   1005
      Width           =   1140
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Please provide all the information needed. Add / Edit User Records."
      Height          =   450
      Left            =   870
      TabIndex        =   5
      Top             =   165
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   75
      Picture         =   "frmUsersAE.frx":16B0
      Top             =   15
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   -45
      TabIndex        =   4
      Top             =   0
      Width           =   7125
   End
End
Attribute VB_Name = "frmUsersAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
 On Error GoTo err:
If Adding = True Then
 execSQL "select * from users"
 
 With RS
 .AddNew
 .Fields!user_id = txtUserNo.Text
 .Fields!username = txtUserName.Text
 .Fields!password = Encode(txtPassword.Text)
 .Fields!usertype = cmbUserType.Text
 .Fields!emp_id = lblEmployees.Caption
 .Update
 .Requery
 End With
 MsgBox "New record has been successfully saved.", vbInformation
 
 ElseIf Editing = True Then
 execSQL "select * from users where user_id=" & txtUserNo.Text & ""

 With RS
 .Fields!username = txtUserName.Text
 .Fields!password = Encode(txtPassword.Text)
 .Fields!usertype = cmbUserType.Text
 .Fields!emp_id = lblEmployees.Caption
 .Update
 .Requery
 End With
 MsgBox "Record changes has been successfully saved.", vbInformation
 Unload Me
 End If

 Set RS = Nothing
 frmUsers.FillListView
   Exit Sub
err:
  MsgBox "Error #" & err.Number & " Description: " & err.Description, vbExclamation
  
End Sub

Private Sub btnSearch_Click()
frmEmpUsers.Show vbModal
End Sub

Private Sub Form_Load()
cmbUserType.List(0) = ""
cmbUserType.List(1) = "User"
cmbUserType.List(2) = "Cashier"
cmbUserType.List(3) = "Manager"
cmbUserType.List(4) = "Administrator"

If Adding = True Then
 intCount
End If

End Sub


Private Sub intCount()
  Dim count As Integer
 On Error GoTo err:
 execSQL "Select * From users"
 count = 999
 
 With RS
 If .RecordCount < 0 Then
 
   txtUserNo.Text = "1000"
    Else
     count = count + .RecordCount + 1
      txtUserNo.Text = count
     
 End If
 End With
 

 Set RS = Nothing
 Exit Sub
err:
 MsgBox "Error # " & err.Number & " Description " & err.Description, vbExclamation
 
End Sub
