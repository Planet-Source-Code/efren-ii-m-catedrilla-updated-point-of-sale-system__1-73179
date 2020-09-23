VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmPrivileges 
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrivileges.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   105
      TabIndex        =   8
      Top             =   3405
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   53
   End
   Begin VB.CheckBox chkAdmin 
      Caption         =   "Allow user to access privileges even if the users is not in administrator mode."
      Height          =   240
      Left            =   270
      TabIndex        =   7
      Top             =   2535
      Width           =   5940
   End
   Begin VB.CheckBox chkReports 
      Caption         =   "Allow user to print transaction reports."
      Height          =   240
      Left            =   270
      TabIndex        =   6
      Top             =   2220
      Width           =   4080
   End
   Begin VB.CheckBox chkDelete 
      Caption         =   "Allow user to delete sensitive data."
      Height          =   240
      Left            =   270
      TabIndex        =   5
      Top             =   1920
      Width           =   4080
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Allow the user to change his password."
      Height          =   240
      Left            =   270
      TabIndex        =   4
      Top             =   1605
      Width           =   4080
   End
   Begin VB.CheckBox chkEdit 
      Caption         =   "Allow user to edit data in the system."
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   1275
      Width           =   4080
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Top             =   3495
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
      MICON           =   "frmPrivileges.frx":058A
      PICN            =   "frmPrivileges.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   4035
      TabIndex        =   10
      Top             =   3495
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
      MICON           =   "frmPrivileges.frx":0B40
      PICN            =   "frmPrivileges.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label lblUsers 
      BackStyle       =   0  'Transparent
      Caption         =   "admin"
      Height          =   375
      Left            =   5865
      TabIndex        =   11
      Top             =   1095
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Privileges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      TabIndex        =   2
      Top             =   240
      Width           =   3165
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   105
      Picture         =   "frmPrivileges.frx":10F6
      Top             =   105
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7545
   End
End
Attribute VB_Name = "frmPrivileges"
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
execSQL "select * from privileges where users='" & lblUsers.Caption & "'"

If Adding = True Then

With RS
.AddNew
.Fields!update_data = chkEdit.Value
.Fields!change_pass = chkPassword.Value
.Fields!delete_data = chkDelete.Value
.Fields!print_reports = chkReports.Value
.Fields!admin = chkAdmin.Value
.Fields!users = lblUsers.Caption
.Update
.Requery
End With

MsgBox "User Privileges has been successfully saved.", vbInformation
Unload Me

ElseIf Editing = True Then
With RS
.Fields!update_data = chkEdit.Value
.Fields!change_pass = chkPassword.Value
.Fields!delete_data = chkDelete.Value
.Fields!print_reports = chkReports.Value
.Fields!admin = chkAdmin.Value
.Update
.Requery
End With
MsgBox "User Privileges has been successfully saved.", vbInformation
End If

Set RS = Nothing
 Exit Sub
err:
 MsgBox "Error #" & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
CenterForm Me
execSQL "select * from privileges where users='" & frmMain.StatusBar1.Panels(4).Text & "'"

If RS.RecordCount < 1 Then
  Adding = True
  Editing = False
  Else
  Editing = True
  Adding = False
Set RS = Nothing
load_privileges
End If
Set RS = Nothing
End Sub


Private Sub load_privileges()
 On Error GoTo err:
  execSQL "select * from privileges"
  
  With RS
  chkEdit.Value = .Fields!update_data
  chkPassword.Value = .Fields!change_pass
  chkDelete.Value = .Fields!delete_data
  chkReports.Value = .Fields!print_reports
  chkAdmin.Value = .Fields!admin
  End With
  
  Set RS = Nothing
  Exit Sub
  
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
  
End Sub
