VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmEmployeesAE 
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "frmEmployeesAE.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4245
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   19
      Top             =   795
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   53
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1335
      TabIndex        =   5
      Top             =   3180
      Width           =   4860
   End
   Begin VB.TextBox txtContactNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   3
      Top             =   2460
      Width           =   2685
   End
   Begin VB.TextBox txtMiddleName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   2
      Top             =   2085
      Width           =   2445
   End
   Begin VB.TextBox txtLastName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   1710
      Width           =   2445
   End
   Begin VB.TextBox txtFirstName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Top             =   1350
      Width           =   2445
   End
   Begin VB.TextBox txtEmployeeID 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   975
      Width           =   2055
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   255
      TabIndex        =   13
      Top             =   3645
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   53
   End
   Begin VB.ComboBox cmbPosition 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEmployeesAE.frx":058A
      Left            =   1350
      List            =   "frmEmployeesAE.frx":0597
      TabIndex        =   4
      Top             =   2820
      Width           =   2400
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   5070
      TabIndex        =   14
      Top             =   3705
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
      MICON           =   "frmEmployeesAE.frx":05B7
      PICN            =   "frmEmployeesAE.frx":05D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   3765
      TabIndex        =   15
      Top             =   3690
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
      MICON           =   "frmEmployeesAE.frx":0B6D
      PICN            =   "frmEmployeesAE.frx":0B89
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Please provide all the information needed. Add / Edit Employee Records."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   945
      TabIndex        =   17
      Top             =   240
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   210
      Picture         =   "frmEmployeesAE.frx":1123
      Top             =   45
      Width           =   720
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   0
      TabIndex        =   16
      Top             =   -15
      Width           =   6975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   225
      TabIndex        =   12
      Top             =   2835
      Width           =   1050
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   11
      Top             =   3180
      Width           =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   10
      Top             =   2475
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MiddleName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   225
      TabIndex        =   9
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LastName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   8
      Top             =   1695
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FirstName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   195
      TabIndex        =   7
      Top             =   1305
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   210
      TabIndex        =   6
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "frmEmployeesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub intCount()
 Dim count As Integer
 
 execSQL "Select * From employees"
 count = 999
 
 With RS
 If .RecordCount < 0 Then
 
    txtEmployeeID.Text = "1000"
    Else
     count = count + .RecordCount + 1
     txtEmployeeID.Text = count
     
 End If
 End With
 

Set RS = Nothing
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
  'On Error GoTo err:
If Adding = True Then

 execSQL "select * from employees"
With RS
.AddNew
.Fields!emp_id = txtEmployeeID.Text
.Fields!firstname = txtFirstName.Text
.Fields!lastname = txtLastName.Text
.Fields!middlename = txtMiddleName.Text
.Fields!contactno = txtContactNo.Text
.Fields!Position = cmbPosition.Text
.Fields!address = txtAddress.Text
.Update
.Requery
End With
  MsgBox "New record has been successfully saved.", vbOKOnly + vbInformation

If MsgBox("Do you want to add another new record?", vbYesNo + vbInformation) = vbNo Then
   
    Else
   ClearText (frmEmployeesAE)
   cmbPosition.Text = ""
   txtEmployeeID.SetFocus
End If



ElseIf Editing = True Then
  execSQL "select * from employees where emp_id=" & txtEmployeeID.Text & ""
  
  With RS
   .Fields!firstname = txtFirstName.Text
   .Fields!lastname = txtLastName.Text
   .Fields!middlename = txtMiddleName.Text
   .Fields!contactno = txtContactNo.Text
   .Fields!Position = cmbPosition.Text
   .Fields!address = txtAddress.Text
   .Update
   .Requery
  End With
  
  MsgBox "Record changes has been successfully saved.", vbInformation
 Unload Me

End If



Set RS = Nothing
frmEmployees.FillListView
intCount
 Exit Sub
 
'err:
 'MsgBox "Error # " & err.Number & "Description: " & err.Description, vbExclamation
 
End Sub

Private Sub Form_Load()
If Adding = True Then
 intCount
End If

End Sub
