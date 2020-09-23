VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmployees 
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmployees.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   11220
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   3675
      TabIndex        =   4
      Top             =   990
      Width           =   2370
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      ItemData        =   "frmEmployees.frx":058A
      Left            =   1110
      List            =   "frmEmployees.frx":059D
      TabIndex        =   3
      Top             =   990
      Width           =   2160
   End
   Begin osenxpsuite.OsenXPButton btnSearch 
      Height          =   315
      Left            =   6060
      TabIndex        =   5
      ToolTipText     =   "Select Category"
      Top             =   990
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
      MICON           =   "frmEmployees.frx":05D6
      PICN            =   "frmEmployees.frx":05F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnRefresh 
      Height          =   375
      Left            =   8025
      TabIndex        =   6
      Top             =   975
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Refresh Records"
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
      MICON           =   "frmEmployees.frx":0B8C
      PICN            =   "frmEmployees.frx":0BA8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   9825
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployees.frx":1142
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   53
   End
   Begin MSComctlLib.ListView lvEmployees 
      Height          =   6525
      Left            =   0
      TabIndex        =   9
      Top             =   1530
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   11509
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "First Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Middle Name"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact No."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Position"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Address"
         Object.Width           =   8819
      EndProperty
   End
   Begin osenxpsuite.OsenXPButton btnAdd 
      Height          =   375
      Left            =   6645
      TabIndex        =   10
      Top             =   8175
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Add Records"
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
      MICON           =   "frmEmployees.frx":16DC
      PICN            =   "frmEmployees.frx":16F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnEdit 
      Height          =   375
      Left            =   5235
      TabIndex        =   11
      Top             =   8175
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "&Edit Records"
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
      MICON           =   "frmEmployees.frx":1C92
      PICN            =   "frmEmployees.frx":1CAE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnDelete 
      Height          =   375
      Left            =   3630
      TabIndex        =   12
      Top             =   8190
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Delete Records"
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
      MICON           =   "frmEmployees.frx":2248
      PICN            =   "frmEmployees.frx":2264
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnClose 
      Height          =   375
      Left            =   8355
      TabIndex        =   13
      Top             =   8175
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Close"
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
      MICON           =   "frmEmployees.frx":27FE
      PICN            =   "frmEmployees.frx":281A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Height          =   480
      Left            =   0
      TabIndex        =   14
      Top             =   8130
      Width           =   9750
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3345
      Picture         =   "frmEmployees.frx":2DB4
      Top             =   1005
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   180
      TabIndex        =   7
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000010&
      Height          =   510
      Left            =   -30
      TabIndex        =   8
      Top             =   900
      Width           =   9855
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   75
      Picture         =   "frmEmployees.frx":333E
      Top             =   15
      Width           =   720
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEmployees.frx":3C67
      Height          =   450
      Left            =   795
      TabIndex        =   1
      Top             =   120
      Width           =   5985
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   720
      Left            =   -45
      TabIndex        =   2
      Top             =   -30
      Width           =   11175
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
Adding = True
Editing = False

frmEmployeesAE.Caption = "Add new record"
frmEmployeesAE.Show vbModal

End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnDelete_Click()
If PrevDelete(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
If MsgBox("Are you sure you want to delete the selected record?.Please click yes to continue." & vbCrLf & vbCrLf & "WARNING: You cannot undo this process.", vbYesNo + vbExclamation) = vbNo Then
 Exit Sub
 Else
    execSQL "select * from employees where emp_id=" & lvEmployees.SelectedItem.Text & ""
With RS
.Delete
.Requery
End With
MsgBox "Record has been successfully deleted.", vbInformation

Set RS = Nothing
End If
FillListView
End Sub

Private Sub btnEdit_Click()
If PrevUpdate(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub



Editing = True
Adding = False
frmEmployeesAE.Caption = "Edit records"

With frmEmployeesAE
.txtEmployeeID.Text = lvEmployees.SelectedItem.Text
.txtFirstName.Text = lvEmployees.SelectedItem.SubItems(1)
.txtLastName.Text = lvEmployees.SelectedItem.SubItems(2)
.txtMiddleName.Text = lvEmployees.SelectedItem.SubItems(3)
.txtContactNo.Text = lvEmployees.SelectedItem.SubItems(4)
.cmbPosition.Text = lvEmployees.SelectedItem.SubItems(5)
.txtAddress.Text = lvEmployees.SelectedItem.SubItems(6)
End With
frmEmployeesAE.Show vbModal
End Sub

Private Sub btnRefresh_Click()
FillListView

End Sub



Private Sub btnSearch_Click()

  On Error GoTo err:
Select Case cmbSearch.Text

  Case "First Name"
   
    execSQL "select * from employees where firstname like '" & txtSearch.Text & "%'"
   
    
  Case "Last Name"
   
    execSQL "select * from employees where lastname like '" & txtSearch.Text & "%'"
    
    
  Case "Middle Name"
    execSQL "select * from employees where middlename like '" & txtSearch.Text & "%'"
    
    
  Case "Employee No."
    execSQL "select * from employees where emp_id=" & txtSearch.Text & ""
    
  End Select
  
  lvEmployees.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvEmployees.ListItems.Add(, , .Fields!emp_id, , 1)
                x.SubItems(1) = .Fields!firstname
                x.SubItems(2) = .Fields!lastname
                x.SubItems(3) = .Fields!middlename
                x.SubItems(4) = .Fields!contactno
                x.SubItems(5) = .Fields!Position
                x.SubItems(6) = .Fields!address
                
                
                .MoveNext
                
        Wend
    End With

  Set RS = Nothing
     Exit Sub
     
err:
   MsgBox "Error # " & err.Number & "Description: " & err.Description, vbExclamation
   
End Sub

Private Sub Form_Load()
FillListView
End Sub

Private Sub Form_Resize()
ctrlLiner1.Width = Me.Width
Label4.Width = Me.Width
End Sub

Public Sub FillListView()
Dim x
    On Error GoTo err:

execSQL "select * from employees"

                                                                         
lvEmployees.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvEmployees.ListItems.Add(, , .Fields!emp_id, , 1)
                x.SubItems(1) = .Fields!firstname
                x.SubItems(2) = .Fields!lastname
                x.SubItems(3) = .Fields!middlename
                x.SubItems(4) = .Fields!contactno
                x.SubItems(5) = .Fields!Position
                x.SubItems(6) = .Fields!address
                
                
                .MoveNext
                
        Wend
    End With
    

  Set RS = Nothing
     Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub lvEmployees_DblClick()
btnEdit_Click
End Sub
