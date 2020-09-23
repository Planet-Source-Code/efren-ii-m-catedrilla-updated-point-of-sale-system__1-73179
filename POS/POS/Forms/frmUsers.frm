VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   Caption         =   "User List"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5445
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   885
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   53
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   6180
      Top             =   210
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
            Picture         =   "frmUsers.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   3645
      Left            =   30
      TabIndex        =   3
      Top             =   975
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   6429
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User Type"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Employee No."
         Object.Width           =   3528
      EndProperty
   End
   Begin osenxpsuite.OsenXPButton btnAdd 
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   4665
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
      MICON           =   "frmUsers.frx":0B24
      PICN            =   "frmUsers.frx":0B40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnEdit 
      Height          =   375
      Left            =   1665
      TabIndex        =   6
      Top             =   4665
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
      MICON           =   "frmUsers.frx":10DA
      PICN            =   "frmUsers.frx":10F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnDelete 
      Height          =   375
      Left            =   45
      TabIndex        =   7
      Top             =   4665
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
      MICON           =   "frmUsers.frx":1690
      PICN            =   "frmUsers.frx":16AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   5355
      TabIndex        =   8
      Top             =   4665
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
      MICON           =   "frmUsers.frx":1C46
      PICN            =   "frmUsers.frx":1C62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   15
      Picture         =   "frmUsers.frx":21FC
      Top             =   5130
      Width           =   240
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "You can also double click the selected record to edit a record."
      Height          =   180
      Left            =   330
      TabIndex        =   4
      Top             =   5145
      Width           =   4890
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Users List."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   2
      Top             =   315
      Width           =   1350
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   30
      Picture         =   "frmUsers.frx":2786
      Top             =   45
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   8895
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()

Adding = True
Editing = False

frmUsersAE.Caption = "Add new records"
frmUsersAE.Show vbModal
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub




Private Sub btnEdit_Click()
If PrevUpdate(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
    
Editing = True
Adding = False

frmUsersAE.Caption = "Edit records"
With frmUsersAE
.txtUserNo.Text = lvUsers.SelectedItem.Text
.txtUserName.Text = lvUsers.SelectedItem.SubItems(1)
.cmbUserType.Text = lvUsers.SelectedItem.SubItems(2)
.lblEmployees.Caption = lvUsers.SelectedItem.SubItems(3)
End With
load_Employees
frmUsersAE.Show vbModal
End Sub


Private Sub load_Employees()
  On Error GoTo err:
execSQL "select * from employees,users where employees.emp_id=users.emp_id"

With RS
frmUsersAE.txtEmployees.Text = .Fields!lastname & ", " & .Fields!firstname
frmUsersAE.txtPassword.Text = DeCode(.Fields!password)
End With

Set RS = Nothing

 Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
FillListView
End Sub


Public Sub FillListView()
Dim x, tot
    On Error GoTo err:
execSQL "select * from users"
                                                                         
lvUsers.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvUsers.ListItems.Add(, , .Fields!user_id, , 1)
                x.SubItems(1) = .Fields!username
                x.SubItems(2) = .Fields!usertype
                x.SubItems(3) = .Fields!emp_id
                .MoveNext
                
        Wend
    End With
    

  Set RS = Nothing
  Exit Sub
err:
    MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
  
End Sub



Private Sub lvUsers_DblClick()
btnEdit_Click
End Sub
