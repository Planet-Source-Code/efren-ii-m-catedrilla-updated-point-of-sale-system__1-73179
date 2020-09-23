VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategoriesMain 
   Caption         =   "Category List"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategoriesMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   1125
      TabIndex        =   6
      Top             =   855
      Width           =   2445
   End
   Begin MSComctlLib.ListView lvCategories 
      Height          =   4335
      Left            =   -30
      TabIndex        =   3
      Top             =   1275
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7646
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   9701
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   795
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPButton btnSearch 
      Height          =   315
      Left            =   3585
      TabIndex        =   7
      ToolTipText     =   "Select Category"
      Top             =   855
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
      MICON           =   "frmCategoriesMain.frx":058A
      PICN            =   "frmCategoriesMain.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnAdd 
      Height          =   375
      Left            =   3015
      TabIndex        =   8
      Top             =   5640
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
      MICON           =   "frmCategoriesMain.frx":0B40
      PICN            =   "frmCategoriesMain.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnEdit 
      Height          =   375
      Left            =   1620
      TabIndex        =   9
      Top             =   5640
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
      MICON           =   "frmCategoriesMain.frx":10F6
      PICN            =   "frmCategoriesMain.frx":1112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnDelete 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   5640
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
      MICON           =   "frmCategoriesMain.frx":16AC
      PICN            =   "frmCategoriesMain.frx":16C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   7845
      TabIndex        =   11
      Top             =   5640
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
      MICON           =   "frmCategoriesMain.frx":1C62
      PICN            =   "frmCategoriesMain.frx":1C7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8400
      Top             =   675
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
            Picture         =   "frmCategoriesMain.frx":2218
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "You can also double click the selected record to edit a record."
      Height          =   180
      Left            =   345
      TabIndex        =   12
      Top             =   6075
      Width           =   4890
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   30
      Picture         =   "frmCategoriesMain.frx":27B2
      Top             =   6060
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by:"
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
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   885
      Width           =   990
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   9210
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   60
      Picture         =   "frmCategoriesMain.frx":2D3C
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "All Category List."
      Height          =   405
      Left            =   795
      TabIndex        =   1
      Top             =   180
      Width           =   3105
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   -15
      Width           =   9240
   End
End
Attribute VB_Name = "frmCategoriesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAdd_Click()
Adding = True
Editing = False
frmCategoriesAE.Caption = "Add new records"
frmCategoriesAE.Show vbModal
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub


Public Sub FillListView()
Dim x
   On Error GoTo err:
execSQL "select * from categories"

                                                                         
lvCategories.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvCategories.ListItems.Add(, , .Fields!cat_id, , 1)
                x.SubItems(1) = .Fields!Description
                
                .MoveNext
                
        Wend
    End With
    

  Set RS = Nothing
    Exit Sub
err:
 MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
 
End Sub


Private Sub Search_Categories()
Dim x
                                  
execSQL "select * from categories"

                                                                         
lvCategories.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvCategories.ListItems.Add(, , .Fields!cat_id, , 1)
                x.SubItems(1) = .Fields!Description
                
                .MoveNext
                
        Wend
    End With
    
  
  Set RS = Nothing
End Sub

Private Sub btnDelete_Click()
If PrevDelete(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
 
 If MsgBox("Are you sure you want to delete the selected record?.Please click yes to continue." & vbCrLf & vbCrLf & "WARNING: You cannot undo this process.", vbYesNo + vbExclamation) = vbNo Then
 Exit Sub
 Else
    execSQL "select * from categories where cat_id=" & lvCategories.SelectedItem.Text & ""
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

With frmCategoriesAE
.Caption = "Edit records"
.txtCategoryID.Text = lvCategories.SelectedItem.Text
.txtDescriptions.Text = lvCategories.SelectedItem.SubItems(1)
.Show vbModal
End With

End Sub

Private Sub btnSearch_Click()
Dim x
execSQL "select * from categories where description like '" & Trim(txtSearch.Text) & "%'"
 If RS.RecordCount < 1 Then: MsgBox "Invalid Entry.No record found!", vbExclamation
 
lvCategories.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvCategories.ListItems.Add(, , .Fields!cat_id, , 1)
                x.SubItems(1) = .Fields!Description
                
                .MoveNext
                
        Wend
    End With
    

  Set RS = Nothing

End Sub

Private Sub Form_Load()
FillListView
End Sub



Private Sub lvCategories_DblClick()
 btnEdit_Click
End Sub
