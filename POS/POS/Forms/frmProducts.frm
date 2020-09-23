VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProducts 
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   11100
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   705
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   53
   End
   Begin MSComctlLib.ListView lvProducts 
      Height          =   6525
      Left            =   75
      TabIndex        =   5
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
         Text            =   "Product No."
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Unit Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SRP/Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Category"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Remarks"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      ItemData        =   "frmProducts.frx":058A
      Left            =   1155
      List            =   "frmProducts.frx":0597
      TabIndex        =   1
      Top             =   1035
      Width           =   2160
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   3735
      TabIndex        =   0
      Top             =   1050
      Width           =   2370
   End
   Begin osenxpsuite.OsenXPButton btnSearch 
      Height          =   315
      Left            =   6135
      TabIndex        =   3
      ToolTipText     =   "Select Category"
      Top             =   1050
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
      MICON           =   "frmProducts.frx":05C1
      PICN            =   "frmProducts.frx":05DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnAdd 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   8160
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
      MICON           =   "frmProducts.frx":0B77
      PICN            =   "frmProducts.frx":0B93
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnEdit 
      Height          =   375
      Left            =   5310
      TabIndex        =   7
      Top             =   8160
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
      MICON           =   "frmProducts.frx":112D
      PICN            =   "frmProducts.frx":1149
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnDelete 
      Height          =   375
      Left            =   3690
      TabIndex        =   8
      Top             =   8160
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
      MICON           =   "frmProducts.frx":16E3
      PICN            =   "frmProducts.frx":16FF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnClose 
      Height          =   375
      Left            =   8430
      TabIndex        =   9
      Top             =   8160
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
      MICON           =   "frmProducts.frx":1C99
      PICN            =   "frmProducts.frx":1CB5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnRefresh 
      Height          =   375
      Left            =   8085
      TabIndex        =   10
      Top             =   1035
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
      MICON           =   "frmProducts.frx":224F
      PICN            =   "frmProducts.frx":226B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   9885
      Top             =   945
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
            Picture         =   "frmProducts.frx":2805
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProducts.frx":2D9F
      Height          =   450
      Left            =   705
      TabIndex        =   14
      Top             =   120
      Width           =   5985
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   90
      Picture         =   "frmProducts.frx":2E31
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   720
      Left            =   -75
      TabIndex        =   12
      Top             =   0
      Width           =   11175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000010&
      Height          =   480
      Left            =   75
      TabIndex        =   11
      Top             =   8115
      Width           =   9750
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
      Left            =   240
      TabIndex        =   2
      Top             =   1035
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3405
      Picture         =   "frmProducts.frx":3577
      Top             =   1050
      Width           =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000010&
      Height          =   510
      Left            =   75
      TabIndex        =   4
      Top             =   975
      Width           =   9720
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()
Adding = True
Editing = False
frmProductsAE.Caption = "Add new record"
frmProductsAE.Show vbModal
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnDelete_Click()
If PrevDelete(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
 
 If MsgBox("Are you sure you want to delete the selected record?.Please click yes to continue." & vbCrLf & vbCrLf & "WARNING: You cannot undo this process.", vbYesNo + vbExclamation) = vbNo Then
 Exit Sub
 Else
    execSQL "select * from products where pro_id=" & lvProducts.SelectedItem.Text & ""
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
On Error GoTo err:

Editing = True
Adding = False

With frmProductsAE
.txtProductNo.Text = lvProducts.SelectedItem.Text
.txtProductCode.Text = lvProducts.SelectedItem.SubItems(1)
.txtProductName.Text = lvProducts.SelectedItem.SubItems(2)
.txtCategory.Text = lvProducts.SelectedItem.SubItems(5)
.txtUnitPrice.Text = lvProducts.SelectedItem.SubItems(3)
.txtSRP.Text = lvProducts.SelectedItem.SubItems(4)
.txtRemarks.Text = lvProducts.SelectedItem.SubItems(6)
End With
load_Category
frmProductsAE.Show vbModal
Exit Sub

err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub btnRefresh_Click()
FillListView
End Sub

Private Sub btnSearch_Click()
Dim x

Select Case cmbSearch.Text

Case "Product Name"
      execSQL "select * from products,categories where categories.cat_id=products.cat_id and productname Like '" & txtSearch.Text & "%'"
      
Case "Product Code"
      execSQL "select * from products,categories where categories.cat_id=products.cat_id and code='" & txtSearch.Text & "'"
      
Case "Category"
     execSQL "select * from products,categories where categories.cat_id=products.cat_id and description='" & txtSearch.Text & "'"
     
End Select

lvProducts.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvProducts.ListItems.Add(, , .Fields!pro_id, , 1)
                x.SubItems(1) = .Fields!code
                x.SubItems(2) = .Fields!ProductName
                x.SubItems(3) = .Fields!unitprice
                x.SubItems(4) = .Fields!srp
                x.SubItems(5) = .Fields!Description
                x.SubItems(6) = .Fields!remarks
                
                .MoveNext
                
        Wend
    End With

Set RS = Nothing
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
execSQL "select * from products,categories where categories.cat_id=products.cat_id"

                                                                         
lvProducts.ListItems.Clear
                                       
With RS
        While Not .EOF
            Set x = lvProducts.ListItems.Add(, , .Fields!pro_id, , 1)
                x.SubItems(1) = .Fields!code
                x.SubItems(2) = .Fields!ProductName
                x.SubItems(3) = .Fields!unitprice
                x.SubItems(4) = .Fields!srp
                x.SubItems(5) = .Fields!Description
                x.SubItems(6) = .Fields!remarks
                
                .MoveNext
                
        Wend
    End With
    

  Set RS = Nothing
  Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub load_Category()

execSQL "select * from categories where description='" & lvProducts.SelectedItem.SubItems(5) & "'"

With RS
frmProductsAE.lblCategoryID.Caption = .Fields!cat_id
End With

Set RS = Nothing
End Sub




Private Sub lvProducts_DblClick()
btnEdit_Click
End Sub
