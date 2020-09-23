VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmProductsAE 
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductsAE.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5100
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   225
      TabIndex        =   18
      Top             =   4560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   53
   End
   Begin VB.TextBox txtProductName 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   1830
      Width           =   3825
   End
   Begin VB.TextBox txtRemarks 
      Height          =   930
      Left            =   1560
      TabIndex        =   5
      Top             =   3330
      Width           =   4545
   End
   Begin VB.TextBox txtSRP 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2955
      Width           =   1590
   End
   Begin VB.TextBox txtUnitPrice 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2595
      Width           =   1605
   End
   Begin VB.TextBox txtCategory 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2190
      Width           =   2970
   End
   Begin VB.TextBox txtProductCode 
      Height          =   315
      Left            =   1545
      TabIndex        =   0
      Top             =   1455
      Width           =   1935
   End
   Begin VB.TextBox txtProductNo 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1095
      Width           =   1605
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   6
      Top             =   810
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   53
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   5190
      TabIndex        =   19
      Top             =   4620
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
      MICON           =   "frmProductsAE.frx":058A
      PICN            =   "frmProductsAE.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   3870
      TabIndex        =   20
      Top             =   4605
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
      MICON           =   "frmProductsAE.frx":0B40
      PICN            =   "frmProductsAE.frx":0B5C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnCategory 
      Height          =   315
      Left            =   4545
      TabIndex        =   21
      ToolTipText     =   "Select Category"
      Top             =   2205
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
      MICON           =   "frmProductsAE.frx":10F6
      PICN            =   "frmProductsAE.frx":1112
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label lblCategoryID 
      Caption         =   "Label10"
      Height          =   345
      Left            =   5010
      TabIndex        =   22
      Top             =   2220
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label as 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   345
      Left            =   270
      TabIndex        =   17
      Top             =   1875
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   345
      Left            =   285
      TabIndex        =   16
      Top             =   3345
      Width           =   1470
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SRP"
      Height          =   345
      Left            =   285
      TabIndex        =   15
      Top             =   2985
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
      Height          =   345
      Left            =   285
      TabIndex        =   14
      Top             =   2640
      Width           =   1470
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   345
      Left            =   285
      TabIndex        =   13
      Top             =   2265
      Width           =   1470
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Code."
      Height          =   345
      Left            =   270
      TabIndex        =   12
      Top             =   1500
      Width           =   1470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product No."
      Height          =   345
      Left            =   285
      TabIndex        =   10
      Top             =   1110
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   135
      Picture         =   "frmProductsAE.frx":16AC
      Top             =   45
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Please provide all the information needed. Add / Edit Product Records."
      Height          =   450
      Left            =   765
      TabIndex        =   9
      Top             =   120
      Width           =   3105
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   0
      TabIndex        =   8
      Top             =   -15
      Width           =   7815
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   "Please provide all the information needed. Add / Edit Category Records."
      Height          =   450
      Left            =   960
      TabIndex        =   7
      Top             =   255
      Width           =   3105
   End
End
Attribute VB_Name = "frmProductsAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnCategory_Click()
frmCategories.Show vbModal
End Sub

Private Sub intCount()
  On Error GoTo err:
  
 Dim count As Integer
 
 execSQL "Select * From products"
 count = 999
 
 With RS
 If .RecordCount < 0 Then
 
    txtProductNo.Text = "1000"
    Else
     count = count + .RecordCount + 1
     txtProductNo.Text = count
     
 End If
 End With
 

 Set RS = Nothing
  Exit Sub
  
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
  
End Sub

Private Sub btnSave_Click()
'On Error GoTo err:
If Adding = True Then
execSQL "Select * from products"

With RS
.AddNew
.Fields!pro_id = txtProductNo.Text
.Fields!ProductName = txtProductName.Text
.Fields!cat_id = lblCategoryID.Caption
.Fields!unitprice = txtUnitPrice.Text
.Fields!srp = txtSRP.Text
.Fields!remarks = txtRemarks.Text
.Fields!code = txtProductCode.Text
.Update
.Requery
End With

MsgBox "New record has been successfully saved.", vbInformation

If MsgBox("Do you want to add another new record?", vbYesNo + vbInformation) = vbNo Then
  Unload Me
  Else
  ClearText Me
  txtProductCode.SetFocus
  End If
  
ElseIf Editing = True Then

execSQL "Select * from products where pro_id=" & txtProductNo.Text & ""

With RS
.Fields!pro_id = txtProductNo.Text
.Fields!ProductName = txtProductName.Text
.Fields!cat_id = lblCategoryID.Caption
.Fields!unitprice = txtUnitPrice.Text
.Fields!srp = txtSRP.Text
.Fields!remarks = txtRemarks.Text
.Fields!code = txtProductCode.Text
.Update
.Requery
End With
 MsgBox "Record changes has been successfully saved.", vbInformation
 Unload Me
End If


Set RS = Nothing
frmProducts.FillListView
intCount
Exit Sub
'err:
 ' MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub

Private Sub Form_Load()
If Adding = True Then
  intCount
  
Else

End If

End Sub
