VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategories 
   Caption         =   "Category List"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4575
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList i16x16 
      Left            =   6240
      Top             =   270
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
            Picture         =   "frmCategories.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCategories 
      Height          =   3645
      Left            =   90
      TabIndex        =   1
      Top             =   375
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descriptions"
         Object.Width           =   9172
      EndProperty
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   5595
      TabIndex        =   2
      Top             =   4080
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
      MICON           =   "frmCategories.frx":0B24
      PICN            =   "frmCategories.frx":0B40
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   4275
      TabIndex        =   3
      Top             =   4065
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BCOL            =   15593969
      BCOLO           =   15593969
      TX              =   "Select"
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
      MICON           =   "frmCategories.frx":10DA
      PICN            =   "frmCategories.frx":10F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   " Select a product category in the list."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   6765
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub FillListView()
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
  MsgBox "Error # " & err.Number & " Description: " & err.Description
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
With frmProductsAE
.lblCategoryID.Caption = lvCategories.SelectedItem.Text
.txtCategory.Text = lvCategories.SelectedItem.SubItems(1)
End With
Unload Me
End Sub

Private Sub Form_Load()
FillListView
End Sub
