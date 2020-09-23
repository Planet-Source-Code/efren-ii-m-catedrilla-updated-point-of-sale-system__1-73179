VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Begin VB.Form frmCategoriesAE 
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategoriesAE.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2190
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   810
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   53
   End
   Begin VB.TextBox txtDescriptions 
      Height          =   315
      Left            =   1455
      TabIndex        =   0
      Top             =   1290
      Width           =   3525
   End
   Begin VB.TextBox txtCategoryID 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   930
      Width           =   2055
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   3945
      TabIndex        =   6
      Top             =   1740
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
      MICON           =   "frmCategoriesAE.frx":058A
      PICN            =   "frmCategoriesAE.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   2610
      TabIndex        =   7
      Top             =   1725
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
      MICON           =   "frmCategoriesAE.frx":0B40
      PICN            =   "frmCategoriesAE.frx":0B5C
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
      Caption         =   "Please provide all the information needed. Add / Edit Category Records."
      Height          =   450
      Left            =   960
      TabIndex        =   5
      Top             =   255
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   210
      Picture         =   "frmCategoriesAE.frx":10F6
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Height          =   825
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category No."
      Height          =   375
      Left            =   75
      TabIndex        =   1
      Top             =   960
      Width           =   1155
   End
End
Attribute VB_Name = "frmCategoriesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub intCount()
 Dim count As Integer
 On Error GoTo err:
 execSQL "Select * From categories"
 count = 999
 
 With RS
 If .RecordCount < 0 Then
 
    txtCategoryID.Text = "1000"
    Else
     count = count + .RecordCount + 1
     txtCategoryID.Text = count
     
 End If
 End With
 

 Set RS = Nothing
 Exit Sub
err:
 MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
 
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If is_Empty(txtDescriptions) = False Then: Exit Sub
'On Error GoTo err:
   
If Adding = True Then
execSQL "select * from categories"

With RS
.AddNew
.Fields!cat_id = txtCategoryID.Text
.Fields!Description = txtDescriptions.Text
.Update
.Requery
End With
 MsgBox "New record has been successfully saved.", vbOKOnly + vbInformation
 
 If MsgBox("Do you want to add another new record?", vbYesNo + vbInformation) = vbNo Then
 Unload Me
 Else
  ClearText frmCategoriesAE
  txtDescriptions.SetFocus
End If

 ElseIf Editing = True Then
 execSQL "select * from categories where cat_id=" & txtCategoryID.Text & ""
 
 With RS
 .Fields!Description = txtDescriptions.Text
 .Update
 .Requery
 End With
 
 MsgBox "Record changes has been successfully saved.", vbInformation
 Unload Me
 End If
 
 Set RS = Nothing
 
intCount
frmCategoriesMain.FillListView
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




