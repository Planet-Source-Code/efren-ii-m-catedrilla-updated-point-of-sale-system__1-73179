VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesAE 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10230
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3390
      TabIndex        =   1
      Top             =   630
      Width           =   2130
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   9105
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   600
      Width           =   945
   End
   Begin VB.TextBox txtQty 
      Height          =   315
      Left            =   7605
      TabIndex        =   2
      Text            =   "0"
      Top             =   615
      Width           =   435
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   285
      Left            =   780
      TabIndex        =   0
      Top             =   660
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      Format          =   50855937
      CurrentDate     =   40303
   End
   Begin VB.TextBox txtSRP 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   630
      Width           =   975
   End
   Begin VB.TextBox txtEmployees 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   135
      Width           =   1755
   End
   Begin osenxpsuite.OsenXPButton btnCancel 
      Height          =   375
      Left            =   8715
      TabIndex        =   12
      Top             =   1080
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frmSalesAE.frx":0000
      PICN            =   "frmSalesAE.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin osenxpsuite.OsenXPButton btnSave 
      Height          =   375
      Left            =   7260
      TabIndex        =   13
      Top             =   1080
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frmSalesAE.frx":05B6
      PICN            =   "frmSalesAE.frx":05D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      XPBlendPicture  =   0   'False
   End
   Begin VB.Label lblSalesID 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   270
      Left            =   3255
      TabIndex        =   14
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      Height          =   345
      Left            =   8115
      TabIndex        =   10
      Top             =   630
      Width           =   990
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      Height          =   345
      Left            =   7305
      TabIndex        =   9
      Top             =   615
      Width           =   540
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "OR Date"
      Height          =   360
      Left            =   135
      TabIndex        =   8
      Top             =   675
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SRP\Unit"
      Height          =   345
      Left            =   5640
      TabIndex        =   6
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name"
      Height          =   360
      Left            =   2310
      TabIndex        =   5
      Top             =   660
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp No."
      Height          =   345
      Left            =   150
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmSalesAE"
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
execSQL "select * from sales"

With RS
.AddNew
.Fields!ProductName = Combo1.Text
.Fields!srp = txtSRP.Text
.Fields!ordate = Format(dtDate.Value, "yyyy-mm-dd")
.Fields!emp_id = txtEmployees.Text
.Fields!qty = txtQty.Text
.Update
.Requery
End With
MsgBox "New record has been successfully saved.", vbInformation

If MsgBox("Do you want to add another new record?", vbYesNo + vbInformation) = vbNo Then
  Unload Me
  Else
 dtDate.Value = Date
 Combo1.Text = ""
 txtSRP.Text = "0.00"
 txtQty.Text = "0"
 txtAmount.Text = "0.00"
End If
ElseIf Editing = True Then
execSQL "select * from sales where sales_id=" & lblSalesID.Caption & ""

With RS
.Fields!ProductName = Combo1.Text
.Fields!srp = txtSRP.Text
.Fields!ordate = Format(dtDate.Value, "yyyy-mm-dd")
.Fields!emp_id = txtEmployees.Text
.Fields!qty = txtQty.Text
.Update
.Requery
End With
MsgBox "Record changes has been successfully saved.", vbInformation
Unload Me
End If

Set RS = Nothing

frmSales.FillListView

  Exit Sub
err:
 MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtQty.SetFocus
load_SRP
End If

KeyAscii = Autocomplete(Combo1, KeyAscii)

End Sub

Private Sub Combo1_LostFocus()

Proper_Case Combo1

End Sub

Private Sub Form_Load()
CenterForm Me
load_Products
End Sub


Private Sub load_Products()

execSQL "select * from products"
 
 With RS
    While Not .EOF
     Combo1.AddItem .Fields!ProductName
     .MoveNext
    Wend
    
 End With

Set RS = Nothing
End Sub

Private Sub load_SRP()

execSQL "select * from products where productname='" & Combo1.Text & "'"
 
 With RS
    txtSRP.Text = .Fields!srp
 End With

Set RS = Nothing
End Sub


Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 txtAmount.Text = Format(txtSRP.Text * txtQty.Text, "#,##0.00")
 End If
 
End Sub
