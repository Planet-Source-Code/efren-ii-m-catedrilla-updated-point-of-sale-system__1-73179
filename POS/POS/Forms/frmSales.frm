VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSales 
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   10455
   Begin VB.PictureBox picToolbar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   10455
      TabIndex        =   3
      Top             =   7905
      Width           =   10455
      Begin osenxpsuite.OsenXPButton btnSales 
         Height          =   765
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1349
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Transaction"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSales.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   0
         TabIndex        =   4
         Top             =   135
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   53
      End
      Begin osenxpsuite.OsenXPButton btnRefresh 
         Height          =   765
         Left            =   1560
         TabIndex        =   6
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1349
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSales.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
      End
      Begin osenxpsuite.OsenXPButton btnQuery 
         Height          =   765
         Left            =   3030
         TabIndex        =   7
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1349
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Query"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSales.frx":05C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
      End
      Begin osenxpsuite.OsenXPButton btnReports 
         Height          =   765
         Left            =   4485
         TabIndex        =   8
         Top             =   270
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1349
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Reports"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSales.frx":05DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
      End
      Begin osenxpsuite.OsenXPButton btnDelete 
         Height          =   765
         Left            =   9030
         TabIndex        =   9
         Top             =   285
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1349
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmSales.frx":05FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         OffsetLeft      =   0
         OffsetTop       =   0
         XPBlendPicture  =   0   'False
      End
   End
   Begin MSComctlLib.ListView lvSales 
      Height          =   6120
      Left            =   0
      TabIndex        =   2
      Top             =   1275
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   10795
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sales No."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SRP\Unit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Employee"
         Object.Width           =   8819
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   900
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   53
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   5190
      Top             =   120
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
            Picture         =   "frmSales.frx":0616
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To edit your sales transactions you can double click the selected record."
      Height          =   255
      Left            =   405
      TabIndex        =   12
      Top             =   990
      Width           =   5820
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   105
      Picture         =   "frmSales.frx":0BB0
      Top             =   975
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All your transactions will be recorded here."
      Height          =   300
      Left            =   1890
      TabIndex        =   11
      Top             =   525
      Width           =   3120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES TRANSACTIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   150
      Width           =   3330
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   105
      Picture         =   "frmSales.frx":113A
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDelete_Click()
 On Error GoTo err:
If PrevDelete(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
 If MsgBox("Are you sure you want to delete the selected record?.Please click yes to continue." & vbCrLf & vbCrLf & "WARNING: You cannot undo this process.", vbYesNo + vbExclamation) = vbNo Then
 Exit Sub
 Else
    execSQL "select * from sales where sales_id=" & lvSales.SelectedItem.Text & ""
With RS
.Delete
.Requery
End With
MsgBox "Record has been successfully deleted.", vbInformation

Set RS = Nothing
End If
FillListView
  Exit Sub
err:
  MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation

End Sub



Private Sub btnRefresh_Click()
FillListView
End Sub



Private Sub btnSales_Click()
Adding = True
Editing = False

frmSalesAE.txtEmployees.Text = frmMain.StatusBar1.Panels(4).Text


frmSalesAE.Show vbModal
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)

Select Case KeyAscii = KeyAscii

Case 20
frmSales.btnSales_Click
Case 6
frmSales.btnRefresh_Click
Case 4
frmSales.btnDelete_Click

End Select

End Sub

Private Sub Form_Load()
FillListView
End Sub

Private Sub Form_Resize()
lvSales.Width = Me.ScaleWidth
lvSales.Height = (Me.ScaleHeight - picToolbar.Height) - lvSales.Top
Label1.Width = Me.Width
ctrlLiner1.Width = Me.Width
ctrlLiner2.Width = Me.Width
btnDelete.Left = Me.ScaleWidth - 1500
End Sub

Public Sub FillListView()
 On Error GoTo err:
 Dim x, tot
 execSQL "select sales.sales_id,sales.emp_id,sales.productname,sales.srp,sales.ordate,sales.qty,(sales.srp * sales.qty) as total from sales"
 lvSales.ListItems.Clear
 With RS
 While Not .EOF
 Set x = lvSales.ListItems.Add(, , .Fields!sales_id, , 1)
         x.SubItems(1) = .Fields!ProductName
         x.SubItems(2) = .Fields!srp
         x.SubItems(3) = .Fields!qty
         x.SubItems(4) = Format(.Fields!total, "#,##0.00")
         x.SubItems(5) = .Fields!ordate
         x.SubItems(6) = .Fields!emp_id
     .MoveNext
     
 Wend
 End With
 

 Set RS = Nothing
  Exit Sub
err:
 MsgBox "Error # " & err.Number & " Description: " & err.Description, vbExclamation
 
End Sub



Private Sub lvSales_DblClick()
  On Error GoTo err:
If PrevUpdate(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub


Editing = True
Adding = False
  With frmSalesAE
  .txtEmployees.Text = lvSales.SelectedItem.SubItems(6)
  .lblSalesID.Caption = lvSales.SelectedItem.Text
  .Combo1.Text = lvSales.SelectedItem.SubItems(1)
  .txtSRP.Text = lvSales.SelectedItem.SubItems(2)
  .txtQty.Text = lvSales.SelectedItem.SubItems(3)
  .txtAmount.Text = lvSales.SelectedItem.SubItems(4)
  .dtDate.Value = lvSales.SelectedItem.SubItems(5)
  .Show vbModal
  End With
  Exit Sub
err:
MsgBox " Error # " & err.Number & " Description: " & err.Description, vbExclamation
End Sub
