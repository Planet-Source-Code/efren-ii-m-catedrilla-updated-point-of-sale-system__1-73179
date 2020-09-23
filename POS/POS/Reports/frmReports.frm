VERSION 5.00
Object = "{B4CAD72F-A7F6-4387-A9E0-12699C4AEE04}#8.0#0"; "osenxpsuite.ocx"
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSalesReport 
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   8115
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   8115
      TabIndex        =   5
      Top             =   6765
      Width           =   8115
      Begin osenxpsuite.OsenXPButton OsenXPButton1 
         Height          =   435
         Left            =   60
         TabIndex        =   7
         Top             =   135
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   767
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Query Date"
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
         MICON           =   "frmReports.frx":058A
         PICN            =   "frmReports.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   53
      End
      Begin osenxpsuite.OsenXPButton btnPrint 
         Height          =   435
         Left            =   1755
         TabIndex        =   8
         Top             =   135
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   767
         BCOL            =   15593969
         BCOLO           =   15593969
         TX              =   "Print Sales"
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
         MICON           =   "frmReports.frx":0B40
         PICN            =   "frmReports.frx":0B5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         XPBlendPicture  =   0   'False
      End
   End
   Begin MSComctlLib.ListView lvReports 
      Height          =   5745
      Left            =   30
      TabIndex        =   4
      Top             =   1005
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   10134
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
      NumItems        =   6
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
         Text            =   "Ordate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SRP\Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Amount"
         Object.Width           =   2822
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   15
      TabIndex        =   1
      Top             =   945
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   53
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   7515
      Top             =   390
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
            Picture         =   "frmReports.frx":10F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Print your daily, monthly and yearly sales. To keep track on your sales"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   510
      Width           =   5160
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Reports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   795
      TabIndex        =   2
      Top             =   165
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   120
      Picture         =   "frmReports.frx":1690
      Top             =   75
      Width           =   660
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPrint_Click()
  execSQL "select sales_id,productname,srp,ordate,emp_id,qty,(srp * qty) as amount from sales"
   Set rptAllSales.DataSource = RS
   rptAllSales.WindowState = vbMaximized
   rptAllSales.Show
   Set RS = Nothing
   
End Sub

Private Sub Form_Load()
FillListView
End Sub

Private Sub Form_Resize()
Label1.Width = Me.Width
ctrlLiner1.Width = Me.Width
lvReports.Width = Me.ScaleWidth
lvReports.Height = (Me.ScaleHeight - Picture1.Height) - lvReports.Top

End Sub


Private Sub FillListView()
  Dim x
  execSQL "select sales_id,productname,srp,ordate,emp_id,qty,(srp * qty) as amount from sales"
   
   lvReports.ListItems.Clear
   
  With RS
   While Not .EOF
    Set x = lvReports.ListItems.Add(, , .Fields!sales_id, , 1)
          x.SubItems(1) = .Fields!ProductName
          x.SubItems(2) = .Fields!ordate
          x.SubItems(3) = .Fields!srp
          x.SubItems(4) = .Fields!qty
          x.SubItems(5) = Format(.Fields!amount, "#,##0.00")
     .MoveNext
   Wend
  End With
  
  Set RS = Nothing
  
End Sub
