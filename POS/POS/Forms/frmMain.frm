VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "COC28D~1.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Point of Sale System [Version 1.0.0.] POS"
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9645
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":058A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Align           =   3  'Align Left
      Height          =   6360
      Left            =   0
      TabIndex        =   3
      Top             =   210
      Width           =   2805
      _Version        =   851970
      _ExtentX        =   4948
      _ExtentY        =   11218
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   6570
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1411
            MinWidth        =   1411
            Picture         =   "frmMain.frx":10CB0
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "6/1/2010"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2117
            MinWidth        =   2117
            Picture         =   "frmMain.frx":1124A
            Text            =   "Username:"
            TextSave        =   "Username:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1773
            MinWidth        =   1764
            Picture         =   "frmMain.frx":117E4
            Text            =   "Time:"
            TextSave        =   "Time:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:43 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   2925
      Top             =   315
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12318
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":133E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13980
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":144B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14FE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   9645
      TabIndex        =   0
      Top             =   0
      Width           =   9645
      Begin CtrlLine.ctrlLiner ctrlLiner1 
         Height          =   30
         Left            =   -60
         TabIndex        =   1
         Top             =   165
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   53
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuWindowsExplorer 
         Caption         =   "Windows Explorer"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReporto 
         Caption         =   "All Sales Transactions Reports"
      End
      Begin VB.Menu mnuReportoo 
         Caption         =   "By Date Sales Transactions Reports"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutSystem 
         Caption         =   "About the system"
      End
      Begin VB.Menu mnuTutorials 
         Caption         =   "Tutorials"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const ID_TASKITEM_CATEGORIES = 1
Const ID_TASKITEM_PRODUCTS = 2
Const ID_TASKITEM_SALES = 3
Const ID_TASKITEM_USERS = 4
Const ID_TASKITEM_PREVILEGES = 5
Const ID_TASKITEM_REPORTS = 6
Const ID_TASKITEM_EMPLOYEES = 7
Const ID_TASKITEM_SYSINFO = 8
Const ID_TASKITEM_DATE = 9
Const ID_TASKITEM_LOGOUT = 10


' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long




Private Sub MDIForm_Load()
CreateTaskPanel

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Are you sure you want to terminate this application?.Click yes to continue.", vbYesNo + vbExclamation) = vbYes Then

 Else
 Cancel = 1
 
 End If
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 1000
        If Me.Height < 4500 Then Me.Height = 4500
        
     ctrlLiner1.Width = Me.Width
         
    End If
    
    

End Sub

Sub CreateTaskPanel()
Dim Group As TaskPanelGroup
Dim Item As TaskPanelItem

Set Group = wndTaskPanel.Groups.Add(0, "View Transactions")
   Group.Tooltip = "View all transactions"
   Group.Special = True
   
   
Group.Items.Add ID_TASKITEM_CATEGORIES, "Category Records", xtpTaskItemTypeLink, 1
Group.Items.Add ID_TASKITEM_PRODUCTS, "Product Records", xtpTaskItemTypeLink, 2
Group.Items.Add ID_TASKITEM_EMPLOYEES, "Employee Records", xtpTaskItemTypeLink, 3
Group.Items.Add ID_TASKITEM_SALES, "Sales Transaction Records", xtpTaskItemTypeLink, 4
Group.Items.Add ID_TASKITEM_REPORTS, "Sales Reports", xtpTaskItemTypeLink, 10


Set Group = wndTaskPanel.Groups.Add(1, "System Settings")
    Group.Tooltip = "View all system settings"
    Group.Special = True
    
    
Group.Items.Add ID_TASKITEM_USERS, "User Settings", xtpTaskItemTypeLink, 5
Group.Items.Add ID_TASKITEM_SYSINFO, "System Information", xtpTaskItemTypeLink, 6
Group.Items.Add ID_TASKITEM_DATE, "Date and Time Settings", xtpTaskItemTypeLink, 7
Group.Items.Add ID_TASKITEM_LOGOUT, "Logout User", xtpTaskItemTypeLink, 8
Group.Items.Add ID_TASKITEM_PREVILEGES, "User Privileges", xtpTaskItemTypeLink, 9


wndTaskPanel.SetImageList i16x16


End Sub


Private Sub mnuAboutSystem_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuCalculator_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuExit_Click()
If MsgBox("Are you sure you want to terminate this application?.Click yes to continue.", vbYesNo + vbExclamation) = vbYes Then
   End
Else
  Exit Sub
End If

End Sub

Private Sub mnuLogout_Click()
Set CN = Nothing

frmLogin.Show vbModal
End Sub

Private Sub mnuNotepad_Click()
Shell "notepad.exe", vbNormalFocus
End Sub

Private Sub mnuTutorials_Click()
MsgBox "To have the full feature of this program" & vbCrLf & _
"You can contact me here:" & vbCrLf & _
"Cellphone#: +639128081019" & vbCrLf & _
"Email: woodenspoon_uto@yahoo.com", vbInformation

End Sub

Private Sub mnuWindowsExplorer_Click()
Shell "Explorer.exe", vbNormalFocus
End Sub

Private Sub wndTaskPanel_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)

Select Case Item.Id
  
  Case ID_TASKITEM_CATEGORIES
     CenterForm frmCategoriesMain
      frmCategoriesMain.Show vbModal
  
  Case ID_TASKITEM_PRODUCTS
       frmProducts.WindowState = vbMaximized
       frmProducts.Show
  Case ID_TASKITEM_EMPLOYEES
      
      frmEmployees.WindowState = vbMaximized
      frmEmployees.Show
      
  Case ID_TASKITEM_USERS
      frmUsers.Show vbModal
  Case ID_TASKITEM_SYSINFO
     StartSysInfo
     
  Case ID_TASKITEM_SALES
      frmSales.WindowState = vbMaximized
      frmSales.Show
      
  Case ID_TASKITEM_LOGOUT
      Set CN = Nothing
      frmLogin.Show vbModal
  Case ID_TASKITEM_PREVILEGES
  If PrevUpdate(frmMain.StatusBar1.Panels(4).Text) = True Then: Exit Sub
    frmPrivileges.Show vbModal
    
  Case ID_TASKITEM_REPORTS
    frmSalesReport.WindowState = vbMaximized
    frmSalesReport.Show
  End Select
  
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Shell SysInfoPath, vbNormalFocus
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

