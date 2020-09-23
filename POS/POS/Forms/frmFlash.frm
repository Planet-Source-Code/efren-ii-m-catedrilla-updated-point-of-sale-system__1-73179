VERSION 5.00
Begin VB.Form frmFlash 
   BorderStyle     =   0  'None
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   LinkTopic       =   "Form2"
   Picture         =   "frmFlash.frx":0000
   ScaleHeight     =   4980
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub
