VERSION 5.00
Begin VB.Form Form000 
   BackColor       =   &H80000009&
   Caption         =   "About"
   ClientHeight    =   7800
   ClientLeft      =   360
   ClientTop       =   495
   ClientWidth     =   7620
   LinkTopic       =   "Form37"
   Picture         =   "Form000.frx":0000
   ScaleHeight     =   7800
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Form000.Hide
Form1.Show
End Sub

