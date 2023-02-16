VERSION 5.00
Begin VB.Form Form00 
   Caption         =   "Help"
   ClientHeight    =   5955
   ClientLeft      =   300
   ClientTop       =   495
   ClientWidth     =   8265
   LinkTopic       =   "Form37"
   Picture         =   "Form00.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form00.Hide
Form1.Show
End Sub
