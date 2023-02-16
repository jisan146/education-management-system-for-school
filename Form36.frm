VERSION 5.00
Begin VB.Form Form36 
   Caption         =   "Student's Results"
   ClientHeight    =   6330
   ClientLeft      =   480
   ClientTop       =   495
   ClientWidth     =   8400
   LinkTopic       =   "Form36"
   Picture         =   "Form36.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   1
      Text            =   "Select Class"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Student's Results"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form36.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form36.Hide
Form1.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem "Baby"
Combo1.AddItem "One"
Combo1.AddItem "Two"
Combo1.AddItem "Three"
Combo1.AddItem "Four"
Combo1.AddItem "Five"
Combo1.AddItem "Six"
Combo1.AddItem "Seven"
Combo1.AddItem "Eight"
Combo1.AddItem "Nine(Science)"
Combo1.AddItem "Nine(Business)"
Combo1.AddItem "Nine(Human)"
Combo1.AddItem "Ten(Science)"
Combo1.AddItem "Ten(Business)"
Combo1.AddItem "Ten(Human)"
Combo1.AddItem "Pree Test(Science)"
Combo1.AddItem "Pree Test(Business)"
Combo1.AddItem "Pree Test(Human)"
Combo1.AddItem "Test(Science)"
Combo1.AddItem "Test(Business)"
Combo1.AddItem "Test(Human)"



End Sub
Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0 'Baby
Form4.Hide
Form37.Show
Case 1 'One
Form4.Hide
Form38.Show
Case 2 'Two
Form4.Hide
Form39.Show
Case 3 'Three
Form4.Hide
Form40.Show
Case 4 'Four
Form4.Hide
Form41.Show
Case 5 'Five
Form4.Hide
Form42.Show
Case 6 'SIX
Form4.Hide
Form43.Show
Case 7 'Seven
Form4.Hide
Form44.Show
Case 8 'Eight
Form4.Hide
Form45.Show
Case 9 'Nine(Science)
Form4.Hide
Form46.Show
Case 10 'Nine(Business)
Form4.Hide
Form47.Show
Case 11 'Nine(Human)
Form4.Hide
Form48.Show
Case 12 'Ten(Science)
Form4.Hide
Form49.Show
Case 13 'Ten(Business)
Form4.Hide
Form50.Show
Case 14 'Ten(Human)
Form4.Hide
Form51.Show
Case 15 'Pree Test(Science)
Form4.Hide
Form52.Show
Case 16 'Pree Test(Business)
Form4.Hide
Form53.Show
Case 17 'Pree Test(Human)
Form4.Hide
Form54.Show
Case 18 'Test(Science)
Form4.Hide
Form55.Show
Case 19 'Test(Business)
Form4.Hide
Form56.Show
Case 20 'Test(Human)
Form4.Hide
Form57.Show
End Select
End Sub


