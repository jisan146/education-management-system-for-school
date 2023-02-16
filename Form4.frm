VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Dues Information"
   ClientHeight    =   6420
   ClientLeft      =   780
   ClientTop       =   495
   ClientWidth     =   8055
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   6420
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2400
      TabIndex        =   1
      Text            =   "Select Class "
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Dues Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form2.Show
End Sub
Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form4.Hide
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
End Sub
Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0 'Baby
Form4.Hide
Form20.Show
Case 1 'One
Form4.Hide
Form21.Show
Case 2 'Two
Form4.Hide
Form22.Show
Case 3 'Three
Form4.Hide
Form23.Show
Case 4 'Four
Form4.Hide
Form24.Show
Case 5 'Five
Form4.Hide
Form25.Show
Case 6 'SIX
Form4.Hide
Form26.Show
Case 7 'Seven
Form4.Hide
Form27.Show
Case 8 'Eight
Form4.Hide
Form28.Show
Case 9 'Nine(Science)
Form4.Hide
Form29.Show
Case 10 'Nine(Business)
Form4.Hide
Form30.Show
Case 11 'Nine(Human)
Form4.Hide
Form31.Show
Case 12 'Ten(Science)
Form4.Hide
Form32.Show
Case 13 'Ten(Business)
Form4.Hide
Form33.Show
Case 14 'Ten(Human)
Form4.Hide
Form34.Show




End Select
End Sub
