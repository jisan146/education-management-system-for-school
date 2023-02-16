VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Student's Information"
   ClientHeight    =   7320
   ClientLeft      =   1260
   ClientTop       =   495
   ClientWidth     =   8685
   LinkTopic       =   "Form3"
   ScaleHeight     =   7320
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7575
      Left            =   -360
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   7515
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   -240
      Width           =   9015
      Begin VB.CommandButton Command3 
         Caption         =   "Main Menu"
         Height          =   495
         Left            =   7440
         TabIndex        =   5
         Top             =   6120
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
         Height          =   495
         Left            =   7440
         TabIndex        =   4
         Top             =   6720
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
         Height          =   495
         Left            =   7440
         TabIndex        =   3
         Top             =   5520
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
         Left            =   3120
         TabIndex        =   2
         Text            =   "Select Class"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Student's Information"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   480
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form3"
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
Form3.Hide
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
Form3.Hide
Form5.Show
Case 1 'One
Form3.Hide
Form6.Show
Case 2 'Two
Form3.Hide
Form7.Show
Case 3 'Three
Form3.Hide
Form8.Show
Case 4 'Four
Form3.Hide
Form9.Show
Case 5 'Five
Form3.Hide
Form10.Show
Case 6 'Six
Form3.Hide
Form11.Show
Case 7 'Seven
Form3.Hide
Form12.Show
Case 8 'Eight
Form3.Hide
Form13.Show
Case 9 'Nine(Science)
Form3.Hide
Form14.Show
Case 10 'Nine(Business)
Form3.Hide
Form15.Show
Case 11 'Nine(Human)
Form3.Hide
Form16.Show
Case 12 'Ten(Science)
Form3.Hide
Form17.Show
Case 13 'Ten(Business)
Form3.Hide
Form18.Show
Case 14 'Ten(Human)
Form3.Hide
Form19.Show
End Select
End Sub
