VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Please Select Information Type"
   ClientHeight    =   7320
   ClientLeft      =   1260
   ClientTop       =   495
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   7335
      Left            =   -360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   7275
      ScaleWidth      =   8955
      TabIndex        =   0
      Top             =   0
      Width           =   9015
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
         Left            =   7560
         TabIndex        =   4
         Top             =   6480
         Width           =   1215
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
         Left            =   4080
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
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
         Left            =   3240
         TabIndex        =   2
         Text            =   "Information of"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Please Select Information Type"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   4935
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form1.Show
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
Combo1.AddItem "Students"
Combo1.AddItem "Dues"
Combo1.AddItem "Teachers"
Combo1.AddItem "Results"
End Sub
Private Sub Combo1_Click()
Select Case Combo1.ListIndex
Case 0 'Students
Form2.Hide
Form3.Show
Case 1 'Payments
Form2.Hide
Form4.Show
Case 2 'Teachers
Form2.Hide
Form35.Show
Case 3 'Results
Form2.Hide
Form36.Show
End Select
End Sub
