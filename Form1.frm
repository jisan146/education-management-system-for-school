VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Titash Software"
   ClientHeight    =   6900
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "School Database"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By MD.Tareq Rahman Jisan"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "TITASH GAS HIGH SCHOOL"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   5535
   End
   Begin VB.Menu Infomenu 
      Caption         =   "Information"
      Begin VB.Menu sub1 
         Caption         =   "Student of Class"
         Begin VB.Menu sub1sub1 
            Caption         =   "Baby"
         End
         Begin VB.Menu sub1sub2 
            Caption         =   "One"
         End
         Begin VB.Menu sub1sub3 
            Caption         =   "Two"
         End
         Begin VB.Menu sub1sub4 
            Caption         =   "Three"
         End
         Begin VB.Menu sub1sub5 
            Caption         =   "Four"
         End
         Begin VB.Menu sub1sub6 
            Caption         =   "Five"
         End
         Begin VB.Menu sub1sub7 
            Caption         =   "Six"
         End
         Begin VB.Menu sub1sub8 
            Caption         =   "Seven"
         End
         Begin VB.Menu sub1sub9 
            Caption         =   "Eight"
         End
         Begin VB.Menu sub1sub10 
            Caption         =   "Nine(Science)"
         End
         Begin VB.Menu sub1sub11 
            Caption         =   "Nine(Business)"
         End
         Begin VB.Menu sub1sub12 
            Caption         =   "Nine(Human)"
         End
         Begin VB.Menu sub1sub13 
            Caption         =   "Ten(Science)"
         End
         Begin VB.Menu sub1sub14 
            Caption         =   "Ten(Business)"
         End
         Begin VB.Menu sub1sub15 
            Caption         =   "Ten(Human)"
         End
      End
   End
   Begin VB.Menu sub2 
      Caption         =   "Dues of student"
      Begin VB.Menu sub2sub1 
         Caption         =   "Baby"
      End
      Begin VB.Menu sub2sub2 
         Caption         =   "One "
      End
      Begin VB.Menu sub2sub3 
         Caption         =   "Two"
      End
      Begin VB.Menu sub2sub4 
         Caption         =   "Three"
      End
      Begin VB.Menu sub2sub5 
         Caption         =   "Four"
      End
      Begin VB.Menu sub2sub6 
         Caption         =   "Five"
      End
      Begin VB.Menu sub2sub7 
         Caption         =   "Six"
      End
      Begin VB.Menu sub2sub8 
         Caption         =   "Seven"
      End
      Begin VB.Menu sub2sub9 
         Caption         =   "Eight"
      End
      Begin VB.Menu sub2sub10 
         Caption         =   "Nine(Science)"
      End
      Begin VB.Menu sub2sub11 
         Caption         =   "Nine(Business)"
      End
      Begin VB.Menu sub2sub12 
         Caption         =   "Nine(Human)"
      End
      Begin VB.Menu sub2sub13 
         Caption         =   "Ten(Science)"
      End
      Begin VB.Menu sub2sub14 
         Caption         =   "Ten(Business)"
      End
      Begin VB.Menu sub2sub15 
         Caption         =   "Ten(Human)"
      End
   End
   Begin VB.Menu sub3 
      Caption         =   "Teacher"
   End
   Begin VB.Menu sub4 
      Caption         =   "Results"
      Begin VB.Menu sub4sub1 
         Caption         =   "Baby"
      End
      Begin VB.Menu sub4sub2 
         Caption         =   "One"
      End
      Begin VB.Menu sub4sub3 
         Caption         =   "Two"
      End
      Begin VB.Menu sub4sub4 
         Caption         =   "Three"
      End
      Begin VB.Menu sub4sub5 
         Caption         =   "Four"
      End
      Begin VB.Menu sub4sub6 
         Caption         =   "Five"
      End
      Begin VB.Menu sub4sub7 
         Caption         =   "Six"
      End
      Begin VB.Menu sub4sub8 
         Caption         =   "Seven"
      End
      Begin VB.Menu sub4sub9 
         Caption         =   "Eight"
      End
      Begin VB.Menu sub4sub10 
         Caption         =   "Nine(Science)"
      End
      Begin VB.Menu sub4sub11 
         Caption         =   "Nine(Busniess)"
      End
      Begin VB.Menu sub4sub12 
         Caption         =   "Nine(Human)"
      End
      Begin VB.Menu sub4sub13 
         Caption         =   "Ten(Science)"
      End
      Begin VB.Menu sub4sub14 
         Caption         =   "Ten(Business)"
      End
      Begin VB.Menu sub4sub15 
         Caption         =   "Ten(Human)"
      End
      Begin VB.Menu sub4sub16 
         Caption         =   "Pree-Test(Science)"
      End
      Begin VB.Menu sub4sub17 
         Caption         =   "Pree-Test(Busniess)"
      End
      Begin VB.Menu sub4sub18 
         Caption         =   "Pree-Test(Human)"
      End
      Begin VB.Menu sub4sub19 
         Caption         =   "Test(Science)"
      End
      Begin VB.Menu sub4sub20 
         Caption         =   "Test(Busniess)"
      End
      Begin VB.Menu sub4sub21 
         Caption         =   "Test(Human)"
      End
   End
   Begin VB.Menu sub6 
      Caption         =   "About"
      Begin VB.Menu sub6sub1 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sub6sub2 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu asd 
      Caption         =   "Present"
   End
   Begin VB.Menu sub5 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asd_Click()
Form99.Show
End Sub

Private Sub command1_click()
Form2.Show
End Sub
Private Sub command2_click()
End
End Sub
Private Sub sub1sub1_Click()
Form5.Show
End Sub
Private Sub sub1sub2_Click()
Form6.Show
End Sub
Private Sub sub1sub3_Click()
Form7.Show
End Sub
Private Sub sub1sub4_Click()
Form8.Show
End Sub
Private Sub sub1sub5_Click()
Form9.Show
End Sub
Private Sub sub1sub6_Click()
Form10.Show
End Sub
Private Sub sub1sub7_Click()
Form11.Show
End Sub
Private Sub sub1sub8_Click()
Form12.Show
End Sub
Private Sub sub1sub9_Click()
Form13.Show
End Sub
Private Sub sub1sub10_Click()
Form14.Show
End Sub
Private Sub sub1sub11_Click()
Form15.Show
End Sub
Private Sub sub1sub12_Click()
Form16.Show
End Sub
Private Sub sub1sub13_Click()
Form17.Show
End Sub
Private Sub sub1sub14_Click()
Form18.Show
End Sub
Private Sub sub1sub15_Click()
Form19.Show
End Sub
Private Sub sub2sub1_Click()
Form20.Show
End Sub
Private Sub sub2sub2_Click()
Form21.Show
End Sub
Private Sub sub2sub3_Click()
Form22.Show
End Sub
Private Sub sub2sub4_Click()
Form23.Show
End Sub
Private Sub sub2sub5_Click()
Form24.Show
End Sub
Private Sub sub2sub6_Click()
Form25.Show
End Sub
Private Sub sub2sub7_Click()
Form26.Show
End Sub
Private Sub sub2sub8_Click()
Form27.Show
End Sub
Private Sub sub2sub9_Click()
Form28.Show
End Sub
Private Sub sub2sub10_Click()
Form29.Show
End Sub
Private Sub sub2sub11_Click()
Form30.Show
End Sub
Private Sub sub2sub12_Click()
Form31.Show
End Sub
Private Sub sub2sub13_Click()
Form32.Show
End Sub
Private Sub sub2sub14_Click()
Form33.Show
End Sub
Private Sub sub2sub15_Click()
Form34.Show
End Sub
Private Sub sub3_Click()
Form35.Show
End Sub
Private Sub sub4sub1_Click()
Form37.Show
End Sub
Private Sub sub4sub2_Click()
Form38.Show
End Sub
Private Sub sub4sub3_Click()
Form39.Show
End Sub
Private Sub sub4sub4_Click()
Form40.Show
End Sub
Private Sub sub4sub5_Click()
Form41.Show
End Sub
Private Sub sub4sub6_Click()
Form42.Show
End Sub

Private Sub sub4sub7_Click()
Form43.Show
End Sub
Private Sub sub4sub8_Click()
Form44.Show
End Sub
Private Sub sub4sub9_Click()
Form45.Show
End Sub
Private Sub sub4sub10_Click()
Form46.Show
End Sub
Private Sub sub4sub11_Click()
Form47.Show
End Sub
Private Sub sub4sub12_Click()
Form48.Show
End Sub
Private Sub sub4sub13_Click()
Form49.Show
End Sub
Private Sub sub4sub14_Click()
Form50.Show
End Sub
Private Sub sub4sub15_Click()
Form51.Show
End Sub
Private Sub sub4sub16_Click()
Form52.Show
End Sub
Private Sub sub4sub17_Click()
Form53.Show
End Sub
Private Sub sub4sub18_Click()
Form54.Show
End Sub
Private Sub sub4sub19_Click()
Form55.Show
End Sub
Private Sub sub4sub20_Click()
Form56.Show
End Sub
Private Sub sub4sub21_Click()
Form57.Show
End Sub
Private Sub sub5_Click()
End
End Sub

Private Sub sub6sub1_Click()
Form1.Hide
Form00.Show
End Sub

Private Sub sub6sub2_Click()
Form1.Hide
Form000.Show
End Sub
