VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   2310
   ClientTop       =   2175
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10215
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   2880
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":55CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":6F30
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":9252
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":AF14
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Form4.Show

Form0.Hide
Form000.Hide
Form1.Hide
Form2.Hide
Form3.Hide

Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide


End Sub

Private Sub Command2_Click()
Form5.Show

Form0.Hide
Form000.Hide
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide

Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide


End Sub

Private Sub Command3_Click()
Form1.Show

Form0.Hide
Form000.Hide

Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide


End Sub

Private Sub Command4_Click()
Form9.Show

Form0.Hide
Form000.Hide
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide



End Sub

Private Sub Command5_Click()
Form11.Show
Form9.Hide
Form0.Hide
Form000.Hide
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
End Sub
