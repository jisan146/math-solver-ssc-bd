VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FF0000&
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   4740
   ClientTop       =   4485
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   1680
      Picture         =   "Form9.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   1680
      Picture         =   "Form9.frx":17E2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Form7.Show

Form0.Hide
Form000.Hide
Form1.Hide
Form2.Hide
Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide

Form8.Hide
Form9.Hide


End Sub

Private Sub Command2_Click()
Form2.Show

Form0.Hide
Form000.Hide
Form1.Hide

Form3.Hide
Form4.Hide
Form5.Hide
Form6.Hide
Form7.Hide
Form8.Hide
Form9.Hide


End Sub
