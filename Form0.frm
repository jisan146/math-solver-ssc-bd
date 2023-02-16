VERSION 5.00
Begin VB.Form Form0 
   BackColor       =   &H00FF0000&
   Caption         =   "ASAT Software"
   ClientHeight    =   6750
   ClientLeft      =   3615
   ClientTop       =   3060
   ClientWidth     =   8265
   Icon            =   "Form0.frx":0000
   LinkTopic       =   "Form37"
   ScaleHeight     =   6750
   ScaleWidth      =   8265
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Are you want to use this software?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ASAT Software"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Menu s1 
      Caption         =   "About"
   End
   Begin VB.Menu s2 
      Caption         =   "Help"
   End
   Begin VB.Menu s3 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form0.Hide

Form8.Show
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub s1_Click()
Form000.Show
Form0.Hide
End Sub

Private Sub s3_Click()
End
End Sub
