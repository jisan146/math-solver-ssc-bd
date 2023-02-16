VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text21 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   30
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text20 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   29
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   28
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3360
      TabIndex        =   27
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   26
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "main menu"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Try again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6480
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text15 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2880
      TabIndex        =   15
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      TabIndex        =   14
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   13
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   12
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   8
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   7
      Text            =   "1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   6
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   5
      Text            =   "1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   4
      Text            =   "1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   2
      Text            =   ".01"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   1
      Text            =   ".1"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   " Sequence value is correct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label19 
      Caption         =   " Sequence value is not correct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label20 
      Caption         =   "Please give me correct value and click try again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "This Sequence has no sum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "    r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "    a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Text19.Text = Text1.Text / (Text4.Text ^ Text5.Text)
Text20.Text = Text2.Text / (Text6.Text ^ Text7.Text)
Text21.Text = Text3.Text / (Text8.Text ^ Text9.Text)
Text15.Text = Text20.Text / Text19.Text
Text16.Text = Text21.Text / Text20.Text
If (Text15.Text) = (Text16.Text) Then
Label18.Visible = True
End If
If (Text15.Text) <> (Text16.Text) Then
Label19.Visible = True
Label20.Visible = True
Command5.Visible = True
End If
Text11.Text = Text19.Text
Text12.Text = Text20.Text / Text19.Text
If (Text12.Text) < 1 Then
Text14.Text = "r < 1"
Text13.Visible = True
Label3.Visible = False
End If
If (Text12.Text) > 1 Then
Text14.Text = "r > 1"
End If
If Text14.Text = "r > 1" Then
Label3.Visible = True
Text13.Visible = False
End If

Text13.Text = Text11.Text / (1 - Text12.Text)
End Sub

Private Sub Command3_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Label19.Visible = False
Label20.Visible = False
Command5.Visible = False
End Sub
