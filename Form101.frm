VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FF0000&
   Caption         =   "Form10"
   ClientHeight    =   8445
   ClientLeft      =   2310
   ClientTop       =   840
   ClientWidth     =   10245
   LinkTopic       =   "Form10"
   ScaleHeight     =   8445
   ScaleWidth      =   10245
   Begin VB.CommandButton Command7 
      Caption         =   "Complete Other style"
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
      Left            =   8400
      TabIndex        =   55
      Top             =   6360
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text31 
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
      Left            =   3960
      TabIndex        =   54
      Text            =   "1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text30 
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
      Left            =   3960
      TabIndex        =   53
      Text            =   "1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text29 
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
      Left            =   2280
      TabIndex        =   49
      Text            =   "1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text28 
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
      Left            =   4320
      TabIndex        =   48
      Text            =   "2"
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text27 
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
      Left            =   6840
      TabIndex        =   47
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text26 
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
      Left            =   240
      TabIndex        =   46
      Text            =   "0"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Other style"
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
      Left            =   8400
      TabIndex        =   45
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text25 
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
      Left            =   8400
      TabIndex        =   44
      Text            =   "1"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8880
      TabIndex        =   43
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "main menu"
      Height          =   495
      Left            =   8880
      TabIndex        =   42
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
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
      Left            =   8400
      TabIndex        =   41
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text24 
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
      Left            =   8400
      TabIndex        =   40
      Text            =   "12"
      Top             =   5160
      Width           =   1695
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
      Height          =   495
      Left            =   8400
      TabIndex        =   39
      Text            =   "12"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text23 
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
      Left            =   3960
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text22 
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
      Left            =   3960
      TabIndex        =   36
      Top             =   6240
      Width           =   1215
   End
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
      Height          =   495
      Left            =   2040
      TabIndex        =   35
      Top             =   5520
      Width           =   1215
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
      Height          =   495
      Left            =   2040
      TabIndex        =   34
      Top             =   6240
      Width           =   1215
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
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   4800
      Width           =   1215
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
      Height          =   495
      Left            =   4440
      TabIndex        =   29
      Top             =   4080
      Width           =   1215
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
      Height          =   495
      Left            =   2040
      TabIndex        =   28
      Top             =   4440
      Width           =   1215
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
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
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
      Height          =   495
      Left            =   4440
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
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
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
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
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   1215
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
      Height          =   495
      Left            =   960
      TabIndex        =   18
      Top             =   2280
      Width           =   1215
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
      Height          =   495
      Left            =   960
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
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
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1215
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
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Complete"
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
      Left            =   8400
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
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
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   960
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
      Left            =   8400
      TabIndex        =   6
      Top             =   3840
      Width           =   1695
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
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   1815
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
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1335
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
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Text            =   "0"
      Top             =   120
      Width           =   1215
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
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Text            =   "3"
      Top             =   120
      Width           =   495
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
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "2"
      Top             =   120
      Width           =   495
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
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line6 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3960
      X2              =   5160
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   52
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   51
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   50
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderWidth     =   3
      X1              =   8400
      X2              =   10080
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000009&
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   38
      Top             =   5880
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   3960
      X2              =   5160
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   2040
      X2              =   3240
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000009&
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   33
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000009&
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   32
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   31
      Top             =   4440
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   27
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3960
      X2              =   6120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000009&
      Caption         =   " -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   26
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      Caption         =   " 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   25
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   24
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      Caption         =   " ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   " a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   " r"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   14
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   " +"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   " ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Text26.Visible = False
Text29.Visible = False
Label18.Visible = False
Text28.Visible = False
Label16.Visible = False
Text27.Visible = False
Label17.Visible = False
Text30.Visible = False
Text31.Visible = False

If (Text3.Text) = "" Then
Text5.Text = Label1.Caption + Text2.Text
End If
If (Text3.Text) <> "" Then
Text5.Text = Label1.Caption + Text2.Text + Text3.Text
End If
If (Text4.Text) = "" Then
Text5.Text = Label1.Caption + Text2.Text + Text3.Text
End If
If (Text4.Text) <> "" Then
Text5.Text = Label1.Caption + Text2.Text + Text3.Text + Text4.Text
End If


If (Text3.Text) = "" Then
Text8.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text + Text2.Text + Text3.Text + Text4.Text + "......."


End If
If (Text3.Text) <> "" Then
Text8.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text + Text2.Text + Text3.Text + Text4.Text + "......."


End If

If (Text4.Text) = "" Then
Text8.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text + Text2.Text + Text3.Text + Text4.Text + "......."


End If
If (Text4.Text) <> "" Then
Text8.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text + Text2.Text + Text3.Text + Text4.Text + "......."


End If


If (Text3.Text) = "" Then
Text6.Text = ".0" + Text2.Text

End If

If (Text3.Text) <> "" Then
Text6.Text = ".00" + Text2.Text + Text3.Text

End If
If (Text4.Text) = "" Then
Text6.Text = ".00" + Text2.Text + Text3.Text

End If

If (Text4.Text) <> "" Then
Text6.Text = ".000" + Text2.Text + Text3.Text + Text4.Text
End If

If (Text3.Text) = "" Then
Text7.Text = ".00" + Text2.Text
End If
If (Text3.Text) <> "" Then
Text7.Text = ".000" + Text2.Text + Text3.Text
End If
If (Text4.Text) = "" Then
Text7.Text = ".0000" + Text2.Text + Text3.Text
End If
If (Text4.Text) <> "" Then
Text7.Text = ".000000" + Text2.Text + Text3.Text + Text4.Text
End If
If (Text3.Text) = (Text4.Text) Then
Text6.Text = ".0" + Text2.Text
Text7.Text = ".00" + Text2.Text
End If
Text9.Text = Text1.Text
Text10.Text = Text5.Text
Text11.Text = Text6.Text / Text5.Text

If (Text3.Text) = "" Then
Text12.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text


End If
If (Text3.Text) <> "" Then
Text12.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text


End If

If (Text4.Text) = "" Then
Text12.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text


End If
If (Text4.Text) <> "" Then
Text12.Text = Text1.Text + Label1.Caption + Text2.Text + Text3.Text + Text4.Text

End If
Text13.Text = Text1.Text
Text14.Text = Text10.Text
Text15.Text = Text11.Text
Text19.Text = Label9.Caption - Text15.Text
Text18.Text = Text14.Text
Text17.Text = Text13.Text
Text20.Text = Text19.Text

Text21.Text = (Text17.Text * Text19.Text) - -Text18.Text


Command2.Visible = True

If (Text2.Text) <> "" Then
Text23.Text = 10
Text22.Text = 10

End If



If (Text3.Text) <> "" Then
Text23.Text = 100
Text22.Text = 100

End If

If (Text4.Text) <> "" Then
Text23.Text = 1000
Text22.Text = 1000

End If



Text16.Text = Text21.Text * Text23.Text
Text24.Text = Text20.Text * Text22.Text

End Sub

Private Sub Command2_Click()
Command1.Visible = False
Text16.Visible = True
Text24.Visible = True
Line5.Visible = True
Text16.Text = Text21.Text * Text23.Text
Text24.Text = Text20.Text * Text22.Text

Text16.Text = Text16.Text / Text25.Text
Text24.Text = Text24.Text / Text25.Text




End Sub

Private Sub Command3_Click()
Command1.Visible = True
Command2.Visible = False


Text1.Visible = True
Label1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Label11.Visible = True
Text8.Visible = True
Label4.Visible = True
Text7.Visible = True
Label2.Visible = True
Text8.Visible = True
Label3.Visible = True
Text26.Visible = False
Text29.Visible = False
Label18.Visible = False
Text28.Visible = False
Label16.Visible = False
Text27.Visible = False
Label17.Visible = False
Text30.Visible = False
Text31.Visible = False








Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text16.Text = ""
Text17.Text = ""
Text18.Text = ""
Text19.Text = ""
Text20.Text = ""
Text21.Text = ""
Text22.Text = ""
Text23.Text = ""
Text24.Text = ""
Text25.Text = 1
End Sub




Private Sub Command5_Click()
End
End Sub

Private Sub Command4_Click()
Form8.Show
Form11.Hide
End Sub

Private Sub Command6_Click()
Command7.Visible = True
Text1.Visible = False
Label1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Label11.Visible = False
Text8.Visible = False
Label4.Visible = False
Text7.Visible = False
Label2.Visible = False
Text8.Visible = False
Label3.Visible = False
Text26.Visible = True
Text29.Visible = True
Label18.Visible = True
Text28.Visible = True
Label16.Visible = True
Text27.Visible = True
Label17.Visible = True
Text30.Visible = True
Text31.Visible = True




Text10.Text = Text29.Text
Text11.Text = Text28.Text / Text29.Text
Text12.Text = Text26.Text
Text13.Text = Text26.Text
Text14.Text = Text10.Text
Text15.Text = Text11.Text
Text17.Text = Text12.Text
Text18.Text = Text14.Text
Text19.Text = Label9 - Text15.Text

Text20.Text = Text19.Text
Text21.Text = (Text17.Text * Text19.Text) - -Text18.Text


Command2.Visible = True

If (Text2.Text) <> "" Then
Text23.Text = 10
Text22.Text = 10

End If



If (Text3.Text) <> "" Then
Text23.Text = 100
Text22.Text = 100

End If

If (Text4.Text) <> "" Then
Text23.Text = 1000
Text22.Text = 1000

End If



Text16.Text = Text21.Text * Text30.Text
Text24.Text = Text20.Text * Text31.Text

End Sub

Private Sub Command7_Click()

Text16.Text = Text21.Text * Text31.Text
Text24.Text = Text20.Text * Text30.Text

Text16.Text = Text16.Text / Text25.Text
Text24.Text = Text24.Text / Text25.Text

End Sub
