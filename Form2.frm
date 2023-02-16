VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   2115
   ClientTop       =   1020
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   11265
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9840
      TabIndex        =   96
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "main menu"
      Height          =   375
      Left            =   9840
      TabIndex        =   95
      Top             =   480
      Width           =   1215
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
      TabIndex        =   77
      Text            =   "51"
      Top             =   1440
      Width           =   855
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
      Left            =   120
      TabIndex        =   76
      Text            =   "61"
      Top             =   2040
      Width           =   855
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
      Left            =   120
      TabIndex        =   75
      Text            =   "71"
      Top             =   2640
      Width           =   855
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
      TabIndex        =   74
      Text            =   "81"
      Top             =   3240
      Width           =   855
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
      Left            =   120
      TabIndex        =   73
      Text            =   "91"
      Top             =   3840
      Width           =   855
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
      Left            =   120
      TabIndex        =   72
      Text            =   "00"
      Top             =   4440
      Width           =   855
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
      Left            =   120
      TabIndex        =   71
      Text            =   "00"
      Top             =   5040
      Width           =   855
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
      Left            =   120
      TabIndex        =   70
      Text            =   "00"
      Top             =   5640
      Width           =   855
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
      Left            =   120
      TabIndex        =   69
      Text            =   "00"
      Top             =   6240
      Width           =   855
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
      Left            =   120
      TabIndex        =   68
      Text            =   "00"
      Top             =   6840
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   67
      Text            =   "60"
      Top             =   1440
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   66
      Text            =   "70"
      Top             =   2040
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   65
      Text            =   "80"
      Top             =   2640
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   64
      Text            =   "90"
      Top             =   3240
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   63
      Text            =   "100"
      Top             =   3840
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   62
      Text            =   "00"
      Top             =   4440
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   61
      Text            =   "00"
      Top             =   5040
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   60
      Text            =   "00"
      Top             =   5640
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   59
      Text            =   "00"
      Top             =   6240
      Width           =   855
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
      Height          =   375
      Left            =   1920
      TabIndex        =   58
      Text            =   "00"
      Top             =   6840
      Width           =   855
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
      Height          =   420
      Left            =   3120
      TabIndex        =   57
      Text            =   "10"
      Top             =   1440
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   56
      Text            =   "15"
      Top             =   2040
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   55
      Text            =   "20"
      Top             =   2640
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   54
      Text            =   "10"
      Top             =   3240
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   53
      Text            =   "5"
      Top             =   3840
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   52
      Text            =   "00"
      Top             =   4440
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   51
      Text            =   "00"
      Top             =   5040
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   50
      Text            =   "00"
      Top             =   5640
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   49
      Text            =   "00"
      Top             =   6240
      Width           =   975
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
      Height          =   420
      Left            =   3120
      TabIndex        =   48
      Text            =   "00"
      Top             =   6840
      Width           =   975
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
      Height          =   420
      Left            =   4320
      TabIndex        =   47
      Text            =   "00"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text32 
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
      Left            =   4320
      TabIndex        =   46
      Text            =   "00"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text33 
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
      Left            =   4320
      TabIndex        =   45
      Text            =   "00"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text34 
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
      Left            =   4320
      TabIndex        =   44
      Text            =   "00"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text35 
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
      Left            =   4320
      TabIndex        =   43
      Text            =   "00"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text36 
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
      Left            =   4320
      TabIndex        =   42
      Text            =   "00"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text37 
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
      Left            =   4320
      TabIndex        =   41
      Text            =   "00"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text38 
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
      Left            =   4320
      TabIndex        =   40
      Text            =   "00"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text39 
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
      Left            =   4320
      TabIndex        =   39
      Text            =   "00"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text40 
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
      Left            =   4320
      TabIndex        =   38
      Text            =   "00"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text41 
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
      Left            =   5640
      TabIndex        =   37
      Text            =   "00"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text42 
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
      Left            =   5640
      TabIndex        =   36
      Text            =   "00"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text43 
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
      Left            =   5640
      TabIndex        =   35
      Text            =   "00"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text44 
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
      Left            =   5640
      TabIndex        =   34
      Text            =   "00"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text45 
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
      Left            =   5640
      TabIndex        =   33
      Text            =   "00"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text46 
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
      Left            =   5640
      TabIndex        =   32
      Text            =   "00"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text47 
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
      Left            =   5640
      TabIndex        =   31
      Text            =   "00"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text48 
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
      Left            =   5640
      TabIndex        =   30
      Text            =   "00"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text49 
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
      Left            =   5640
      TabIndex        =   29
      Text            =   "00"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text50 
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
      Left            =   5640
      TabIndex        =   28
      Text            =   "00"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text00 
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
      Left            =   6720
      TabIndex        =   27
      Text            =   "00"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text51 
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
      Left            =   8160
      TabIndex        =   26
      Text            =   "00"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text52 
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
      Left            =   8160
      TabIndex        =   25
      Text            =   "00"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text53 
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
      Left            =   8160
      TabIndex        =   24
      Text            =   "00"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text54 
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
      Left            =   8160
      TabIndex        =   23
      Text            =   "00"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text55 
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
      Left            =   8160
      TabIndex        =   22
      Text            =   "00"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text56 
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
      Left            =   8160
      TabIndex        =   21
      Text            =   "00"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text57 
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
      Left            =   8160
      TabIndex        =   20
      Text            =   "00"
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text58 
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
      Left            =   8160
      TabIndex        =   19
      Text            =   "00"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text59 
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
      Left            =   8160
      TabIndex        =   18
      Text            =   "00"
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text60 
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
      Left            =   8160
      TabIndex        =   17
      Text            =   "00"
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text61 
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
      Left            =   9480
      TabIndex        =   16
      Text            =   "00"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text62 
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
      Left            =   9480
      TabIndex        =   15
      Text            =   "00"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text63 
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
      Left            =   9480
      TabIndex        =   14
      Text            =   "00"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text64 
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
      Left            =   9480
      TabIndex        =   13
      Text            =   "00"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text65 
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
      Left            =   9480
      TabIndex        =   12
      Text            =   "00"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text66 
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
      Left            =   9480
      TabIndex        =   11
      Text            =   "00"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text67 
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
      Left            =   9480
      TabIndex        =   10
      Text            =   "00"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text68 
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
      Left            =   9480
      TabIndex        =   9
      Text            =   "00"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text69 
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
      Left            =   9480
      TabIndex        =   8
      Text            =   "00"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text70 
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
      Left            =   9480
      TabIndex        =   7
      Text            =   "00"
      Top             =   6840
      Width           =   1575
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
      Left            =   7680
      TabIndex        =   6
      Top             =   8160
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   7440
      Width           =   11055
      Begin VB.TextBox Text72 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3000
         TabIndex        =   4
         Text            =   "00"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text73 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   5520
         TabIndex        =   3
         Text            =   "00"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox Text74 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   9360
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox Text75 
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
      Left            =   7680
      TabIndex        =   0
      Text            =   "00"
      Top             =   8760
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   94
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   93
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   92
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   91
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   90
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   89
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   88
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   87
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   86
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "To"
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
      Left            =   1200
      TabIndex        =   85
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "            Class"
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
      TabIndex        =   84
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   83
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "    Xi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   82
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "   Fi x Xi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   81
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Arithmetic      Mean"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   80
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   " | Xi - x |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   79
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "     Fi | ( Xi - x) |"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   78
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
Text31.Text = (Text1.Text - -Text11.Text) / 2
Text32.Text = (Text2.Text - -Text12.Text) / 2
Text33.Text = (Text3.Text - -Text13.Text) / 2
Text34.Text = (Text4.Text - -Text14.Text) / 2
Text35.Text = (Text5.Text - -Text15.Text) / 2
Text36.Text = (Text6.Text - -Text16.Text) / 2
Text37.Text = (Text7.Text - -Text17.Text) / 2
Text38.Text = (Text8.Text - -Text18.Text) / 2
Text39.Text = (Text9.Text - -Text19.Text) / 2
Text40.Text = (Text10.Text - -Text20.Text) / 2
Text41.Text = (Text21.Text * Text31.Text)
Text42.Text = (Text22.Text * Text32.Text)
Text43.Text = (Text23.Text * Text33.Text)
Text44.Text = (Text24.Text * Text34.Text)
Text45.Text = (Text25.Text * Text35.Text)
Text46.Text = (Text26.Text * Text36.Text)
Text47.Text = (Text27.Text * Text37.Text)
Text48.Text = (Text28.Text * Text38.Text)
Text49.Text = (Text29.Text * Text39.Text)
Text50.Text = (Text30.Text * Text40.Text)
Text72.Text = (Text21.Text - -Text22.Text - -Text23.Text - -Text24.Text - -Text25.Text - -Text26.Text - -Text27.Text - -Text28.Text - -Text29.Text - -Text30.Text)
Text73.Text = (Text41.Text - -Text42.Text - -Text43.Text - -Text44.Text - -Text45.Text - -Text46.Text - -Text47.Text - -Text48.Text - -Text49.Text - -Text50.Text)
Text00.Text = Text73.Text / Text72.Text
Text51.Text = Text31.Text - Text00.Text
Text52.Text = Text32.Text - Text00.Text
Text53.Text = Text33.Text - Text00.Text
Text54.Text = Text34.Text - Text00.Text
Text55.Text = Text35.Text - Text00.Text
Text56.Text = Text36.Text - Text00.Text
Text57.Text = Text37.Text - Text00.Text
Text58.Text = Text38.Text - Text00.Text
Text59.Text = Text39.Text - Text00.Text
Text60.Text = Text40.Text - Text00.Text
Text61.Text = (Text51.Text) * (Text21.Text)
Text62.Text = (Text52.Text) * Text22.Text
Text63.Text = (Text53.Text) * Text23.Text
Text64.Text = (Text54.Text) * Text24.Text
Text65.Text = (Text55.Text) * Text25.Text
Text66.Text = (Text56.Text) * Text26.Text
Text67.Text = (Text57.Text) * Text27.Text
Text68.Text = (Text58.Text) * Text28.Text
Text69.Text = (Text59.Text) * Text29.Text
Text70.Text = (Text60.Text) * Text30.Text

If (Text61.Text) <= -1 Then
Text61.Text = Text61.Text * (-1)
End If
If (Text62.Text) <= -1 Then
Text62.Text = Text62.Text * (-1)
End If
If (Text63.Text) <= -1 Then
Text63.Text = Text63.Text * (-1)
End If
If (Text64.Text) <= -1 Then
Text64.Text = Text64.Text * (-1)
End If
If (Text65.Text) <= -1 Then
Text65.Text = Text65.Text * (-1)
End If
If (Text66.Text) <= -1 Then
Text66.Text = Text66.Text * (-1)
End If
If (Text67.Text) <= -1 Then
Text67.Text = Text67.Text * (-1)
End If
If (Text68.Text) <= -1 Then
Text68.Text = Text68.Text * (-1)
End If
If (Text69.Text) <= -1 Then
Text69.Text = Text69.Text * (-1)
End If
If (Text70.Text) <= -1 Then
Text70.Text = Text70.Text * (-1)
End If

If (Text51.Text) <= -1 Then
Text51.Text = Text51.Text * (-1)
End If

If (Text52.Text) <= -1 Then
Text52.Text = Text52.Text * (-1)
End If

If (Text53.Text) <= -1 Then
Text53.Text = Text53.Text * (-1)
End If

If (Text54.Text) <= -1 Then
Text54.Text = Text54.Text * (-1)
End If

If (Text55.Text) <= -1 Then
Text55.Text = Text55.Text * (-1)
End If

If (Text56.Text) <= -1 Then
Text56.Text = Text56.Text * (-1)
End If

If (Text57.Text) <= -1 Then
Text57.Text = Text57.Text * (-1)
End If

If (Text58.Text) <= -1 Then
Text58.Text = Text58.Text * (-1)
End If

If (Text59.Text) <= -1 Then
Text59.Text = Text59.Text * (-1)
End If

If (Text60.Text) <= -1 Then
Text60.Text = Text60.Text * (-1)
End If
Text74.Text = Text61.Text - -Text62.Text - -Text63.Text - -Text64.Text - -Text65.Text - -Text66.Text - -Text67.Text - -Text68.Text - -Text69.Text - -Text70.Text
Text75.Text = (Text74.Text / Text72.Text)
End Sub



Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form7.Hide
Form8.Show
End Sub
