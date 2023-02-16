VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   1935
   ClientTop       =   1215
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11220
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9840
      TabIndex        =   58
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "main menu"
      Height          =   375
      Left            =   8520
      TabIndex        =   57
      Top             =   720
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
      Left            =   9120
      TabIndex        =   56
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   5640
      TabIndex        =   52
      Top             =   1200
      Visible         =   0   'False
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
      Left            =   4560
      TabIndex        =   51
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Details"
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
      Left            =   9120
      TabIndex        =   50
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Visible         =   0   'False
      Width           =   10335
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
         Left            =   7560
         TabIndex        =   49
         Top             =   600
         Width           =   1095
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
         Left            =   6840
         TabIndex        =   46
         Text            =   "1"
         Top             =   600
         Width           =   255
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
         Left            =   5280
         TabIndex        =   43
         Top             =   600
         Width           =   1095
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
         Left            =   1560
         TabIndex        =   35
         Top             =   1920
         Width           =   1095
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
         Left            =   1680
         TabIndex        =   34
         Top             =   1320
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
         Left            =   1680
         TabIndex        =   33
         Top             =   600
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
         Left            =   3480
         TabIndex        =   32
         Top             =   600
         Width           =   975
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
         TabIndex        =   31
         Text            =   "2"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         Caption         =   "}"
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
         Left            =   8760
         TabIndex        =   48
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   ")"
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
         Left            =   7200
         TabIndex        =   47
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   " -"
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
         TabIndex        =   45
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   "("
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
         Left            =   4920
         TabIndex        =   44
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   "+"
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
         Left            =   4560
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "{"
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
         Left            =   2760
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderWidth     =   4
         X1              =   1680
         X2              =   2640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "="
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
         TabIndex        =   38
         Top             =   1920
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "="
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
         TabIndex        =   37
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "  (S) n"
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
         TabIndex        =   36
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   7815
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
         Left            =   4560
         TabIndex        =   26
         Top             =   600
         Width           =   975
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
         Left            =   3120
         TabIndex        =   25
         Top             =   600
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
         Left            =   1680
         TabIndex        =   24
         Top             =   600
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
         Left            =   3120
         TabIndex        =   23
         Top             =   1440
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
         Left            =   6000
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   " -"
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
         Left            =   4200
         TabIndex        =   41
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "+"
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
         Left            =   2760
         TabIndex        =   40
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "    n"
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
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "="
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
         TabIndex        =   28
         Top             =   960
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderWidth     =   4
         X1              =   1560
         X2              =   5520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "="
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
         Left            =   5640
         TabIndex        =   27
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Details"
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
      Left            =   9120
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(S)n"
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
      Left            =   9120
      TabIndex        =   19
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   9120
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   1200
      TabIndex        =   17
      Text            =   "99"
      Top             =   2520
      Width           =   1575
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
      TabIndex        =   13
      Text            =   "1"
      Top             =   240
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
      TabIndex        =   12
      Text            =   "2"
      Top             =   240
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
      TabIndex        =   11
      Text            =   "3"
      Top             =   240
      Width           =   975
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
      Left            =   3360
      TabIndex        =   10
      Top             =   240
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
      Left            =   4440
      TabIndex        =   9
      Top             =   240
      Width           =   975
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
      Left            =   5520
      TabIndex        =   8
      Top             =   240
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
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Width           =   975
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
      Left            =   7680
      TabIndex        =   6
      Top             =   240
      Width           =   975
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
      Left            =   8760
      TabIndex        =   5
      Top             =   240
      Width           =   1095
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
      Left            =   9960
      TabIndex        =   4
      Top             =   240
      Width           =   975
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
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
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
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "n"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
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
      Left            =   9120
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   6720
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   4335
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
      Left            =   6720
      TabIndex        =   54
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
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
      Left            =   7200
      TabIndex        =   53
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "    n"
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
      TabIndex        =   16
      Top             =   2520
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
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "    d"
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
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()

Text29.Text = Text2.Text - Text1.Text
Text30.Text = Text3.Text - Text2.Text
If (Text29.Text) = (Text30.Text) Then
Label18.Visible = True


End If
If (Text29.Text) <> (Text30.Text) Then
Label19.Visible = True

Label20.Visible = True

Command5.Visible = True
End If









Text11.Text = Text1.Text
Text12.Text = Text2.Text - Text1.Text

Text17.Text = Text14.Text
Text16.Text = Text12.Text
Text15.Text = Text11.Text
Text18.Text = Text12.Text

Text19.Text = ((Text17.Text - -Text16.Text) - Text15.Text) / Text18.Text
Text20.Text = Text19.Text
Command3.Visible = True
Command2.Visible = True
Text19.Visible = True





End Sub

Private Sub Command2_Click()


Text23.Text = Text19.Text
Text24.Text = 2
Text22.Text = Text11.Text
Text26.Text = Text19.Text
Text28.Text = Text12.Text
Text25.Text = Text23.Text * (2 * Text22.Text + (Text26.Text - 1) * Text28.Text) / Text24.Text
Text13.Text = Text23.Text * (2 * Text22.Text + (Text26.Text - 1) * Text28.Text) / Text24.Text
Command4.Visible = True
Text13.Visible = True












End Sub

Private Sub Command3_Click()
Frame1.Visible = True
End Sub

Private Sub Command4_Click()
Frame1.Visible = True
Frame2.Visible = True
End Sub

Private Sub Command5_Click()
Command5.Visible = False
Label19.Visible = False
Label20.Visible = False

End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
Form5.Hide
Form8.Show
End Sub
