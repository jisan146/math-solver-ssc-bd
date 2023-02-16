VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "main menu"
      Height          =   495
      Left            =   360
      TabIndex        =   65
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   360
      TabIndex        =   64
      Top             =   7440
      Width           =   1215
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
      Left            =   7320
      TabIndex        =   57
      Text            =   "00"
      Top             =   7920
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   52
      Top             =   6600
      Width           =   11055
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
         Left            =   8520
         TabIndex        =   55
         Top             =   120
         Width           =   2175
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
         Left            =   2880
         TabIndex        =   54
         Text            =   "00"
         Top             =   120
         Width           =   1335
      End
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
         Left            =   1560
         TabIndex        =   53
         Text            =   "00"
         Top             =   120
         Width           =   975
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
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
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
      Left            =   7320
      TabIndex        =   51
      Top             =   7320
      Width           =   3495
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
      Left            =   8640
      TabIndex        =   50
      Text            =   "00"
      Top             =   6000
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   49
      Text            =   "00"
      Top             =   5400
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   48
      Text            =   "00"
      Top             =   4800
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   47
      Text            =   "00"
      Top             =   4200
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   46
      Text            =   "00"
      Top             =   3600
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   45
      Text            =   "00"
      Top             =   3000
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   44
      Text            =   "00"
      Top             =   2400
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   43
      Text            =   "00"
      Top             =   1800
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   42
      Text            =   "00"
      Top             =   1200
      Width           =   2175
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
      Left            =   8640
      TabIndex        =   41
      Text            =   "00"
      Top             =   600
      Width           =   2175
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
      Left            =   6360
      TabIndex        =   40
      Text            =   "00"
      Top             =   6000
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   39
      Text            =   "00"
      Top             =   5400
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   38
      Text            =   "00"
      Top             =   4800
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   37
      Text            =   "00"
      Top             =   4200
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   36
      Text            =   "00"
      Top             =   3600
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   35
      Text            =   "00"
      Top             =   3000
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   34
      Text            =   "00"
      Top             =   2400
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   33
      Text            =   "00"
      Top             =   1800
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   32
      Text            =   "00"
      Top             =   1200
      Width           =   1935
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
      Left            =   6360
      TabIndex        =   31
      Text            =   "00"
      Top             =   600
      Width           =   1935
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
      Left            =   4560
      TabIndex        =   30
      Text            =   "00"
      Top             =   600
      Width           =   1575
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
      Left            =   3000
      TabIndex        =   29
      Text            =   "00"
      Top             =   6000
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   28
      Text            =   "00"
      Top             =   5400
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   27
      Text            =   "00"
      Top             =   4800
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   26
      Text            =   "00"
      Top             =   4200
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   25
      Text            =   "00"
      Top             =   3600
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   24
      Text            =   "00"
      Top             =   3000
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   23
      Text            =   "00"
      Top             =   2400
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   22
      Text            =   "00"
      Top             =   1800
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   21
      Text            =   "00"
      Top             =   1200
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   20
      Text            =   "00"
      Top             =   600
      Width           =   1335
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
      Left            =   240
      TabIndex        =   19
      Text            =   "50"
      Top             =   6000
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
      Left            =   240
      TabIndex        =   18
      Text            =   "45"
      Top             =   5400
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
      Left            =   240
      TabIndex        =   17
      Text            =   "40"
      Top             =   4800
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
      Left            =   240
      TabIndex        =   16
      Text            =   "35"
      Top             =   4200
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
      Left            =   240
      TabIndex        =   15
      Text            =   "30"
      Top             =   3600
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
      Left            =   240
      TabIndex        =   14
      Text            =   "25"
      Top             =   3000
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
      Left            =   240
      TabIndex        =   13
      Text            =   "20"
      Top             =   2400
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
      Left            =   240
      TabIndex        =   12
      Text            =   "15"
      Top             =   1800
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
      Left            =   240
      TabIndex        =   11
      Text            =   "10"
      Top             =   1200
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
      Left            =   240
      TabIndex        =   10
      Text            =   "5"
      Top             =   600
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
      Left            =   1680
      TabIndex        =   9
      Text            =   "2"
      Top             =   6000
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
      Left            =   1680
      TabIndex        =   8
      Text            =   "3"
      Top             =   5400
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
      Left            =   1680
      TabIndex        =   7
      Text            =   "8"
      Top             =   4800
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
      Left            =   1680
      TabIndex        =   6
      Text            =   "13"
      Top             =   4200
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
      Left            =   1680
      TabIndex        =   5
      Text            =   "18"
      Top             =   3600
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
      Left            =   1680
      TabIndex        =   4
      Text            =   "20"
      Top             =   3000
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
      Left            =   1680
      TabIndex        =   3
      Text            =   "10"
      Top             =   2400
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
      TabIndex        =   2
      Text            =   "15"
      Top             =   1800
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
      Left            =   1680
      TabIndex        =   1
      Text            =   "12"
      Top             =   1200
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
      Left            =   1680
      TabIndex        =   0
      Text            =   "8"
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "         Fi | ( Xi - x) |"
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
      Left            =   8640
      TabIndex        =   63
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "          | Xi - x |"
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
      Left            =   6360
      TabIndex        =   62
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "   Arithmetic         Mean"
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
      Left            =   4800
      TabIndex        =   61
      Top             =   120
      Width           =   1215
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
      Left            =   3000
      TabIndex        =   60
      Top             =   120
      Width           =   1335
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
      Left            =   240
      TabIndex        =   59
      Top             =   120
      Width           =   975
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
      Left            =   1680
      TabIndex        =   58
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()
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




Private Sub Command4_Click()
Form2.Hide
Form8.Show
End Sub

Private Sub Command5_Click()
End
End Sub
