VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   65
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "main menu"
      Height          =   495
      Left            =   120
      TabIndex        =   64
      Top             =   7800
      Width           =   1215
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
      Left            =   120
      TabIndex        =   57
      Text            =   "0"
      Top             =   600
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
      Left            =   120
      TabIndex        =   56
      Text            =   "1"
      Top             =   1200
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
      Left            =   120
      TabIndex        =   55
      Text            =   "2"
      Top             =   1800
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
      Left            =   120
      TabIndex        =   54
      Text            =   "3"
      Top             =   2400
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
      Left            =   120
      TabIndex        =   53
      Text            =   "4"
      Top             =   3000
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
      Left            =   120
      TabIndex        =   52
      Text            =   "5"
      Top             =   3600
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
      Left            =   120
      TabIndex        =   51
      Text            =   "6"
      Top             =   4200
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
      Left            =   120
      TabIndex        =   50
      Text            =   "7"
      Top             =   4800
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
      Left            =   120
      TabIndex        =   49
      Text            =   "0"
      Top             =   5400
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
      Left            =   120
      TabIndex        =   48
      Text            =   "0"
      Top             =   6000
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
      Left            =   2640
      TabIndex        =   47
      Text            =   "00"
      Top             =   600
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   46
      Text            =   "00"
      Top             =   1200
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   45
      Text            =   "00"
      Top             =   1800
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   44
      Text            =   "00"
      Top             =   2400
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   43
      Text            =   "00"
      Top             =   3000
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   42
      Text            =   "00"
      Top             =   3600
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   41
      Text            =   "00"
      Top             =   4200
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   40
      Text            =   "00"
      Top             =   4800
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   39
      Text            =   "00"
      Top             =   5400
      Width           =   1095
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
      Left            =   2640
      TabIndex        =   38
      Text            =   "00"
      Top             =   6000
      Width           =   1095
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
      Left            =   3840
      TabIndex        =   37
      Text            =   "00"
      Top             =   600
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
      Left            =   6000
      TabIndex        =   36
      Text            =   "00"
      Top             =   600
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
      Left            =   6000
      TabIndex        =   35
      Text            =   "00"
      Top             =   1200
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
      Left            =   6000
      TabIndex        =   34
      Text            =   "00"
      Top             =   1800
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
      Left            =   6000
      TabIndex        =   33
      Text            =   "00"
      Top             =   2400
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
      Left            =   6000
      TabIndex        =   32
      Text            =   "00"
      Top             =   3000
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
      Left            =   6000
      TabIndex        =   31
      Text            =   "00"
      Top             =   3600
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
      Left            =   6000
      TabIndex        =   30
      Text            =   "00"
      Top             =   4200
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
      Left            =   6000
      TabIndex        =   29
      Text            =   "00"
      Top             =   4800
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
      Left            =   6000
      TabIndex        =   28
      Text            =   "00"
      Top             =   5400
      Width           =   1935
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
      Left            =   6000
      TabIndex        =   27
      Text            =   "00"
      Top             =   6000
      Width           =   1935
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
      Left            =   8400
      TabIndex        =   26
      Text            =   "00"
      Top             =   600
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   25
      Text            =   "00"
      Top             =   1200
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   24
      Text            =   "00"
      Top             =   1800
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   23
      Text            =   "00"
      Top             =   2400
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   22
      Text            =   "00"
      Top             =   3000
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   21
      Text            =   "00"
      Top             =   3600
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   20
      Text            =   "00"
      Top             =   4200
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   19
      Text            =   "00"
      Top             =   4800
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   18
      Text            =   "00"
      Top             =   5400
      Width           =   2295
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
      Left            =   8400
      TabIndex        =   17
      Text            =   "00"
      Top             =   6000
      Width           =   2295
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
      Left            =   7560
      TabIndex        =   16
      Top             =   7200
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   6480
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
         Left            =   1440
         TabIndex        =   14
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
         Left            =   2640
         TabIndex        =   13
         Text            =   "00"
         Top             =   120
         Width           =   1095
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
         Left            =   8400
         TabIndex        =   12
         Text            =   "00"
         Top             =   120
         Width           =   2295
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
         TabIndex        =   15
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
      Left            =   7560
      TabIndex        =   10
      Text            =   "00"
      Top             =   7800
      Width           =   3495
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
      Left            =   1440
      TabIndex        =   9
      Text            =   "5"
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
      Left            =   1440
      TabIndex        =   8
      Text            =   "10"
      Top             =   1200
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
      Left            =   1440
      TabIndex        =   7
      Text            =   "15"
      Top             =   1800
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
      Left            =   1440
      TabIndex        =   6
      Text            =   "18"
      Top             =   2400
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
      Left            =   1440
      TabIndex        =   5
      Text            =   "25"
      Top             =   3000
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
      Left            =   1440
      TabIndex        =   4
      Text            =   "19"
      Top             =   3600
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
      Left            =   1440
      TabIndex        =   3
      Text            =   "11"
      Top             =   4200
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
      Left            =   1440
      TabIndex        =   2
      Text            =   "6"
      Top             =   4800
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "00"
      Top             =   5400
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "00"
      Top             =   6000
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
      Left            =   1440
      TabIndex        =   63
      Top             =   120
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
      Left            =   120
      TabIndex        =   62
      Top             =   120
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
      Left            =   2640
      TabIndex        =   61
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "     Arithmetic                Mean"
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
      Left            =   4080
      TabIndex        =   60
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label16 
      Caption         =   "            Xi - x"
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
      Left            =   6000
      TabIndex        =   59
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label17 
      Caption         =   "          Fi( Xi - x)^2"
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
      Left            =   8400
      TabIndex        =   58
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
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
Text61.Text = (Text51.Text) ^ 2 * Text21.Text
Text62.Text = (Text52.Text) ^ 2 * Text22.Text
Text63.Text = (Text53.Text) ^ 2 * Text23.Text
Text64.Text = (Text54.Text) ^ 2 * Text24.Text
Text65.Text = (Text55.Text) ^ 2 * Text25.Text
Text66.Text = (Text56.Text) ^ 2 * Text26.Text
Text67.Text = (Text57.Text) ^ 2 * Text27.Text
Text68.Text = (Text58.Text) ^ 2 * Text28.Text
Text69.Text = (Text59.Text) ^ 2 * Text29.Text
Text70.Text = (Text60.Text) ^ 2 * Text30.Text
Text74.Text = (Text61.Text - -Text62.Text - -Text63.Text - -Text64.Text - -Text65.Text - -Text66.Text - -Text67.Text - -Text68.Text - -Text69.Text - -Text70.Text)
Text75.Text = (Text74.Text / Text72.Text) ^ (1 / 2)

End Sub



Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form8.Show
Form3.Hide
End Sub
