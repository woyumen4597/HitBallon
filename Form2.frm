VERSION 5.00
Begin VB.Form rankshow 
   Caption         =   "排行榜"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   7620
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label textfs 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   4
      Left            =   4200
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label textfs 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   3
      Left            =   4200
      TabIndex        =   14
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label textfs 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   2
      Left            =   4200
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label textfs 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   1
      Left            =   4200
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label textfs 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Index           =   0
      Left            =   4200
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label textname 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label textname 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label textname 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label textname 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label textname 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "rankshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
rankshow.Hide
'Form1.Show
End Sub
Private Sub Form_Load()
Dim i As Integer
Open "排行榜.txt" For Random As #1
For i = 1 To 5
Get #1, i, Rank(i)
Next
For i = 1 To 5
If Rank(i).fs > 0 Then
    textname(i - 1).Caption = Rank(i).name
    textfs(i - 1).Caption = Rank(i).fs
    End If
    Next i
    Close #1
End Sub

