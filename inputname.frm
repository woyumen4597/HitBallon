VERSION 5.00
Begin VB.Form inputname 
   Caption         =   "                    ������Ϣ"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   5070
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdcancel 
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtfs 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblname 
      Caption         =   "���������ǣ�(������4����)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblfs 
      Caption         =   "���ķ����ǣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "inputname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdcancel_Click()
inputname.Hide
End Sub
Private Sub cmdok_Click()
Dim i, flag, place As Integer
Open "���а�.txt" For Random As #1
For i = 1 To 5
Get #1, i, Rank(i)
Next i
Close #1
Dim temp As user
temp.name = txtname.text
temp.fs = t
i = 6
flag = 1
While i > 1 And flag = 1
i = i - 1
If i <> 1 Then
    If t <= Rank(i - 1).fs Then
    flag = 0
    End If
 End If
Wend
place = i
If place < 5 Then
i = 5
While i >= (place + 1)
Rank(i).name = Rank(i - 1).name
Rank(i).fs = Rank(i - 1).fs
i = i - 1
Wend
End If
Rank(place).name = temp.name
Rank(place).fs = temp.fs
Open "���а�.txt" For Random As #1
For i = 1 To 5
Put #1, i, Rank(i)
Next i
Close #1
inputname.Hide
Dim words, answer As Variant
words = "     ��ϲ��Ϊ���а��" + Str(place) + "��"
answer = MsgBox(words, vbOKOnly, "��ϲ")
rankshow.Show
End Sub

