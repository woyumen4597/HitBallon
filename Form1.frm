VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9990
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   15795
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15240
      Top             =   5040
   End
   Begin VB.PictureBox Picture1 
      Height          =   11295
      Left            =   0
      Picture         =   "Form1.frx":33B8F
      ScaleHeight     =   11235
      ScaleWidth      =   13275
      TabIndex        =   6
      Top             =   -720
      Width           =   13335
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   3840
         Top             =   5520
      End
      Begin VB.Image Image5 
         Height          =   1380
         Index           =   1
         Left            =   9840
         Picture         =   "Form1.frx":6771E
         Stretch         =   -1  'True
         Top             =   9240
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   1380
         Index           =   0
         Left            =   11160
         Picture         =   "Form1.frx":680D2
         Stretch         =   -1  'True
         Top             =   9240
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   1380
         Index           =   1
         Left            =   1800
         Picture         =   "Form1.frx":68A86
         Stretch         =   -1  'True
         Top             =   9360
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   1380
         Index           =   1
         Left            =   3360
         Picture         =   "Form1.frx":6A801
         Stretch         =   -1  'True
         Top             =   9360
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   1380
         Index           =   0
         Left            =   8520
         Picture         =   "Form1.frx":6B553
         Stretch         =   -1  'True
         Top             =   9240
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   1380
         Index           =   0
         Left            =   6840
         Picture         =   "Form1.frx":6D128
         Stretch         =   -1  'True
         Top             =   9240
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   1380
         Index           =   0
         Left            =   4920
         Picture         =   "Form1.frx":6DE7A
         Stretch         =   -1  'True
         Top             =   9240
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   840
         Picture         =   "Form1.frx":6FBF5
         Stretch         =   -1  'True
         Top             =   500
         Width           =   2265
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   615
      Index           =   2
      Left            =   13560
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "暂停"
      Height          =   615
      Index           =   1
      Left            =   13560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Index           =   0
      Left            =   13560
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Score1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   13560
      TabIndex        =   9
      Top             =   5760
      Width           =   240
   End
   Begin VB.Label Time 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "60"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   13560
      TabIndex        =   8
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "难度：简单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   13560
      TabIndex        =   7
      Top             =   6600
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   13560
      TabIndex        =   5
      Top             =   5280
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   14280
      TabIndex        =   4
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Time     Left"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   13560
      TabIndex        =   3
      Top             =   3480
      Width           =   1170
   End
   Begin VB.Menu start 
      Caption         =   "开始(&S)"
      Begin VB.Menu start1 
         Caption         =   "开始"
      End
      Begin VB.Menu pause 
         Caption         =   "暂停"
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "工具(&T)"
      Begin VB.Menu txt 
         Caption         =   "记事本"
      End
      Begin VB.Menu calculator 
         Caption         =   "计算器"
      End
   End
   Begin VB.Menu Level 
      Caption         =   "难度(&L)"
      Begin VB.Menu easy 
         Caption         =   "简单"
      End
      Begin VB.Menu normal 
         Caption         =   "一般"
      End
      Begin VB.Menu difficult 
         Caption         =   "困难"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu text 
         Caption         =   "操作说明"
      End
      Begin VB.Menu about 
         Caption         =   "关于"
      End
   End
   Begin VB.Menu Rank 
      Caption         =   "排行榜(&R)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim x As Integer
Dim Score As Double
Dim index, i, k As Integer
Dim msg, msg1, msg2, msg3, msg4, msg5, msg6, msg7, msg8, Title As Variant
Const vbKeyLeft = 37  '方向键左
Const vbKeyRight = 39 '方向键右
Private Sub about_Click() '关于作者
msg = "版本1.0.0" & Chr(13) & Chr(10) & "作者：金荣钏" & Chr(13) & Chr(10) & "学号：14124879" & Chr(13) & Chr(10) & "警告：版权所有，复制必究"
Title = "关于打气球"
MsgBox msg, vbOKOnly, Title
End Sub
Private Sub calculator_Click()  '打开计算器
x = Shell("c:\windows\system32\calc.exe", 1)
End Sub
Private Sub Command1_Click(index As Integer)  '开始键
If index = 0 Then
Score = 0            '初始化分数
Time.Caption = 60   '初始化时间
Picture1.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Command1(0).Caption = "重新开始"
Command1(1).Caption = "暂停"
Call ks
Picture1.SetFocus
Command1(0).Enabled = False
Command1(0).Enabled = True
Image2(0).Visible = True
Image2(1).Visible = True
Image3(0).Visible = True
Image3(1).Visible = True
Image5(0).Visible = True
Image5(1).Visible = True
Image4(0).Visible = True
End If
If index = 1 Then   '暂停键
If Timer1.Enabled = True Then
Command1(1).Caption = "恢复"
Timer1.Enabled = False
Timer2.Enabled = False
Else
Command1(1).Caption = "暂停"
Timer1.Enabled = True
Timer2.Enabled = True
End If
Command1(0).Enabled = True '开始键
End If
If index = 2 Then  '退出键
End
End If
End Sub
Public Sub ks() '随机气球位置
Dim x As Double
Dim y As Double
Dim flag As Boolean
Dim i As Integer
Dim sx(7) As Double
Dim sy(7) As Double
Randomize  '随机
sx(0) = 10935 * Rnd + 1400
sy(0) = 8000 * Rnd + 2000
For i = 1 To 6
 x = 10935 * Rnd + 1400
 y = 8000 * Rnd + 2000
 flag = False
 While flag = False
 Randomize
   x = 10935 * Rnd + 1400
   y = 8000 * Rnd + 2000
   flag = True
 For index = 0 To i - 1
  If Abs(sx(index) - x) < 1100 And Abs(sy(index) - y) < 1400 Then flag = False
 Next
 Wend
sx(i) = x
sy(i) = y
Next
Image2(0).Left = sx(0): Image2(0).Top = sy(0)
Image2(1).Left = sx(1): Image2(1).Top = sy(1)
Image3(0).Left = sx(2): Image3(0).Top = sy(2)
Image3(1).Left = sx(3): Image3(1).Top = sy(3)
Image4(0).Left = sx(4): Image4(0).Top = sy(4)
Image5(0).Left = sx(5): Image5(0).Top = sy(5)
Image5(1).Left = sx(6): Image5(1).Top = sy(6)
For index = 0 To 1
Image2(index).Visible = True
Image3(index).Visible = True
Image5(index).Visible = True
Next
Image4(0).Visible = True
End Sub
'选择难度
Private Sub difficult_Click()
Timer2.Interval = 600
Label4.Caption = "难度：困难"
End Sub
Private Sub easy_Click()
Timer2.Interval = 1000
Label4.Caption = "难度：简单"
End Sub
Private Sub exit_Click()
End
End Sub


Private Sub normal_Click()
Timer2.Interval = 800
Label4.Caption = "难度：一般"
End Sub
'菜单上的暂停键
Private Sub pause_Click()
If Timer1.Enabled = True Then
Command1(1).Caption = "恢复"
Timer1.Enabled = False
Timer2.Enabled = False
Else
Command1(1).Caption = "暂停"
Timer1.Enabled = True
Timer2.Enabled = True
End If
Command1(0).Enabled = True
End Sub
'移动飞艇
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyLeft
Image1.Move Image1.Left - 200, Image1.Top, Image1.Width, Image1.Height
If Image1.Left + Image1.Width <= Picture1.ScaleLeft Then
Image1.Left = Picture1.ScaleWidth - Image1.Width
End If
Case vbKeyRight
Image1.Move Image1.Left + 200, Image1.Top, Image1.Width, Image1.Height
If Image1.Left + Image1.Width / 2 >= Picture1.ScaleWidth Then
Image1.Left = 0
End If
End Select
End Sub
'初始化游戏界面
Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
Call ks
Dim index As Integer
For index = 0 To 1
Image2(index).Visible = False
Image3(index).Visible = False
Image5(index).Visible = False
Next
Image4(0).Visible = False
Picture1.Enabled = False
End Sub
'显示排行榜
Private Sub Rank_Click()
'Form1.Hide
rankshow.Show
End Sub
'菜单上的开始键
Private Sub start1_Click()
Score = 0            '初始化分数
Time.Caption = 60   '初始化时间
Picture1.Enabled = True
Timer1.Enabled = True
Timer2.Enabled = True
Command1(0).Caption = "重新开始"
Command1(1).Caption = "暂停"
Call ks
Picture1.SetFocus
Command1(0).Enabled = False
Command1(0).Enabled = True
Image2(0).Visible = True
Image2(1).Visible = True
Image3(0).Visible = True
Image3(1).Visible = True
Image5(0).Visible = True
Image5(1).Visible = True
Image4(0).Visible = True
End Sub
'使用指南
Private Sub text_Click()
msg = "可以选择：简单，一般，困难。难度键盘左右控制飞艇的方向，红色的气球加50分，爱心气球分数翻倍，紫色气球分数减半，热气球直接爆炸结束程序，时间一共60秒，时间结束显示出分数。分数高的可进入排行榜。"
Title = "                                                         操作说明                      "
MsgBox msg, vbOKOnly, Title
End Sub
'时间设置
Private Sub Timer1_Timer()
Time.Caption = Time.Caption - 1
If Time.Caption = 0 Then
EndGame
End If
End Sub
'气球移动以及判断相遇条件
Private Sub Timer2_Timer()
For k = 0 To 1
Image2(k).Top = Image2(k).Top - 500
Next
For k = 0 To 1
msg1 = Image2(k).Top >= Image1.Top And Image2(k).Top <= (Image1.Top + Image1.Height) And ((Image2(k).Left >= Image1.Left And Image2(k).Left <= (Image1.Left + Image1.Width)) Or (Image2(k).Left + Image2(k).Width >= Image1.Left And Image2(k).Left + Image2(k).Width <= (Image1.Left + Image1.Width)))
msg2 = Image2(k).Top + Image2(k).Height >= Image1.Top And Image2(k).Top + Image2(k).Height <= (Image1.Top + Image1.Height) And ((Image2(k).Left >= Image1.Left And Image2(k).Left <= (Image1.Left + Image1.Width)) Or (Image2(k).Left + Image2(k).Width >= Image1.Left And Image2(k).Left + Image2(k).Width <= (Image1.Left + Image1.Width)))

If msg1 Or msg2 Then
Timer2.Enabled = False

MsgBox "你碰到爆炸气球，游戏结束,你的得分是" & Score, vbOKOnly, "结果"
EndGame
End If
If Image2(k).Top <= 0 Then
Image2(k).Top = 10000: Image2(k).Left = Int(12000 * Rnd)
End If
Next
For k = 0 To 1
Image3(k).Top = Image3(k).Top - 400
Next
For k = 0 To 1
msg3 = Image3(k).Top >= Image1.Top And Image3(k).Top <= (Image1.Top + Image1.Height) And ((Image3(k).Left >= Image1.Left And Image3(k).Left <= (Image1.Left + Image1.Width)) Or (Image3(k).Left + Image3(k).Width >= Image1.Left And Image3(k).Left + Image3(k).Width <= (Image1.Left + Image1.Width)))
msg4 = Image3(k).Top + Image3(k).Height >= Image1.Top And Image3(k).Top + Image3(k).Height <= (Image1.Top + Image1.Height) And ((Image3(k).Left >= Image1.Left And Image3(k).Left <= (Image1.Left + Image1.Width)) Or (Image3(k).Left + Image3(k).Width >= Image1.Left And Image3(k).Left + Image3(k).Width <= (Image1.Left + Image1.Width)))
If msg3 Or msg4 Then
Score = Score / 2
Image3(k).Top = 10000
Image3(k).Left = Int(12000 * Rnd)
End If
Next
For k = 0 To 1
If Image3(k).Top <= 0 Then
Image3(k).Top = 10000
Image3(k).Left = Int(12000 * Rnd)
End If
Next
Image4(0).Top = Image4(0).Top - 300
msg5 = Image4(0).Top >= Image1.Top And Image4(0).Top <= (Image1.Top + Image1.Height) And ((Image4(0).Left >= Image1.Left And Image4(0).Left <= (Image1.Left + Image1.Width)) Or (Image4(0).Left + Image4(0).Width >= Image1.Left And Image4(0).Left + Image4(0).Width <= (Image1.Left + Image1.Width)))
msg6 = Image4(0).Top + Image4(0).Height >= Image1.Top And Image4(0).Top + Image4(0).Height <= (Image1.Top + Image1.Height) And ((Image4(0).Left >= Image1.Left And Image4(0).Left <= (Image1.Left + Image1.Width)) Or (Image4(0).Left + Image4(0).Width >= Image1.Left And Image4(0).Left + Image4(0).Width <= (Image1.Left + Image1.Width)))
If msg5 Or msg6 Then
Score = Score * 2
Image4(0).Top = 10000
Image4(0).Left = Int(12000 * Rnd)
End If
If Image4(0).Top <= 0 Then
Image4(0).Top = 10000
Image4(0).Left = Int(12000 * Rnd)
End If
For k = 0 To 1
 Image5(k).Top = Image5(k).Top - 350
 Next
 For k = 0 To 1
 msg7 = Image5(k).Top >= Image1.Top And Image5(k).Top <= (Image1.Top + Image1.Height) And ((Image5(k).Left >= Image1.Left And Image5(k).Left <= (Image1.Left + Image1.Width)) Or (Image3(k).Left + Image3(k).Width >= Image1.Left And Image3(k).Left + Image3(k).Width <= (Image1.Left + Image1.Width)))
 msg8 = Image5(k).Top + Image5(k).Height >= Image1.Top And Image5(k).Top + Image5(k).Height <= (Image1.Top + Image1.Height) And ((Image5(k).Left >= Image1.Left And Image5(k).Left <= (Image1.Left + Image1.Width)) Or (Image5(k).Left + Image5(k).Width >= Image1.Left And Image5(k).Left + Image5(k).Width <= (Image1.Left + Image1.Width)))
 If msg7 Or msg8 Then
 Score = Score + 50
 Image5(k).Top = 10000
Image5(k).Left = Int(12000 * Rnd)
End If
Next

For k = 0 To 1
If Image5(k).Top <= 0 Then
Image5(k).Top = 10000
Image5(k).Left = Int(12000 * Rnd)
End If
Next



Score1.Caption = Score
End Sub
'打开记事本
Private Sub txt_Click()
x = Shell("c:\windows\system32\notepad.exe", 1)
End Sub
'得到排行榜的最低分
Private Sub GetMinScore()
Dim temp(1 To 5) As user
    If Dir$("排行榜.txt") = "" Then
        Open "排行榜.txt" For Random As #1
        For i = 1 To 5
        temp(i).fs = 0
        Put #1, i, temp(i)
        Next i
        Close #1
        MinScore = 0
    Else
        Open "排行榜.txt" For Random As #1
        'For i = 1 To 5
            Get #1, 5, temp(5)
        'Next i
        Close #1
        MinScore = temp(5).fs
    End If
End Sub
'未进入排行榜
Private Sub EndGame1()
If Time.Caption = 0 Then      '时间结束的情况
msg = "时间到了！你的得分是 " & Score
Title = "结果"
MsgBox msg, vbOKOnly, Title
End If
Command1(0).Caption = "重新开始"
Command1(0).Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False
Call ks
Time.Caption = 60
Score1.Caption = 0
Image2(0).Visible = False
Image2(1).Visible = False
Image3(0).Visible = False
Image3(1).Visible = False
Image4(0).Visible = False
Image5(0).Visible = False
Image5(1).Visible = False
End Sub
'进入排行榜
Private Sub EndGame2()
Timer1.Enabled = False
Timer2.Enabled = False
inputname.txtfs = Score
t = Score
inputname.Show
inputname.txtname.SetFocus
'Form1.Hide
End Sub
'结束模块
Private Sub EndGame()
Command1(0).Caption = "重新开始"
GetMinScore
If Score < MinScore Then
EndGame1
Else
EndGame2

End If

End Sub
