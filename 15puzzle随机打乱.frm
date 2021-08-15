VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "15puzzle随机打乱"
   ClientHeight    =   2895
   ClientLeft      =   5040
   ClientTop       =   3390
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Snake"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      ToolTipText     =   "蛇形顺序打乱"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "随机打乱"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Class"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "自然顺序打乱"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2280
      Top             =   0
   End
   Begin VB.CommandButton Command7 
      Caption         =   "待定"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "关闭"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "开始"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "记录"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Snake"
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "蛇形顺序打乱"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "随机打乱"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Class"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "自然顺序打乱"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ver:0.1.0.191109"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Author: x-13112"
      Top             =   1080
      Width           =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'i,j,k均为变量
't为循环次数，最大值为32000
'wel和upd的字符串长度为127
Dim pzrand15(15) As Single, pz15(15) As Integer, stan(15) As Integer, temp1(15) As Integer, temp2(15) As Integer, _
    nxs%, nxsx(32767) As Integer, mun%, i%, j%, k%, t%, dt%, m%, n%
Dim abc(1) As Integer, wel As String * 127, upd As String * 127, sec!, ext
Sub welc()
wel = MsgBox("欢迎使用15puzzle随机打乱程序！" + vbCrLf + "1.点击[Class]、[Snake]或[All]生成打乱；" + _
      vbCrLf + "2.点击[关闭]结束应用；" + vbCrLf + "3.双击版本号可再次查看。", , "使用说明")
End Sub
Sub upda()
upd = MsgBox("ver 0.0.1.191106" + vbCrLf + "生成随机打乱，并判定逆序数为偶数时显示结果。" + _
      vbCrLf + "ver 0.1.0.191106" + vbCrLf + "1.对于不同的逆序数，扩展为Class和Snake模式，取消逆序数显示；" + _
      vbCrLf + "2.增加使用说明。", , "更新日志")
End Sub
Function puzzle15()                                   '生成15p并计算逆序数
m = 4                                                 '阶数
n = m ^ 2 - 1
Cls
''随机数生成过程
Randomize                                             '完全随机
For i = 0 To n
  pzrand15(i) = Rnd()                                 '生成16个随机数
Next i
For i = 0 To n
  mun = 0
  For j = 0 To n
    If pzrand15(i) > pzrand15(j) Then mun = mun + 1
  Next j
  pz15(i) = mun
Next i
'''生成结束
''逆序数计算过程
For i = 0 To n
  If Int(i / m) = 0 Then
      stan(i) = i + 1
    ElseIf Int(i / m) = 1 Then
      stan(i) = (2 * m ^ 2 - m - i) Mod 16
    ElseIf Int(i / m) = 2 Then
      stan(i) = i + 1
    Else
      stan(i) = 2 * m ^ 2 - m - i                     '-1是为了去掉数字0
  End If
Next i
For i = 0 To n
  temp1(i) = pz15((stan(i) + 15) Mod 16)              '避免下标越界
Next i
j = 0
Do Until temp1(j) = 0                                 '计算0的位置
  j = j + 1
Loop
i = j
Do Until i = n
  temp1(i) = temp1(i + 1)
  i = i + 1
Loop
For i = m ^ 2 - m To n - 1
  stan(i) = stan(i + 1)                               '调整标准序列，去掉0或m^2
Next i
For i = 0 To n - 1
  j = 0
  Do Until temp1(i) = stan(j)
    j = j + 1
  Loop
  temp2(i) = j
Next i
For i = 0 To n - 1
  k = 0
  For j = 0 To n - i - 1
    If temp2(i) > temp2(i + j) Then k = k + 1
  Next j
  temp1(i) = k
Next i
nxs = 0
For i = 0 To n - 1
  nxs = nxs + temp1(i)
Next i
'''计算结束
End Function
Sub prn15p()                                         '输出15p
i = 0
Do While i <= n
  If pz15(i) >= 10 Then
      Print pz15(i);
      Print vbTab;                                   'tab排版
    ElseIf pz15(i) > 0 Then
      Print Space(1);
      Print pz15(i);
      Print vbTab;                                   'tab排版
    Else
      Print vbTab;                                   '屏蔽0
  End If
  i = i + 1
  If i Mod m = 0 Then Print                          '输入阶数个数字后换行
Loop
End Sub
Sub stat()                                           '测试
sec = Format(0#)
sec = InputBox("请输入秒数：" + vbCrLf + "说明：" + vbCrLf + "1.请输入实际秒数" + _
      vbCrLf + "2.请输入大于0的数字，小于等于0或者其他字符无效", "秒数")
Do Until sec > 0
  sec = InputBox("秒数输入错误！" + vbCrLf + "请重新输入秒数：" + vbCrLf + "说明：" + _
        vbCrLf + "请输入大于0的数字，小于等于0或者其他字符无效", "错误！")
Loop
ext = InputBox("请输入罚秒：" + vbCrLf + "说明：" + vbCrLf + "1.完全还原请输入0" + _
      vbCrLf + "2.完全还原但最后一个数字块的中心点未归位请输入2" + _
      vbCrLf + "3.未完成请输入DNF" + vbCrLf + "4.输入其他字符无效", "罚秒")
inputext:
  If ext = 2 Or ext = 0 Then
      MsgBox "此次最终成绩为：", sec + ext
    ElseIf ext = DNF Then
      MsgBox "此次最终成绩为：", ext
    Else
      ext = InputBox("罚秒输入错误！" + vbCrLf + "请重新输入罚秒：" + vbCrLf + "说明：" + _
            vbCrLf + "1.完全还原请输入0" + vbCrLf + "2.完全还原但最后一个数字块的中心点未归位请输入2" + _
            vbCrLf + "3.未完成请输入DNF" + vbCrLf + "4.输入其他字符无效", "错误！")
      GoTo inputext
  End If
End Sub
Private Sub Command1_Click()
Class:
  puzzle15
If nxs Mod 2 = 1 Then GoTo Class
Print "模式：class"
  prn15p
End Sub
Private Sub Command2_Click()
  puzzle15
If nxs Mod 2 = 0 Then
    Print "模式：class"
  Else
    Print "模式：snake"
  End If
  prn15p
End Sub
Private Sub Command3_Click()
snake:
  puzzle15
If nxs Mod 2 = 0 Then GoTo snake
Print "模式：snake"
  prn15p
End Sub
Private Sub Command4_Click()                             '测试
  stat

'abc(0) = InputBox("请输入最小值：", "逆序数范围筛选", 40)
'abc(1) = InputBox("请输入最大值：", "逆序数范围筛选", 70)
'MsgBox ("您输入的逆序数范围为：" + abc(0) + "--" + abc(1),vbOKOnly)
End Sub
Private Sub Command5_Click()                             '测试
If sec = 0 And Command5.Caption = "开始" Then
  Timer1.Enabled = True
  Timer1.Interval = 1
  Command5.Caption = "停止"
End If
If sec = 10 Then
End If
If Command5.Caption = "停止" Then
  Timer1.Enabled = False
  Timer1.Interval = 0
  Label1.Caption = sec
  Command5.Caption = "开始"
End If
End Sub
Private Sub Command6_Click()
  End
End Sub
Private Sub Command7_Click()
Cls
Label1.Caption = "ver: 0.1.0.191107"                     '版本号
'dt = InputBox("循环次数(最大值为32000，建议小于10000)：", , 1000)
'For t = 1 To dt
'  Call Command1_Click
'  nxsx(t) = nxs
'Next t
'Open "Reverse order num.csv" For Output As #1
'For t = 1 To dt
'  Print #1, nxsx(t)
'Next t
'Close #1
'MsgBox "循环已结束！" + vbCrLf + "数据保存到Reverse order num.csv"
End Sub
Private Sub Form_Load()
  welc
End Sub
Private Sub Label1_DblClick()
  welc
End Sub
Private Sub Timer1_Timer()
Dim sec1!, sec2!
Do Until Timer1.Interval = 0
  sec1 = Time
Loop
sec = (sec2 - sec1) * 86400
If Timer1.Interval = 1 Then sec2 = Time
  Label1.Caption = Format(sec, "0.000")
If sec >= 10 Then Timer1.Interval = 0                    'sec超时
End Sub
