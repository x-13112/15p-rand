VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "15puzzle�������"
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
      ToolTipText     =   "����˳�����"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   9
      ToolTipText     =   "�������"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Class"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "��Ȼ˳�����"
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
      Caption         =   "����"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�ر�"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��ʼ"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��¼"
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
      ToolTipText     =   "����˳�����"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "All"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "�������"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Class"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "��Ȼ˳�����"
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
'i,j,k��Ϊ����
'tΪѭ�����������ֵΪ32000
'wel��upd���ַ�������Ϊ127
Dim pzrand15(15) As Single, pz15(15) As Integer, stan(15) As Integer, temp1(15) As Integer, temp2(15) As Integer, _
    nxs%, nxsx(32767) As Integer, mun%, i%, j%, k%, t%, dt%, m%, n%
Dim abc(1) As Integer, wel As String * 127, upd As String * 127, sec!, ext
Sub welc()
wel = MsgBox("��ӭʹ��15puzzle������ҳ���" + vbCrLf + "1.���[Class]��[Snake]��[All]���ɴ��ң�" + _
      vbCrLf + "2.���[�ر�]����Ӧ�ã�" + vbCrLf + "3.˫���汾�ſ��ٴβ鿴��", , "ʹ��˵��")
End Sub
Sub upda()
upd = MsgBox("ver 0.0.1.191106" + vbCrLf + "����������ң����ж�������Ϊż��ʱ��ʾ�����" + _
      vbCrLf + "ver 0.1.0.191106" + vbCrLf + "1.���ڲ�ͬ������������չΪClass��Snakeģʽ��ȡ����������ʾ��" + _
      vbCrLf + "2.����ʹ��˵����", , "������־")
End Sub
Function puzzle15()                                   '����15p������������
m = 4                                                 '����
n = m ^ 2 - 1
Cls
''��������ɹ���
Randomize                                             '��ȫ���
For i = 0 To n
  pzrand15(i) = Rnd()                                 '����16�������
Next i
For i = 0 To n
  mun = 0
  For j = 0 To n
    If pzrand15(i) > pzrand15(j) Then mun = mun + 1
  Next j
  pz15(i) = mun
Next i
'''���ɽ���
''�������������
For i = 0 To n
  If Int(i / m) = 0 Then
      stan(i) = i + 1
    ElseIf Int(i / m) = 1 Then
      stan(i) = (2 * m ^ 2 - m - i) Mod 16
    ElseIf Int(i / m) = 2 Then
      stan(i) = i + 1
    Else
      stan(i) = 2 * m ^ 2 - m - i                     '-1��Ϊ��ȥ������0
  End If
Next i
For i = 0 To n
  temp1(i) = pz15((stan(i) + 15) Mod 16)              '�����±�Խ��
Next i
j = 0
Do Until temp1(j) = 0                                 '����0��λ��
  j = j + 1
Loop
i = j
Do Until i = n
  temp1(i) = temp1(i + 1)
  i = i + 1
Loop
For i = m ^ 2 - m To n - 1
  stan(i) = stan(i + 1)                               '������׼���У�ȥ��0��m^2
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
'''�������
End Function
Sub prn15p()                                         '���15p
i = 0
Do While i <= n
  If pz15(i) >= 10 Then
      Print pz15(i);
      Print vbTab;                                   'tab�Ű�
    ElseIf pz15(i) > 0 Then
      Print Space(1);
      Print pz15(i);
      Print vbTab;                                   'tab�Ű�
    Else
      Print vbTab;                                   '����0
  End If
  i = i + 1
  If i Mod m = 0 Then Print                          '������������ֺ���
Loop
End Sub
Sub stat()                                           '����
sec = Format(0#)
sec = InputBox("������������" + vbCrLf + "˵����" + vbCrLf + "1.������ʵ������" + _
      vbCrLf + "2.���������0�����֣�С�ڵ���0���������ַ���Ч", "����")
Do Until sec > 0
  sec = InputBox("�����������" + vbCrLf + "����������������" + vbCrLf + "˵����" + _
        vbCrLf + "���������0�����֣�С�ڵ���0���������ַ���Ч", "����")
Loop
ext = InputBox("�����뷣�룺" + vbCrLf + "˵����" + vbCrLf + "1.��ȫ��ԭ������0" + _
      vbCrLf + "2.��ȫ��ԭ�����һ�����ֿ�����ĵ�δ��λ������2" + _
      vbCrLf + "3.δ���������DNF" + vbCrLf + "4.���������ַ���Ч", "����")
inputext:
  If ext = 2 Or ext = 0 Then
      MsgBox "�˴����ճɼ�Ϊ��", sec + ext
    ElseIf ext = DNF Then
      MsgBox "�˴����ճɼ�Ϊ��", ext
    Else
      ext = InputBox("�����������" + vbCrLf + "���������뷣�룺" + vbCrLf + "˵����" + _
            vbCrLf + "1.��ȫ��ԭ������0" + vbCrLf + "2.��ȫ��ԭ�����һ�����ֿ�����ĵ�δ��λ������2" + _
            vbCrLf + "3.δ���������DNF" + vbCrLf + "4.���������ַ���Ч", "����")
      GoTo inputext
  End If
End Sub
Private Sub Command1_Click()
Class:
  puzzle15
If nxs Mod 2 = 1 Then GoTo Class
Print "ģʽ��class"
  prn15p
End Sub
Private Sub Command2_Click()
  puzzle15
If nxs Mod 2 = 0 Then
    Print "ģʽ��class"
  Else
    Print "ģʽ��snake"
  End If
  prn15p
End Sub
Private Sub Command3_Click()
snake:
  puzzle15
If nxs Mod 2 = 0 Then GoTo snake
Print "ģʽ��snake"
  prn15p
End Sub
Private Sub Command4_Click()                             '����
  stat

'abc(0) = InputBox("��������Сֵ��", "��������Χɸѡ", 40)
'abc(1) = InputBox("���������ֵ��", "��������Χɸѡ", 70)
'MsgBox ("���������������ΧΪ��" + abc(0) + "--" + abc(1),vbOKOnly)
End Sub
Private Sub Command5_Click()                             '����
If sec = 0 And Command5.Caption = "��ʼ" Then
  Timer1.Enabled = True
  Timer1.Interval = 1
  Command5.Caption = "ֹͣ"
End If
If sec = 10 Then
End If
If Command5.Caption = "ֹͣ" Then
  Timer1.Enabled = False
  Timer1.Interval = 0
  Label1.Caption = sec
  Command5.Caption = "��ʼ"
End If
End Sub
Private Sub Command6_Click()
  End
End Sub
Private Sub Command7_Click()
Cls
Label1.Caption = "ver: 0.1.0.191107"                     '�汾��
'dt = InputBox("ѭ������(���ֵΪ32000������С��10000)��", , 1000)
'For t = 1 To dt
'  Call Command1_Click
'  nxsx(t) = nxs
'Next t
'Open "Reverse order num.csv" For Output As #1
'For t = 1 To dt
'  Print #1, nxsx(t)
'Next t
'Close #1
'MsgBox "ѭ���ѽ�����" + vbCrLf + "���ݱ��浽Reverse order num.csv"
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
If sec >= 10 Then Timer1.Interval = 0                    'sec��ʱ
End Sub
