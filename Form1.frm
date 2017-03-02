VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "无名益智小游戏"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10965
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "游戏简介"
      Height          =   495
      Left            =   9000
      TabIndex        =   25
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新开始"
      Height          =   495
      Left            =   9000
      TabIndex        =   24
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Height          =   495
      Left            =   9000
      TabIndex        =   23
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   4
      Left            =   2760
      TabIndex        =   7
      Top             =   6120
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3960
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   3030
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   2100
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   1170
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   4800
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   2160
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   1200
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "|"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "用户"
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "电脑"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "玩家"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xxxx, i As Integer '上次选中个数
 
Private Sub Check1_Click(index As Integer)
 '第一行
Dim num As Integer
 
If win(0) = 0 And Option1(1).Value = True Then MsgBox "亲，你输了！", 64, "游戏结束"
If index = 0 Then
   Dim i1 As Integer
   '选中
   If Check1(index).Value = 1 Then
        
        For i1 = 1 To 14
          Check1(i1).Enabled = False
        Next
    Else
     Call jinyong(1)
    End If
'第二行
ElseIf index >= 1 And index <= 2 Then
    Dim i2 As Integer
    For i2 = 1 To 2
        If Check1(i2).Value = 1 And Check1(i2).Enabled = True Then
        '复选框选中的个数
          num = num + 1
        End If
    Next
   
    If num >= 1 Then
         
     '有一个选中
        Call check(1, 2, 0)
    ElseIf num = 0 Then
    '全未选中
         Call jinyong(1)
    End If
'第三行
ElseIf index >= 3 And index <= 5 Then
    Dim i3 As Integer
    For i3 = 3 To 5
        If Check1(i3).Value = 1 And Check1(i3).Enabled = True Then
          num = num + 1
        End If
    Next
    If num >= 1 Then
        Call check(3, 5, 0)
    ElseIf num = 0 Then
          Call jinyong(1)
    End If
'第四行
ElseIf index >= 6 And index <= 9 Then
    Dim i4 As Integer
    For i4 = 6 To 9
        If Check1(i4).Value = 1 And Check1(i4).Enabled = True Then
          num = num + 1
        End If
    Next
    If num >= 1 Then
           Call check(6, 9, 0)
    ElseIf num = 0 Then
          Call jinyong(1)
    End If
'第五行
ElseIf index >= 10 And index <= 14 Then
    Dim i5 As Integer
    For i5 = 10 To 14
        If Check1(i5).Value = 1 And Check1(i5).Enabled = True Then
          num = num + 1
        End If
    Next
    If num >= 1 Then
           Call check(10, 14, 0)
    ElseIf num = 0 Then
         Call jinyong(1)
    End If
End If
End Sub

Private Sub Command1_Click()
If win(1) = i And win(1) <> 15 Or win(0) = 0 Then MsgBox "亲，该你走啦！", 48, "提示": Exit Sub
 '更改状态显示
Option1(0).Enabled = True
Option1(0).Value = True
Option1(1).Enabled = False
'禁用用户操作
Call jinyong(0)
'电脑分析开始
'Check1(5).Value = 1
Call compute(0) '电脑分析 0傻瓜级别
Call jinyong(1)
'电脑分析结束更改状态
Option1(1).Enabled = True
Option1(1).Value = True
Option1(0).Enabled = False
i = win(1)
End Sub

Private Sub Command2_Click()
i = 0
Call Form_Load
End Sub

Private Sub Command3_Click()
Load Form2
Form2.Show
End Sub

Private Sub Form_Activate()

''获取不定个数随机数不能重复
 
End Sub

Private Sub Form_DblClick()
'Dim x
'x = row_one()
'If Not IsEmpty(x) And x <> 0 Then Print x
End Sub


'窗体获得焦点事件
Private Sub Form_GotFocus()

End Sub

Private Sub Form_Load()
 
 '初始化
Option1(1).Value = True
 Option1(0).Enabled = False
Dim i As Integer

For i = 0 To 14
 
  Check1(i).Enabled = True
  Check1(i).Value = 0
  Check1(i).Caption = "|"
Next
End Sub

'电脑分析
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
 '禁用第一个复选框
 Dim i, num
 For i = 0 To 14
    If Check1(i).Value = 1 Then num = num + 1
 Next i
 If num = 0 Then Check1(0).Enabled = False
End Sub

Private Sub Frame1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'Print Index
End Sub
 
'累加
Public Function add(a As Integer) As Integer
add = a + 1
End Function
'打印数组
Public Sub dump(a, b As Integer)
Dim i, j As Integer
If b = 1 Then
    For i = LBound(a) To UBound(a)
        Print a(i);
    Next i
ElseIf b = 2 Then
    For i = LBound(a, 1) To UBound(a, 1)
        For j = LBound(a, 2) To UBound(a, 2)
         Print a(i, j);
        Next j
        Print " "
    Next i
End If
End Sub
'求数组的长度
Private Function arr_count(a)
arr_count = UBound(a) - LBound(a) + 1
End Function
'获取禁用的复选框索引
Private Sub check(a As Integer, b As Integer, c As Integer)
  Dim i As Integer
    For i = 0 To a - 1
    If c = 0 Then
        Check1(i).Enabled = False
         Else
        Check1(i).Enabled = True
    End If
    Next
    For i = b + 1 To 14
    If c = 0 Then
        Check1(i).Enabled = False
        Else
        Check1(i).Enabled = True
    End If
    Next
End Sub

'禁用复选框
Private Sub jinyong(a As Integer)
Dim i As Integer
For i = 0 To 14
    If a = 0 Then '全部禁用
        Check1(i).Enabled = False
    ElseIf a = 1 Then '打开未被选中的
        If Check1(i).Value <> 1 Then
             Check1(i).Enabled = True
         End If
    ElseIf a = 2 Then
        If Check1(i).Value = 1 Then
             Check1(i).Enabled = False
         End If
    End If
Next
End Sub
'最后下的一方
Private Function win(a As Integer)
Dim i, num1, num2, tmp As Integer
For i = 0 To 14
If Check1(i).Value = 1 Then num1 = num1 + 1 Else num2 = num2 + 1
Next
If a = 1 Then
tmp = num1
ElseIf a = 0 Then
tmp = num2
End If
win = tmp
End Function
'电脑思考
Private Sub compute(a As Integer)
If a = 0 Then Call fool '傻瓜级别
End Sub
'获取未选中的复选框存入二维数组
Private Function arr_check()

Dim arr(4, 4), i, j As Integer
For i = 0 To 14
    If Check1(i).Value <> 1 Then
        If i = 0 Then
        arr(0, i) = 0 ' 此行特殊处理
        ElseIf i >= 1 And i <= 2 Then
        arr(1, i - 1) = i
        ElseIf i >= 3 And i <= 5 Then
        arr(2, i - 3) = i
        ElseIf i >= 6 And i <= 9 Then
        arr(3, i - 6) = i
        ElseIf i >= 10 And i <= 14 Then
        arr(4, i - 10) = i
        End If
    End If
Next
arr_check = arr
End Function

 
 

Private Sub test1()
Dim num, i As Integer
'Randomize
'num = Int(Rnd * 3 + 1) '随机次数
'Print num
For i = 0 To 3
Print i
Next
End Sub
'去重复
Function Array_unique(arr As Variant) As Variant
    arr = QuickSort(arr)
    Dim k As Integer, i As Integer

    For i = 0 To UBound(arr)
        If arr(k) <> arr(i) Then
            arr(k + 1) = arr(i)
            k = k + 1
        End If

    Next
    Dim NewArr() As Variant
    ReDim NewArr(k)
    For i = 0 To k
        NewArr(i) = arr(i)
    Next

    Array_unique = NewArr
End Function
'排序
Function QuickSort(arr)
    Dim i, j
    Dim bound, t
    bound = UBound(arr)

    For i = 0 To bound - 1
        For j = i + 1 To bound
            If arr(i) > arr(j) Then
                t = arr(i)
                arr(i) = arr(j)
                arr(j) = t
            End If
        Next
    Next
    QuickSort = arr
End Function

Sub test()
    Dim s
    s = Array_unique(Array(1, 2, 2))
    'Debug.Print Join(s, "|")
    Debug.Print arr_count(s)
End Sub
'数组中非空元素个数
Private Function arr_num(a)
    Dim i, num As Integer
    For i = LBound(a) To UBound(a)
        If TypeName(a(i)) <> "Empty" Then num = num + 1
    Next
    arr_num = num
End Function

Private Function arr_new(arr, k, z) 'k为行号 arr为二维数组
''获取不定个数随机数不能重复

Dim num, i, j, tmp(5), test As Integer
Dim arr_tmp, arr_one(5)
'将二维数组的一维赋给变量
    Do
        Randomize
        If z = 0 Then
            If remain_list(k) = 0 And remain_one(k) <> 1 Then
             num = remain_one(k) - 1
            Else
             num = Int(Rnd * remain_one(k) + 1) '随机次数
            End If
        Else
            num = z   '随机次数
        End If
    
        For i = 1 To num
           test = Int(Rnd * 5)
           tmp(i - 1) = arr(k, test)
        Next
        '比较是否有重复
        arr_tmp = Array_unique(tmp)
    Loop While arr_num(arr_tmp) <> num
    arr_new = arr_tmp
 
End Function
'傻瓜级别随机选中
Private Sub fool()
If win(0) = 0 Then MsgBox "无路可走 ", 48, "提示": Exit Sub
Dim num, i, arr1, arr2, row_num, tmp 'row_num 为剩余行数
Dim tmp_arr, tmp_arr1, tmp_arr3, tmp_arr4, status, tmp_num '当前未选中的行
Dim tmp_arr5, arr_tmp(1), num_tmp, max, tmp_num1
Dim rnd_sum, tmp_sum, max_id, tmp_arr13, tmp_arr14, max2
Dim mid, num2, min, min2, tmp_arr114, tmp_arr24
rnd_sum = 0
Randomize
row_num = arr_num(arr_one(remain_all(), 0)) '剩余行数
tmp_arr3 = Array_unique(arr_one(remain_all(), 0)) '剩余行去重复
tmp_arr13 = tmp_arr3
max = remain_one(tmp_arr3(0))
tmp_arr4 = arr_one(remain_all(), 1) '剩余数量
tmp_arr14 = tmp_arr4

min = arr_min(tmp_arr4)
min2 = arr_key(tmp_arr4, min) '目标行号
max2 = arr_max(tmp_arr4)
num2 = arr_key(tmp_arr14, max2) '目标行号
tmp_arr114 = Array_unique(tmp_arr4)
'只选中一个
If Not IsEmpty(row_one()) Then
    If row_one() = 1 Then
        num = 2: rnd_sum = 2: GoTo n200
    ElseIf row_one() = 2 Then
        num = 3: rnd_sum = 1: GoTo n200
    ElseIf row_one() = 3 Then
        num = 4: rnd_sum = 2: GoTo n200
    ElseIf row_one() = 4 Then
     num = 4: rnd_sum = 1: GoTo n200
    End If
End If
If row_num = 2 And arr_num(tmp_arr3) = arr_num(tmp_arr114) Then
'剩两行时
    '1
    '1111
    If arr_min(tmp_arr14) = 1 Then
        num = arr_key(tmp_arr14, arr_max(tmp_arr14))
        rnd_sum = arr_max(tmp_arr14)
        GoTo n200
     
    ElseIf remain_one(tmp_arr3(1)) > remain_one(tmp_arr3(2)) Then
         num = tmp_arr3(1): rnd_sum = remain_one(tmp_arr3(1)) - remain_one(tmp_arr3(2))
        GoTo n200
    ElseIf remain_one(tmp_arr3(1)) < remain_one(tmp_arr3(2)) Then
         num = tmp_arr3(2): rnd_sum = remain_one(tmp_arr3(2)) - remain_one(tmp_arr3(1))
        GoTo n200
    End If
End If
'1          1
'1         1
'1         11
'1111     1
'1
If row_num > 2 And arr_num(Array_unique(tmp_arr4)) = 2 And arr_num(tmp_arr3) > 2 Then
  '仅有一行剩余选项数不为1，其余行剩余为1
     For i = LBound(tmp_arr3) To UBound(tmp_arr3)
         If remain_one(tmp_arr3(i)) = 1 Then num_tmp = num_tmp + 1
         If remain_one(tmp_arr3(i)) > max Then
            max = remain_one(tmp_arr3(i)): max_id = tmp_arr3(i)
         End If
    Next
    If num_tmp = arr_num(tmp_arr3) - 1 Then
        If row_num Mod 2 = 0 Then
            rnd_sum = remain_one(max_id)
        Else
            rnd_sum = remain_one(max_id) - 1
        End If
        num = max_id: GoTo n200
    End If
End If

'剩余3行时
'11
'11
'1111
If row_num = 3 And arr_num(tmp_arr114) = 2 Then
    For i = LBound(tmp_arr14) To UBound(tmp_arr14)
       If tmp_arr14(i) > 1 Then num_tmp = num_tmp + 1
       If repet_num(tmp_arr14, tmp_arr14(i)) = 1 Then
            num = i
            rnd_sum = remain_one(num)
       End If
    Next
    If num_tmp >= 2 Then GoTo n200
End If
'剩余3行时
'11
'1111
'11111
If row_num = 3 And arr_num(tmp_arr114) = 3 And arr_max(tmp_arr14) = 5 And arr_min(tmp_arr14) > 1 Then
tmp_arr24 = tmp_arr14
For i = 0 To 4
        min = min - 1
        tmp_arr24(min2) = min
        If min = 1 And arr_sum(tmp_arr24) = 10 Then
            rnd_sum = i + 1: num = min2: GoTo n200
        End If
    Next i
End If
'剩余3行时
'1
'11
'11111
If row_num = 3 And arr_num(tmp_arr114) = 3 And max2 > 3 Then
   tmp_arr24 = tmp_arr14
     For i = 0 To 4
        max2 = max2 - 1
        tmp_arr24(num2) = max2
        mid = arr_sum(tmp_arr24) - arr_max(tmp_arr24) - arr_min(tmp_arr24)
        If arr_max(tmp_arr24) - mid = 1 And mid - arr_min(tmp_arr24) = 1 And arr_min(tmp_arr24) = 1 Then
            rnd_sum = i + 1: num = num2: GoTo n200
        End If
    Next i
End If
'剩余3行时
'11
'11
'11
If row_num = 3 And arr_num(tmp_arr114) = 1 And arr_min(tmp_arr4) >= 2 Then
    rnd_sum = max2
End If
'剩余4行时
'11
'11
'1
'1111
If row_num = 4 And arr_num(arr_repeat(tmp_arr14)) = 2 Then
    rnd_sum = arr_max(arr_repeat(tmp_arr14)) - arr_min(arr_repeat(tmp_arr14))
    num = arr_key(tmp_arr14, arr_max(arr_repeat(tmp_arr14))): GoTo n200
End If
'剩余4行
'11
'11
'11
'111
If row_num = 4 And arr_num(arr_repeat(tmp_arr14)) = 1 And arr_min(tmp_arr14) >= 2 And arr_max(arr_repeat(tmp_arr14)) = arr_max(tmp_arr14) Then
    rnd_sum = arr_max(tmp_arr14) - arr_min(tmp_arr14)
    num = num2: GoTo n200
End If
'剩余4行
'1
'11
'111
'1111
If row_num = 4 And arr_num(tmp_arr114) = 4 Then
    tmp_num = 0
    For Each i In Array(1, 2, 3, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 4 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 4
        GoTo n200
    End If
    tmp_num = 0
    For Each i In Array(1, 2, 3, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 4 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 5
        GoTo n200
    End If
    tmp_num = 0
    For Each i In Array(1, 2, 4, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 4 Then
        num = arr_key(tmp_arr14, 2)
        rnd_sum = 2
        GoTo n200
    End If
    tmp_num = 0
    For Each i In Array(1, 3, 4, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 4 Then
        num = arr_key(tmp_arr14, 3)
        rnd_sum = 3
        GoTo n200
    End If
   
End If
'剩余5行
'1
'1
'11
'11
'1111
If row_num = 5 And arr_num(arr_repeat(tmp_arr14)) = 1 And repet_num(tmp_arr14, 1) <> 4 Then
rnd_sum = arr_max(arr_repeat(tmp_arr14)) '目标数量
num = arr_key(tmp_arr14, rnd_sum) '目标行号
GoTo n200
End If
'剩余5行
'1
'1
'1
'11
'11
If row_num = 5 And arr_num(tmp_arr114) = 2 Then
For Each i In tmp_arr114
    If repet_num(tmp_arr14, i) = 3 Then
        num = arr_max(arr_key(tmp_arr14, i))
        rnd_sum = i
        GoTo n200
    End If
Next i
End If
'Print row_num
'剩余5行Array(1, 2, 3, 1, 5)
If row_num = 5 Then
    tmp_num = 0
    For Each i In Array(2, 3, 5)
       If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 1) = 2 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 4
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 2, 3, 1, 4)
    For Each i In Array(2, 3, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 1) = 2 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 3
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 2, 3, 2, 5)
    For Each i In Array(1, 3, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 2) = 2 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 3
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 2, 3, 3, 5)
    For Each i In Array(1, 2, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 3) = 2 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 2
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 2, 3, 3, 4)
    For Each i In Array(1, 2, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 3) = 2 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 1
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 2, 3, 4, 4)
    For Each i In Array(1, 2, 3)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 4) = 2 Then
        num = arr_max(arr_key(tmp_arr14, 4))
        rnd_sum = 3
        GoTo n200
    End If
      tmp_num = 0
    'Array(1, 1, 1, 3, 5)
    For Each i In Array(3, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 1) = 3 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 3
        GoTo n200
    End If
    tmp_num = 0
    'Array(1, 1, 1, 4, 5)
    For Each i In Array(4, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 1) = 3 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 2
        GoTo n200
    End If
    
      tmp_num = 0
    'Array(1, 1, 1, 2, 5)
    For Each i In Array(2, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 1) = 3 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 2
        GoTo n200
    End If
     tmp_num = 0
     'Array(1, 1, 1, 3, 4)
    For Each i In Array(3, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 1) = 3 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 2
        GoTo n200
    End If
     tmp_num = 0
     'Array(2,2,2,1,5)
    For Each i In Array(1, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 2) = 3 Then
        num = arr_key(tmp_arr14, 5)
        rnd_sum = 2
        GoTo n200
    End If
     tmp_num = 0
     'Array(2,2,2,1,4)
    For Each i In Array(1, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 2 And repet_num(tmp_arr14, 2) = 3 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 1
        GoTo n200
    End If
     tmp_num = 0
    'Array(1,1,3,4,5)
    For Each i In Array(3, 4, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 1) = 2 Then
        num = arr_key(tmp_arr14, 3)
        rnd_sum = 2
        GoTo n200
    End If
     tmp_num = 0
    'Array(1,1,2,4,5)
    For Each i In Array(2, 4, 5)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 1) = 2 Then
        num = arr_key(tmp_arr14, 2)
        rnd_sum = 1
        GoTo n200
    End If
      tmp_num = 0
    'Array(1,2,2,3,4)
    For Each i In Array(1, 3, 4)
        If repet_num(tmp_arr14, i) = 1 Then
            tmp_num = tmp_num + 1
        End If
    Next i
    If tmp_num = 3 And repet_num(tmp_arr14, 2) = 2 Then
        num = arr_key(tmp_arr14, 4)
        rnd_sum = 2
        GoTo n200
    End If
   
     
End If
num = rnd_row() '随机行数
n200:
arr2 = arr_new(arr_check(), num, rnd_sum) '待选中二维数组
For i = LBound(arr2) To UBound(arr2)
   If TypeName(arr2(i)) = "Empty" Or TypeName(i) = "Empty" Then GoTo n100
        Check1(arr2(i)).Value = 1
        Check1(arr2(i)).Caption = "||"
n100:
Next i
Call jinyong(2)
If win(0) = 0 Then MsgBox "好吧，我输了！", 64, "游戏结束"
End Sub
'每行剩余数
Private Function remain_one(ByVal a As Integer)
    Dim i, num As Integer
    For i = 0 To 14
        If Check1(i).Value <> 1 Then
            If a = 0 And i = 0 Then
                num = num + 1
            ElseIf a = 1 And i >= 1 And i <= 2 Then
                num = num + 1
            ElseIf a = 2 And i >= 3 And i <= 5 Then
                num = num + 1
            ElseIf a = 3 And i >= 6 And i <= 9 Then
                num = num + 1
            ElseIf a = 4 And i >= 10 And i <= 14 Then
                num = num + 1
            End If
        End If
    Next i
    remain_one = num
End Function
'除当前行外的所有剩余复选框个数
Private Function remain_list(ByVal a As Integer)
    Dim i, num As Integer
    For i = 0 To 4
        If i <> a Then num = num + remain_one(i)
    Next
    remain_list = num
End Function
'可选的行，可选的个数 返回二维数组
Private Function remain_all()
Dim i, arr(4, 4), num, tmp
    For i = 0 To 4
        tmp = remain_one(i)
        If tmp > 0 Then
        arr(0, i) = i
        arr(1, i) = tmp
        num = num + 1
        End If
    Next
If num = 0 Then
remain_all = 0
Else
remain_all = arr
End If
End Function
'二维取一维
Private Function arr_one(a, k)
Dim arr(100)
For i = LBound(a, 2) To UBound(a, 2)
    arr(i) = a(k, i)
Next i
'arr_one = Array_unique(arr)
arr_one = arr
End Function

'生成可用随机行数
Private Function rnd_row()
Dim num
Do
num = Int(Rnd * 5)  '随机行号
Loop While remain_one(num) = 0
rnd_row = num
End Function
'根据值键名
Private Function arr_key(arr, val)
    Dim tmp(100), num, key
    For i = LBound(arr) To UBound(arr)
        If arr(i) = val Then
             num = num + 1: tmp(num) = i
            If num = 1 Then key = i
        End If
    Next i
    If num = 1 Then
        arr_key = key
    Else
        arr_key = tmp
    End If
End Function

'重复次数
Private Function repet_num(arr, val)
Dim num, j
For j = LBound(arr) To UBound(arr)
     If arr(j) = val And Not IsEmpty(arr(j)) Then num = num + 1
Next j
repet_num = num
End Function

 

'获取数组中最大值
Private Function arr_max(arr)
    Dim max, j
    max = arr(LBound(arr))
    For j = LBound(arr) To UBound(arr)
        If Not IsEmpty(arr(j)) And arr(j) > max Then max = arr(j)
    Next j
    arr_max = max
End Function

'求最小值
Private Function arr_min(arr)
Dim min, j
min = 10
    For j = LBound(arr) To UBound(arr)
        If Not IsEmpty(arr(j)) And arr(j) < min Then min = arr(j)
    Next j
arr_min = min
End Function

'数组求和
Private Function arr_sum(arr)
Dim j, num
    For j = LBound(arr) To UBound(arr)
         num = num + arr(j)
    Next j
arr_sum = num
End Function

'统计数组中不重复的元素
Private Function arr_repeat(arr)
Dim i, j, num, tmp
Dim Count(10)
tmp = arr
For Each i In tmp
    num = 0
    For Each j In arr
        If i = j Then
            num = num + 1
        End If
    Next j
    If num = 1 Then Count(i) = i
Next i
arr_repeat = Count
End Function

'判断行号
Private Function row_one()
Dim i, num, index, tmp
For i = 0 To 14
        If Check1(i).Value = 1 Then
        num = num + 1
        tmp = i
        End If
Next i
If num = 1 Then
    If tmp = 0 Then
    row_one = 0
    ElseIf tmp >= 1 And tmp <= 2 Then
    row_one = 1
    ElseIf tmp >= 3 And tmp <= 5 Then
    row_one = 2
    ElseIf tmp >= 6 And tmp <= 9 Then
    row_one = 3
    ElseIf tmp >= 10 And tmp <= 14 Then
    row_one = 4
    End If
End If
End Function
