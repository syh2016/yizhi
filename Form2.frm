VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "游戏简介"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "楷体"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10965
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label8 
      Caption         =   "版本:A15"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   7680
      Width           =   3855
   End
   Begin VB.Label Label7 
      Caption         =   "开发时间:2016-9-23"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   6960
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "作者：宋宇航"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   6120
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   4800
      Width           =   7455
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Width           =   7455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   7455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   7455
   End
   Begin VB.Label Label1 
      Caption         =   "游戏简介"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Label1.Caption = "游戏简介"
    Label2.Caption = "玩家与电脑轮流走棋(选择复选框)"
    Label3.Caption = "每次走棋最多只能全选一行"
    Label4.Caption = "棋盘填满为止"
    Label5.Caption = "填满棋盘的一方判负"
End Sub

