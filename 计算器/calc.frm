VERSION 5.00
Begin VB.Form calculator 
   Caption         =   "简单计算器"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9690
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   19
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CommandButton Calc 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Clear 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Exit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   12
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "开方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      TabIndex        =   17
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   "乘方"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "数2"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "数1"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Clear.Enabled = False
Calc.Enabled = False
Label9.Caption = ""
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label9.Caption = ""
Clear.Enabled = False
Calc.Enabled = True
End Sub

Private Sub Calc_Click()
If Len(Text1.Text) = 0 Or Len(Text2.Text) = 0 Then
Label9.Caption = "请先输入数字"
ElseIf Text2.Text <> 0 Then
Text3.Text = Val(Text1.Text) + Val(Text2.Text)
Text4.Text = Val(Text1.Text) - Val(Text2.Text)
Text5.Text = Val(Text1.Text) * Val(Text2.Text)
Text6.Text = Val(Text1.Text) / Val(Text2.Text)
Text7.Text = Val(Text1.Text) ^ Val(Text2.Text)
Text8.Text = Val(Text1.Text) ^ Val(1 / Text2.Text)
Label9.Caption = ""
ElseIf Text2.Text = 0 Then
Label9.Caption = "除数不能为零，请重新输入"
End If
Clear.Enabled = True
Calc.Enabled = False
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 And Len(Text2.Text) Then
Calc.Enabled = True
Else
Calc.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Len(Text1.Text) > 0 And Len(Text2.Text) Then
Calc.Enabled = True
Else
Calc.Enabled = False
End If
End Sub
