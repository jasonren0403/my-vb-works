VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "调色板"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   10050
   StartUpPosition =   3  '窗口缺省
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      LargeChange     =   10
      Left            =   2280
      Max             =   255
      TabIndex        =   6
      Top             =   3840
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      LargeChange     =   10
      Left            =   2280
      Max             =   255
      TabIndex        =   5
      Top             =   3120
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   10
      Left            =   2280
      Max             =   255
      TabIndex        =   4
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton ExitButton 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton ResetButton 
      Caption         =   "重置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
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
      Left            =   7080
      TabIndex        =   0
      Text            =   "我会变色哦！"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "颜色显示区域"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "文字效果"
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   5760
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
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
      Left            =   5760
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ResetButton_Click()
HScroll1.Value = 0
HScroll2.Value = 0
HScroll3.Value = 0
Label1.BackColor = RGB(0, 0, 0)
Text2.ForeColor = Label1.BackColor
End Sub

Private Sub ExitButton_Click()
End
End Sub

Private Sub HScroll1_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.ForeColor = Label1.BackColor
Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.ForeColor = Label1.BackColor
Label3.Caption = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Text2.ForeColor = Label1.BackColor
Label4.Caption = HScroll3.Value
End Sub

