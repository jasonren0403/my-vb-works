VERSION 5.00
Begin VB.Form main 
   BackColor       =   &H80000016&
   Caption         =   "指针式钟表"
   ClientHeight    =   8025
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   24
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   11640
   Begin VB.CommandButton Exit 
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
      Height          =   615
      Left            =   9360
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9480
      Top             =   1080
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   615
      Left            =   8880
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   6
      Height          =   255
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   6735
      Left            =   1800
      Shape           =   3  'Circle
      Tag             =   "钟面内圈"
      Top             =   840
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   8
      Height          =   7485
      Left            =   1560
      Shape           =   3  'Circle
      Tag             =   "钟面外圈"
      Top             =   480
      Width           =   7245
   End
   Begin VB.Line Line3 
      BorderWidth     =   15
      Tag             =   "分针"
      X1              =   5160
      X2              =   5160
      Y1              =   4200
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   15
      Tag             =   "时针"
      X1              =   5160
      X2              =   5160
      Y1              =   4200
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Tag             =   "秒针"
      X1              =   5160
      X2              =   5160
      Y1              =   4200
      Y2              =   1080
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   13
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000016&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   3480
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000016&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   2410
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000016&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   2040
      TabIndex        =   10
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000016&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   2520
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000016&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   3600
      TabIndex        =   8
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000016&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   4960
      TabIndex        =   7
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000016&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   6360
      TabIndex        =   6
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000016&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   7320
      TabIndex        =   5
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000016&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   7920
      TabIndex        =   4
      Top             =   3780
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   7410
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000016&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6360
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4750
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pi, a
Dim s As Integer, f As Integer, m As Integer

Private Sub Exit_Click()
End
End Sub

Private Sub main_Load()
main.Show
End Sub

Private Sub Timer1_Timer()
pi = 4 * Atn(1)
a = Time
m = Val(Right(a, 2))
f = Val(Left(Right(a, 5), 2))
If Right(Left(a, 3), 1) = ":" Then
s = Val(Left(a, 2))
Else
s = Val(Left(a, 1))
End If
f = f + m / 60
s = s + f / 60
Line1.X2 = 5160 + 3000 * Sin(m * pi / 30)
Line1.Y2 = 4200 - 3000 * Cos(m * pi / 30)
Line3.X2 = 5160 + 2400 * Sin(f * pi / 30)
Line3.Y2 = 4200 - 2400 * Cos(f * pi / 30)
Line2.X2 = 5160 + 1800 * Sin(s * pi / 6)
Line2.Y2 = 4100 - 1800 * Cos(s * pi / 6)
Label13.Caption = Time
End Sub
