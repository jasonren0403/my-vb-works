VERSION 5.00
Begin VB.Form demo 
   Caption         =   "演示循环使用"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9555
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ExitButton 
      Caption         =   "退出"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "简单实现太极图"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "快速计算从1加到100"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "从2到10，偶数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "从2到11"
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "从1到10"
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
Show
For x = 1 To 10
Print x
Next x
End Sub

Private Sub Command2_Click()
Cls
Show
For x = 2 To 11
Print x
Next x
End Sub

Private Sub Command3_Click()
Cls
Show
For x = 2 To 10 Step 2
Print x
Next x
End Sub

Private Sub Command4_Click()
Cls
Show
t = 0
For n = 1 To 100
t = t + n
Next n
Print t
End Sub

Private Sub Command5_Click()
Cls
Show
Circle (3500, 3000), 2000
Circle (3500, 2000), 1000, , 3.1415926 / 2, 3 * 3.1415926 / 2
Circle (3500, 4000), 1000, , 3 * 3.1415926 / 2, 3.1415926 / 2
For x = 1 To 200 Step 0.1
Circle (3500, 2000), 200 - x
Circle (3500, 4000), 200 - x
Next x
End Sub

Private Sub ExitButton_Click()
End
End Sub
