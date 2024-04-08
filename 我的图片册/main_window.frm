VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "图片查看器"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   4920
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   810
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   720
      Left            =   360
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "图像预览"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "打开的路径"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5880
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label4 
      Caption         =   "文件名"
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
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "目录名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "驱动器名"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form2.Image1.Width = 12000
Form2.Image1.Height = 10000
Form2.Image1.Picture = Form1.Image1.Picture
Form2.Width = 12000
Form2.Height = 10000
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Dir1_Change()
ChDir Dir1.Path
File1.Path = Dir1.Path
File1.Pattern = "*.bmp;*.jpg;*.gif"
Label3.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
ChDrive Drive1.Drive
Dir1.Path = Drive1.Drive
File1.Pattern = "*.bmp;*.jpg;*.gif"
End Sub

Private Sub File1_Click()
Image1.Picture = LoadPicture(File1.FileName)
End Sub

Private Sub Form_Load()
File1.Pattern = "*.bmp;*.jpg;*.gif"
Label3.Caption = Dir1.Path
End Sub

