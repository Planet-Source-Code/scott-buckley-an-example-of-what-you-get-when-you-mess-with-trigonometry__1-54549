VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigonometric Graphing"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8775
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdchange 
      Caption         =   "Other"
      Height          =   1095
      Left            =   6480
      TabIndex        =   18
      Top             =   5160
      Width           =   2175
   End
   Begin MSComctlLib.Slider sldscale 
      Height          =   675
      Left            =   6360
      TabIndex        =   17
      Top             =   4200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1191
      _Version        =   393216
      LargeChange     =   50
      SmallChange     =   50
      Min             =   1
      Max             =   3
      SelStart        =   2
      TickStyle       =   2
      Value           =   2
   End
   Begin VB.Timer Timerstart 
      Interval        =   100
      Left            =   0
      Top             =   6120
   End
   Begin VB.CommandButton cmdnum2d 
      Caption         =   "<"
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdnum1d 
      Caption         =   "<"
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdnum2p 
      Caption         =   ">"
      Height          =   255
      Left            =   8400
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdnum1p 
      Caption         =   ">"
      Height          =   255
      Left            =   8400
      TabIndex        =   13
      Top             =   480
      Width           =   255
   End
   Begin VB.Frame frmopt2 
      Caption         =   "Operation 2"
      Height          =   1095
      Left            =   7680
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
      Begin VB.OptionButton optopt2 
         Caption         =   "Tan"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optopt2 
         Caption         =   "Cos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optopt2 
         Caption         =   "Sin"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame frmopt1 
      Caption         =   "Operation 1"
      Height          =   1095
      Left            =   6360
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
      Begin VB.OptionButton optopt1 
         Caption         =   "Tan"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optopt1 
         Caption         =   "Cos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optopt1 
         Caption         =   "Sin"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.Slider sldnum1 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   20
      SelStart        =   3
      TickStyle       =   1
      Value           =   3
   End
   Begin VB.TextBox txtnum2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      Text            =   "5"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtnum1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   1
      Text            =   "3"
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox graph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   6060
      Left            =   240
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   240
      Width           =   6060
   End
   Begin MSComctlLib.Slider sldnum2 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   20
      SelStart        =   5
      TickStyle       =   1
      Value           =   5
   End
   Begin VB.Label lbleq 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X = Function1(Number1 * Ø), Y = Function2(Number2 * Ø)"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   6360
      Width           =   8775
   End
   Begin VB.Label lblnum2 
      BackStyle       =   0  'Transparent
      Caption         =   "Number 2:"
      Height          =   255
      Left            =   6360
      TabIndex        =   20
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblnum1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number 1:"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menuchange 
         Caption         =   "Go To Other Equation"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuabout 
      Caption         =   "About"
      Begin VB.Menu menuabout2 
         Caption         =   "About Trigonometric Graphing"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim th As Double
Dim scal As Integer
Dim op1 As Integer
Dim op2 As Integer
Dim num1 As Double
Dim num2 As Double
Dim x As Double
Dim y As Double
Dim step As Double

Private Sub drawgraph()
    GraphCls
    For th = -3.2 To 3.2 Step step
        Select Case op1
        Case 0
            x = Sin(num1 * th)
        Case 1
            x = Cos(num1 * th)
        Case 2
            x = Tan(num1 * th)
        End Select
        Select Case op2
        Case 0
            y = Sin(num2 * th)
        Case 1
            y = Cos(num2 * th)
        Case 2
            y = Tan(num2 * th)
        End Select
            SetPixel graph.hdc, 200 + scal * x, 200 - scal * y, RGB(0, 0, 0)
    Next th
    graph.Refresh
End Sub

Private Sub cmdnum1d_Click()
    If num1 > 1 Then txtnum1 = txtnum1 - 1
End Sub

Private Sub cmdnum1p_Click()
    If num1 < 50 Then txtnum1 = txtnum1 + 1
End Sub

Private Sub cmdnum2d_Click()
    If num2 > 1 Then txtnum2 = txtnum2 - 1
End Sub

Private Sub cmdnum2p_Click()
    If num2 < 20 Then txtnum2 = txtnum2 + 1
End Sub

Private Sub cmdchange_Click()
    Form1.Hide
    Form2.Top = Form1.Top
    Form2.Left = Form1.Left
    Form2.Show
End Sub

Private Sub form_load()
    GraphCls
    step = 0.001
    scal = 100
End Sub

Private Sub GraphCls()
    graph.Cls
    graph.Line (200, 0)-(200, 400), RGB(255, 0, 0)
    graph.Line (0, 200)-(400, 200), RGB(255, 0, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub menuabout2_Click()
    frmAbout.Show
End Sub

Private Sub menuchange_Click()
    cmdchange_Click
End Sub

Private Sub menuexit_Click()
    End
End Sub

Private Sub optopt1_Click(Index As Integer)
    op1 = Index
    drawgraph
End Sub

Private Sub optopt2_Click(Index As Integer)
    op2 = Index
    drawgraph
End Sub

Private Sub sldnum1_Scroll()
    txtnum1.Text = sldnum1.Value
End Sub

Private Sub sldnum2_Scroll()
    txtnum2.Text = sldnum2.Value
End Sub

Private Sub pull()
    num1 = txtnum1.Text
    num2 = txtnum2.Text
    sldnum1.Value = num1
    sldnum2.Value = num2
    drawgraph
End Sub

Private Sub sldscale_Scroll()
    scal = sldscale.Value * 50
    drawgraph
End Sub

Private Sub Timerstart_Timer()
    pull
    drawgraph
    Timerstart.Enabled = False
End Sub

Private Sub txtnum1_Change()
    pull
End Sub

Private Sub txtnum2_Change()
    pull
End Sub
