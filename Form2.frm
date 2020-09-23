VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigonometric Graphing"
   ClientHeight    =   6615
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8775
   Icon            =   "Form2.frx":0000
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
      TabIndex        =   10
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtamp 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   9
      Text            =   "-10"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtfrq 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   8
      Text            =   "3"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdampp 
      Caption         =   ">"
      Height          =   255
      Left            =   8400
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdfrqp 
      Caption         =   ">"
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   1800
      Width           =   255
   End
   Begin VB.CommandButton cmdampd 
      Caption         =   "<"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdfrqd 
      Caption         =   "<"
      Height          =   255
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin MSComctlLib.Slider sldradius 
      Height          =   675
      Left            =   6360
      TabIndex        =   3
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
   Begin MSComctlLib.Slider sldamp 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   -50
      Max             =   50
      SelStart        =   -10
      TickStyle       =   1
      TickFrequency   =   5
      Value           =   -10
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
   Begin MSComctlLib.Slider sldfrq 
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Max             =   20
      SelStart        =   3
      TickStyle       =   1
      Value           =   3
   End
   Begin VB.Label lbleq 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X = (Radius - (Amplitude * sin(Ø * Frequency))) * sin(Ø), X = (Radius - (Amplitude * sin(Ø * Frequency))) * cos(Ø)"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6360
      Width           =   8775
   End
   Begin VB.Label lblamp 
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude:"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblfrq 
      BackStyle       =   0  'Transparent
      Caption         =   "Frequency:"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   1560
      Width           =   855
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Index           =   1
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim th As Double
Dim radius As Integer
Dim scal As Double
Dim op1 As Integer
Dim op2 As Integer
Dim amp As Double
Dim frq As Double
Dim x As Double
Dim y As Double
Dim step As Double

Private Sub drawgraph()
    GraphCls
    For th = -3.2 To 3.2 Step step
        scal = radius - amp * Sin(th * frq)
        x = scal * Sin(th)
        y = scal * Cos(th)
        SetPixel graph.hdc, 200 + x, 200 - y, RGB(0, 0, 0)
    Next th
    graph.Refresh
End Sub

Private Sub cmdampd_Click()
    If amp > -50 Then txtamp = txtamp - 1
End Sub

Private Sub cmdampp_Click()
    If amp < 50 Then txtamp = txtamp + 1
End Sub

Private Sub cmdchange_Click()
    Form2.Hide
    Form1.Top = Form2.Top
    Form1.Left = Form2.Left
    Form1.Show
End Sub

Private Sub cmdfrqd_Click()
    If frq > 0 Then txtfrq = txtfrq - 1
End Sub

Private Sub cmdfrqp_Click()
    If frq < 20 Then txtfrq = txtfrq + 1
End Sub

Private Sub form_load()
    GraphCls
    step = 0.001
    radius = 100
End Sub

Public Sub GraphCls()
    graph.Cls
    graph.Line (200, 0)-(200, 400), RGB(255, 0, 0)
    graph.Line (0, 200)-(400, 200), RGB(255, 0, 0)
End Sub

Private Sub optopt1_Click(Index As Integer)
    op1 = Index
    drawgraph
End Sub

Private Sub optopt2_Click(Index As Integer)
    op2 = Index
    drawgraph
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

Private Sub sldamp_Scroll()
    txtamp.Text = sldamp.Value
End Sub

Private Sub sldfrq_Scroll()
    txtfrq.Text = sldfrq.Value
End Sub

Public Sub pull()
    amp = txtamp.Text
    frq = txtfrq.Text
    sldamp.Value = amp
    sldfrq.Value = frq
    drawgraph
End Sub

Private Sub sldradius_scroll()
    radius = sldradius.Value * 50
    drawgraph
End Sub

Private Sub Timerstart_Timer()
    pull
    drawgraph
    Timerstart.Enabled = False
End Sub

Private Sub txtamp_Change()
    pull
End Sub

Private Sub txtfrq_Change()
    pull
End Sub
