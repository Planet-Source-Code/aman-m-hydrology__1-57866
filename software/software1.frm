VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Designing aid for Hydro-Projects :"
   ClientHeight    =   7860
   ClientLeft      =   4350
   ClientTop       =   525
   ClientWidth     =   6975
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H8000000A&
   Icon            =   "software1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   6975
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Table"
      Height          =   375
      Left            =   5640
      TabIndex        =   46
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   2280
      TabIndex        =   40
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   2280
      TabIndex        =   39
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   2280
      TabIndex        =   37
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2280
      TabIndex        =   36
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   2280
      TabIndex        =   35
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Average value of the above values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   33
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "By Moody's formula"
      Height          =   495
      Left            =   5400
      TabIndex        =   31
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "By Nagi's formula"
      Height          =   495
      Left            =   5400
      TabIndex        =   28
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   5280
      Width           =   855
   End
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   0
      TabIndex        =   25
      Top             =   4560
      Width           =   6975
      Begin VB.CommandButton Command4 
         Caption         =   "By RW Abetic formula"
         Height          =   495
         Left            =   5400
         TabIndex        =   27
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Rated Head of Turbine :"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Enter the value of P.F :"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Operating Head :"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "No. of Turbine Units :"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Capacity to be installed :"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Calculation of Specific Speed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3480
      TabIndex        =   23
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3480
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3480
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Forebay"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Power"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3480
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Design of a Turbine :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   4095
      End
      Begin VB.Label Label5 
         Caption         =   "Watts"
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Value of Load Factor :"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Efficiency of the Turbine :"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Value of Head Entered (m) :"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Value of Discharge Entered (m3/s) :"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   4920
         Picture         =   "software1.frx":0442
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   6975
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   6360
         Picture         =   "software1.frx":952B4
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Design of Transition of Forebay (Mitra's Method) :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label10 
         Caption         =   "Value of Bf Entered :"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Value of Lf Entered :"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label Label8 
         Caption         =   "Value of Bc Entered :"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Value of x Entered :"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim Lf, l, Q, power, H As Double
Dim CG, efficiency, e As Double
Q = Text5.Text
H = Text6.Text
e = Text7.Text
Lf = Text8.Text
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
power = (9.81 * e * Q * H)
CG = (power / Lf)
Text1.Text = CG
End Sub

Private Sub Command2_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim x, Bx1, Bx2, Bx, g, Bxx, e, m, N, Bf, Bc, Lf, xq, l, b, c As Double
x = Text4.Text
Lf = Text9.Text
Bf = Text10.Text
Bc = Text3.Text
m = (Bc ^ 1.5)
N = (Bf ^ 1.5)
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
Bxx = Lf * m
e = m - N
Bx1 = (x * e) / Bxx
Bx2 = ((1 - Bx1) ^ 0.666666666666667)
Bx = Bf / Bx2
Text2.Text = Bx
On Error Resume Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
Load Form2
On Error Resume Next
Form2.Visible = True
On Error Resume Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim j, P, PF, Hmax, hr, z, N, a, b As Double
P = Text13.Text
N = Text14.Text
PF = Text18.Text
z = P / N
Hmax = Text19.Text
hr = Hmax * ((PF / 1.15) ^ 0.6666667)
Text11.Text = hr
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
Dim Ns, v, w As Double
hr = Text11.Text
Ns = 1700 / (hr ^ 0.5)
Text12.Text = Ns
On Error Resume Next
End Sub

Private Sub Command5_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim j, P, PF, Hmax, hr, z, N, a, b As Double
P = Text13.Text
N = Text14.Text
PF = Text18.Text
z = P / N
Hmax = Text19.Text
hr = Hmax * ((PF / 1.15) ^ 0.6666667)
Text11.Text = hr
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
Dim Ns, v, w As Double
hr = Text11.Text
Ns = 1640 / (hr ^ 0.5)
Text15.Text = Ns
On Error Resume Next
End Sub

Private Sub Command6_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim j, P, PF, Hmax, hr, z, N, a, b As Double
P = Text13.Text
N = Text14.Text
PF = Text18.Text
z = P / N
Hmax = Text19.Text
hr = Hmax * ((PF / 1.15) ^ 0.6666667)
Text11.Text = hr
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
Dim Ns, v, w As Double
hr = Text11.Text
Ns = (6780 / (hr + 9.5)) + 83.6
Text16.Text = Ns
On Error Resume Next
End Sub

Private Sub Command7_Click()
On Error Resume Next
ProgressBar1.Visible = True
Dim j, P, PF, Hmax, hr, z, N, a, b As Double
P = Text13.Text
N = Text14.Text
PF = Text18.Text
z = P / N
Hmax = Text19.Text
hr = Hmax * ((PF / 1.15) ^ 0.6666667)
Text11.Text = hr
Dim Counter As Integer
ProgressBar1.Value = 0
For Counter = 1 To 999
ProgressBar1.Value = ProgressBar1.Value + 0.1
Next Counter
ProgressBar1.Visible = False
Dim Ns1, Ns2, Ns3, Ns, v, w As Double
Ns1 = 1700 / (hr ^ 0.5)
Ns2 = 1640 / (hr ^ 0.5)
Ns3 = (6780 / (hr + 9.5)) + 83.6
Ns = (Ns1 + Ns2 + Ns3) / 3
Text17.Text = Ns
On Error Resume Next
End Sub

Private Sub Command8_Click()
On Error Resume Next
Load frmdam
frmdam.Visible = True
On Error Resume Next
End Sub

Private Sub Command9_Click()
On Error Resume Next
Load Form3
Form3.Visible = True
End Sub

