VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "About"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Image Image1 
         Height          =   2550
         Left            =   0
         Picture         =   "software2.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1830
      End
   End
   Begin VB.Label Label3 
      Caption         =   "A++ Inc."
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Thanking You"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"software2.frx":6AE1
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
