VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Electro Zone 2.0"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   Icon            =   "EZ2-1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "EZ2-1.frx":185A
   ScaleHeight     =   7245
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "end"
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1080
      Top             =   240
   End
   Begin VB.PictureBox picSplash 
      BackColor       =   &H00FF00FF&
      Height          =   5895
      Left            =   4560
      ScaleHeight     =   5835
      ScaleWidth      =   3315
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   360
      TabIndex        =   0
      Top             =   2840
      Width           =   8625
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   840
         Top             =   240
      End
      Begin VB.Image Image3 
         Height          =   1035
         Left            =   -8640
         Picture         =   "EZ2-1.frx":19AC
         Top             =   0
         Width           =   17250
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A BIG THANKS IN ADVANCE !!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   6000
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "if u have gmail/orkut then meet me :  najeebputhiyallam@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   8535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "if anyone can help me please please mail me :  najeeb_puthiyallam@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5400
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"EZ2-1.frx":3BC5C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   8895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8280
      TabIndex        =   2
      Top             =   6900
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   1035
      Left            =   120
      Picture         =   "EZ2-1.frx":3BCE5
      Top             =   600
      Visible         =   0   'False
      Width           =   8625
   End
   Begin VB.Image Image1 
      Height          =   7245
      Left            =   0
      Picture         =   "EZ2-1.frx":3C5DC
      Top             =   0
      Width           =   9450
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(114, 0, 49)
'Me.Width = 3375
'Me.Height = 5895
picSplash.Picture = Image2.Picture
picSplash.BorderStyle = 0
Call createSkinnedForm(Frame1, picSplash)
End Sub

Private Sub Form_Resize()
'Me.Width = 8595
'Me.Height = 5865
End Sub

Private Sub Timer1_Timer()
If Image3.Left >= 0 Then
Image3.Left = -8640
Else
Image3.Left = Image3.Left + 50
End If
End Sub

Private Sub Timer2_Timer()
i = i + 1
If i > 99 Then
'End
Else
Label1.Caption = "Loadin " & i & "%"
End If
End Sub
