VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Start Credits"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Credits"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin Project1.Credits Credits1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3625
      CreditForeColor =   -2147483644
      CreditBackColor =   -2147483647
      Align           =   2
      BorderColor     =   8421504
      BorderStyle     =   1
      CreditAppearance=   1
      Text            =   "Credits1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483626
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Credits1.sTimer = False
End Sub

Private Sub Command2_Click()
Credits1.sTimer = True
End Sub

Private Sub Form_Load()
Credits1.Text = "This is a Credit Example" & vbNewLine & _
"By Rob" & vbNewLine & "Hope you Like it! =)" & vbNewLine & _
"To Change the Other Stuff its in the Properties, Like Colors"

End Sub
