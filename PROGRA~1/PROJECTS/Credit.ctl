VERSION 5.00
Begin VB.UserControl Credits 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H80000006&
   KeyPreview      =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   3135
   ToolboxBitmap   =   "Credit.ctx":0000
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FontTransparent =   0   'False
      Height          =   2175
      Left            =   60
      ScaleHeight     =   2175
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000004&
         Height          =   525
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "Credit.ctx":0312
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   120
         Top             =   1440
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000004&
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Enum iBorder
  [none] = 0
  [Fixed Single] = 1
End Enum
Enum iApperance
  [Flat] = 0
  [3D] = 1
End Enum
Enum iBack
  [Transparent] = 0
  [Opaque] = 1
End Enum
Private m_Text As String
Dim p As Integer
Private Sub delay()
  Dim PauseTime, Start, Finish
  PauseTime = 2.5
  Start = Timer
  Do While Timer < Start + PauseTime
      DoEvents
  Loop
  Finish = Timer
End Sub
Public Property Get sTimer() As Boolean
  sTimer = Timer1.Enabled
End Property
Public Property Let sTimer(ByVal New_timer As Boolean)
  Timer1.Enabled() = New_timer
  Text1.Top = Picture1.Height
  PropertyChanged "sTimer"
End Property
Public Property Get CreditBorderStyle() As iBorder
  CreditBorderStyle = Picture1.BorderStyle
End Property
Public Property Let CreditBorderStyle(ByVal New_CreditBorderStyle As iBorder)
  Picture1.BorderStyle() = New_CreditBorderStyle
  PropertyChanged "CreditBorderStyle"
End Property
Public Property Get Font() As Font
  Set Font = Text1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
  Set Text1.Font = New_Font
  PropertyChanged "Font"
End Property
Public Property Get Appearance() As iApperance
  Appearance = Label1.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As iApperance)
  Label1.Appearance() = New_Appearance
  PropertyChanged "Appearance"
End Property
Public Property Get CreditAppearance() As iApperance
  CreditAppearance = Picture1.Appearance
End Property
Public Property Let CreditAppearance(ByVal New_CreditAppearance As iApperance)
  Picture1.Appearance() = New_CreditAppearance
  PropertyChanged "CreditAppearance"
End Property
Public Property Get BorderStyle() As iBorder
  BorderStyle = Label1.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As iBorder)
  Label1.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = UserControl.BackColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  UserControl.BackColor() = New_BorderColor
  PropertyChanged "BorderColor"
End Property
Public Property Get Align() As AlignmentConstants
  Align = Text1.Alignment
End Property
Public Property Let Align(ByVal New_Align As AlignmentConstants)
  Text1.Alignment() = New_Align
  PropertyChanged "Align"
End Property
Public Property Get CreditForeColor() As OLE_COLOR
  CreditForeColor = Text1.ForeColor
End Property
Public Property Let CreditForeColor(ByVal New_CreditForeColor As OLE_COLOR)
  Text1.ForeColor() = New_CreditForeColor
  PropertyChanged "CreditForeColor"
End Property
Public Property Get CreditBackColor() As OLE_COLOR
  CreditBackColor = Text1.BackColor
End Property
Public Property Let CreditBackColor(ByVal New_CreditBackColor As OLE_COLOR)
  Text1.BackColor() = New_CreditBackColor
  PropertyChanged "CreditBackColor"
End Property
Public Property Get Text() As String
  Text = m_Text
End Property
Public Property Let Text(ByVal New_text As String)
  Text1 = New_text
  m_Text = New_text
End Property
Public Property Get BackColor() As OLE_COLOR
  BackColor = Picture1.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  Picture1.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

Private Sub Timer1_Timer()
DoEvents
If Text1.Top < p Then
    p = -9999
    delay
End If
If Text1.Top = 0 - Text1.Height Then
    Text1.Top = Picture1.Height + 100
End If

Text1.Top = Text1.Top - 10

End Sub
Private Sub UserControl_Initialize()
DoEvents
p = 10
End Sub
Private Sub UserControl_InitProperties()
  m_Text = Extender.Name
  Text1.Text = m_Text
  Timer1.Enabled = False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Text1.ForeColor = PropBag.ReadProperty("CreditForeColor", &H8000000F)
  Text1.BackColor = PropBag.ReadProperty("CreditBackColor", &H80000006)
  Text1.Alignment = PropBag.ReadProperty("Align", 0)
  UserControl.BackColor = PropBag.ReadProperty("BorderColor", &HC00000)
  Picture1.BorderStyle = PropBag.ReadProperty("CreditBorderStyle", 0)
  Label1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  Label1.Appearance = PropBag.ReadProperty("Appearance", 0)
  Picture1.Appearance = PropBag.ReadProperty("CreditAppearance", 0)
  Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
  Text = PropBag.ReadProperty("Text", Extender.Name)
  Picture1.BackColor = PropBag.ReadProperty("BackColor", &H80000004)
  Timer1.Enabled = PropBag.ReadProperty("sTimer", False)
End Sub
Private Sub UserControl_Resize()
  Picture1.Height = UserControl.Height - 120
  Picture1.Width = UserControl.Width - 125
  Text1.Width = Picture1.Width
  Text1.Top = Picture1.Height
  Text1.Height = Picture1.Height
  Label1.Width = UserControl.Width
  Label1.Height = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("CreditForeColor", Text1.ForeColor, &H8000000F)
  Call PropBag.WriteProperty("CreditBackColor", Text1.BackColor, &H80000006)
  Call PropBag.WriteProperty("Align", Text1.Alignment, 0)
  Call PropBag.WriteProperty("BorderColor", UserControl.BackColor, &HC00000)
  Call PropBag.WriteProperty("CreditBorderStyle", Picture1.BorderStyle, 0)
  Call PropBag.WriteProperty("BorderStyle", Label1.BorderStyle, 0)
  Call PropBag.WriteProperty("Appearance", Label1.Appearance, 0)
  Call PropBag.WriteProperty("CreditAppearance", Picture1.Appearance, 0)
  Call PropBag.WriteProperty("Text", m_Text)
  Call PropBag.WriteProperty("Font", Text1.Font)
  Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H80000004)
  Call PropBag.WriteProperty("sTimer", Timer1.Enabled, False)
End Sub


