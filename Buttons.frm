VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool buttons and scrollbars: by Johannes.B 2000"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      Height          =   1890
      Left            =   3960
      TabIndex        =   42
      Top             =   3000
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      Height          =   1230
      Left            =   3120
      TabIndex        =   40
      Top             =   3240
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   3000
      Width           =   255
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3120
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   38
      Top             =   3000
      Width           =   495
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "C:\"
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1440
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   36
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   35
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   960
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   34
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   4920
      Width           =   255
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   1530
      TabIndex        =   31
      Top             =   4680
      Width           =   1590
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "BUTTON3"
         Height          =   270
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   6360
      ScaleHeight     =   570
      ScaleWidth      =   1620
      TabIndex        =   29
      Top             =   3765
      Width           =   1620
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "BUTTON2"
         Height          =   270
         Left            =   45
         TabIndex        =   30
         Top             =   195
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6360
      ScaleHeight     =   555
      ScaleWidth      =   1545
      TabIndex        =   27
      Top             =   2880
      Width           =   1575
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "BUTTON1"
         Height          =   270
         Left            =   15
         TabIndex        =   28
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   1695
      Left            =   5160
      ScaleHeight     =   1635
      ScaleWidth      =   1515
      TabIndex        =   25
      Top             =   240
      Width           =   1575
      Begin VB.PictureBox Picture10 
         Height          =   4355
         Left            =   0
         Picture         =   "Buttons.frx":0000
         ScaleHeight     =   4290
         ScaleWidth      =   1515
         TabIndex        =   26
         Top             =   -1320
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Default"
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      Top             =   1920
      Width           =   1815
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   6720
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   23
      Top             =   240
      Width           =   285
   End
   Begin VB.CommandButton CLR 
      Caption         =   "CLR"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   5400
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   20
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   2040
      Width           =   255
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   2040
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   480
      ScaleHeight     =   1395
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "+"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New backcolor (some values = 0)"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Default"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   4440
      Top             =   1800
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   960
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      Height          =   615
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   1800
      ScaleHeight     =   555
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "+"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "B"
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "G"
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "R"
      Height          =   255
      Left            =   1800
      TabIndex        =   44
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "BUTTON4"
      Height          =   255
      Left            =   6720
      TabIndex        =   43
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   1920
      TabIndex        =   37
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   315
      Left            =   690
      Top             =   4890
      Width           =   1020
   End
   Begin VB.Line Line5 
      X1              =   6360
      X2              =   7920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line4 
      X1              =   7920
      X2              =   8040
      Y1              =   4560
      Y2              =   4680
   End
   Begin VB.Line Line3 
      X1              =   6360
      X2              =   6480
      Y1              =   4560
      Y2              =   4680
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6360
      Y1              =   4560
      Y2              =   5040
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6480
      Y1              =   5040
      Y2              =   5160
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   3720
      TabIndex        =   21
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Unchecked"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim JB As Integer
Dim JBB As Integer

Dim A, B As Integer
Dim N As Integer
Dim U As Integer

Dim JV As Integer

Dim RR, GG, BB As Integer
Dim JP As Integer

Dim JH As Integer

Dim o As Integer
Dim JK As Integer

Dim AB, AC As Integer
Dim COL As Integer
Private Sub CLR_Click()

JBB = 0
Picture2.Cls

Picture6.Cls
Picture7.Cls
Picture8.Cls
End Sub

Private Sub Command1_Click()
If JB > 0 Then
Picture1.ForeColor = Picture1.BackColor
Picture1.Circle (JB * 150, Picture1.Height / 3), 50
JB = JB - 1
End If
End Sub

Private Sub Command10_Click()
If Command10.Caption = "+" Then
List1.Visible = True
Command10.Caption = "-"
Else
List1.Visible = False
Command10.Caption = "+"
End If
End Sub

Private Sub Command2_Click()
If JB < 8 Then
JB = JB + 1
Picture1.ForeColor = RGB(JB * 30, 0, 0)
Picture1.Circle (JB * 150, Picture1.Height / 3), 50
End If
End Sub

Private Sub Command3_Click()
If JBB > 0 Then
For Counter = 1 To 2 Step 1
Picture2.ForeColor = Picture2.BackColor
Picture2.Line (JBB * 60, Picture1.Height * 2)-(JBB * 60, 0)
JBB = JBB - 1
Next
End If
End Sub

Private Sub Command4_Click()
If JBB < 22 Then
For Counter = 1 To 2 Step 1
JBB = JBB + 1
If Label6.Caption = 1 Then Picture2.ForeColor = RGB(JBB * 12, 0, 0)
If Label6.Caption = 2 Then Picture2.ForeColor = RGB(0, JBB * 12, 0)
If Label6.Caption = 3 Then Picture2.ForeColor = RGB(0, 0, JBB * 12)
Picture2.Line (JBB * 60, Picture1.Height * 2)-(JBB * 60, 0)
Next
End If
End Sub

Private Sub Command5_Click()
Picture3.ForeColor = Picture3.BackColor
Picture3.Circle (A, B), 5


A = Picture3.ScaleWidth / 2
B = Picture3.ScaleHeight / 2

Picture3.ForeColor = vbBlack
Picture3.Circle (A, B), 5

End Sub

Private Sub Command6_Click()

RR = Rnd * 255
GG = Rnd * 255
BB = Rnd * 255

Picture1.BackColor = RGB(RR, GG, BB)
Picture2.BackColor = RGB(RR, GG, BB)
Picture3.BackColor = RGB(RR, GG, BB)
Picture4.BackColor = RGB(RR, GG, BB)
Picture5.BackColor = RGB(RR, GG, BB)
Picture6.BackColor = RGB(RR, GG, BB)
Picture7.BackColor = RGB(RR, GG, BB)
Picture8.BackColor = RGB(RR, GG, BB)
Picture9.BackColor = RGB(RR, GG, BB)

Picture12.BackColor = RGB(RR, GG, BB)
Picture13.BackColor = RGB(RR, GG, BB)
Picture14.BackColor = RGB(RR, GG, BB)
Picture15.BackColor = RGB(RR, GG, BB)
Picture16.BackColor = RGB(RR, GG, BB)
Picture17.BackColor = RGB(RR, GG, BB)
Picture18.BackColor = RGB(RR, GG, BB)
Picture19.BackColor = RGB(RR, GG, BB)
List1.BackColor = RGB(RR, GG, BB)
Dir1.BackColor = RGB(RR, GG, BB)

Command9.Value = True

JB = 0
JBB = 0
JP = 0

Command5.Value = True

'Reset checkbox
JV = 0
Label4.Caption = "Unchecked"

'Option 1's value = true
Label6.Caption = 1
Picture6.Circle (Picture6.Width / 3, Picture6.Height / 3), 60
'''

End Sub

Private Sub Command7_Click()
If JP > 0 Then

Picture5.ForeColor = Picture5.BackColor
Picture5.Circle (Picture5.Width / 3, Picture5.Height - JP * 165), 50

JP = JP - 1

End If
End Sub

Private Sub Command8_Click()
If JP < 8 Then
JP = JP + 1
Picture5.ForeColor = RGB(JP * 30, 0, 0)
Picture5.Circle (Picture5.Width / 3, Picture5.Height - JP * 165), 50
End If
End Sub

Private Sub Command9_Click()
Picture9.ForeColor = Picture9.BackColor
Picture9.Circle (Picture9.ScaleWidth / 2, JH), 4

JH = Picture9.ScaleHeight / 2

Picture10.Move (0), (JH * 23 - Picture10.Height + 1700)

Picture9.ForeColor = vbBlack
Picture9.Circle (Picture9.ScaleWidth / 2, JH), 4

End Sub

Private Sub Form_Activate()
Picture3.Circle (A, B), 5
'Option 1's value = true
Label6.Caption = 1
Picture6.Circle (Picture6.Width / 3, Picture6.Height / 3), 60
'''
Picture9.Circle (Picture9.ScaleWidth / 2, JH), 4
List1.Text = "C:\"
Dir1.Path = "C:\"
End Sub

Private Sub Form_Load()
A = Picture3.ScaleWidth / 2
B = Picture3.ScaleHeight / 2

Picture3.ForeColor = vbBlack

JH = Picture9.ScaleHeight / 2


List1.AddItem "A:\"
List1.AddItem "B:\"
List1.AddItem "C:\"
List1.AddItem "D:\"
List1.AddItem "E:\"
List1.AddItem "F:\"
List1.AddItem "G:\"
List1.AddItem "H:\"
List1.AddItem "I:\"
List1.AddItem "J:\"
List1.AddItem "K:\"
List1.AddItem "L:\"
List1.AddItem "M:\"
List1.AddItem "N:\"
List1.AddItem "O:\"
List1.AddItem "P:\"
List1.AddItem "Q:\"
List1.AddItem "R:\"
List1.AddItem "S:\"
List1.AddItem "T:\"
List1.AddItem "U:\"
List1.AddItem "V:\"
List1.AddItem "W:\"
List1.AddItem "X:\"
List1.AddItem "Y:\"
List1.AddItem "Z:\"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BorderStyle = 1
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.BorderStyle = 0
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Appearance = 1
Picture12.BackColor = Picture1.BackColor
End Sub


Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Appearance = 0
Picture12.BackColor = Picture1.BackColor
End Sub


Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.BorderStyle = 1
End Sub


Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.BorderStyle = 0
End Sub


Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AB = Picture14.Left
AC = Picture14.Top

Line1.Visible = False
Line4.Visible = False

Picture14.Left = AB - 70
Picture14.Top = AC - 70
End Sub


Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture14.Left = AB
Picture14.Top = AC

Line1.Visible = True
Line4.Visible = True

End Sub


Private Sub List1_Click()
On Error GoTo io
Label11.Caption = List1.Text
List1.Visible = False
Command10.Caption = "+"
Dir1.Path = List1.Text
Exit Sub
io:
MsgBox "Drive not ready/found"
Exit Sub
End Sub

Private Sub Picture12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Appearance = 1
Picture12.BackColor = Picture1.BackColor
End Sub


Private Sub Picture12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Appearance = 0
Picture12.BackColor = Picture1.BackColor
End Sub


Private Sub Picture13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.BorderStyle = 1
End Sub


Private Sub Picture13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.BorderStyle = 0
End Sub


Private Sub Picture14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

AB = Picture14.Left
AC = Picture14.Top

Line1.Visible = False
Line4.Visible = False

Picture14.Left = AB - 70
Picture14.Top = AC - 70
End Sub


Private Sub Picture14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


Picture14.Left = AB
Picture14.Top = AC


Line1.Visible = True
Line4.Visible = True
End Sub


Private Sub Picture15_Click()
Picture15.BorderStyle = 1
Picture16.BorderStyle = 0
Picture17.BorderStyle = 0
Picture18.BorderStyle = 0
Label10.Caption = 1
End Sub

Private Sub Picture16_Click()
Picture15.BorderStyle = 0
Picture16.BorderStyle = 1
Picture17.BorderStyle = 0
Picture18.BorderStyle = 0
Label10.Caption = 2
End Sub


Private Sub Picture17_Click()
Picture15.BorderStyle = 0
Picture16.BorderStyle = 0
Picture17.BorderStyle = 1
Picture18.BorderStyle = 0
Label10.Caption = 3
End Sub


Private Sub Picture18_Click()
Picture15.BorderStyle = 0
Picture16.BorderStyle = 0
Picture17.BorderStyle = 0
Picture18.BorderStyle = 1
Label10.Caption = 4
End Sub


Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
N = 1
End Sub


Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If N = 1 Then

Picture3.ForeColor = Picture3.BackColor
Picture3.Circle (A, B), 5


A = X
B = Y

If A < 0 Then A = 0
If B < 0 Then B = 0
If A > Picture3.ScaleWidth Then A = Picture3.ScaleWidth
If B > Picture3.ScaleHeight Then B = Picture3.ScaleHeight

Picture3.ForeColor = vbBlack
Picture3.Circle (A, B), 5
End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
N = 0
End Sub


Private Sub Picture4_Click()
If JV = 0 Then
JV = 1
Label4.Caption = "Checked"
Picture4.ForeColor = vbBlack
Picture4.Circle (Picture4.Width / 3, Picture4.Height / 3), 60

Else
JV = 0
Label4.Caption = "Unchecked"

Picture4.ForeColor = Picture4.BackColor
Picture4.Circle (Picture4.Width / 3, Picture4.Height / 3), 60


End If
End Sub

Private Sub Picture6_Click()

CLR.Value = True

Label6.Caption = 1
Picture6.Circle (Picture6.Width / 3, Picture6.Height / 3), 60


End Sub


Private Sub Picture7_Click()
CLR.Value = True

Label6.Caption = 2
Picture7.Circle (Picture7.Width / 3, Picture7.Height / 3), 60
End Sub


Private Sub Picture8_Click()
CLR.Value = True

Label6.Caption = 3
Picture8.Circle (Picture8.Width / 3, Picture8.Height / 3), 60
End Sub


Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
U = 1
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If U = 1 Then
Picture9.ForeColor = Picture9.BackColor
Picture9.Circle (Picture9.ScaleWidth / 2, JH), 4

JH = Y

If JH > Picture9.ScaleHeight Then JH = Picture9.ScaleHeight
If JH < 0 Then JH = 0

Picture10.Move (0), (JH * 23 - Picture10.Height + 1700)

Picture9.ForeColor = vbBlack
Picture9.Circle (Picture9.ScaleWidth / 2, JH), 4
End If
End Sub


Private Sub Picture9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
U = 0
End Sub

Private Sub Timer1_Timer()
Label1.Caption = JB
Label2.Caption = JBB
Label3.Caption = A & ", " & B
Label5.Caption = JP
End Sub


