VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Index           =   0
      Left            =   240
      ScaleHeight     =   1665
      ScaleWidth      =   4065
      TabIndex        =   4
      Top             =   3000
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2535
      Index           =   1
      Left            =   240
      ScaleHeight     =   2475
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click me to move or resize"
      Height          =   1095
      Left            =   2520
      TabIndex        =   0
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "If you use control arrays as I did,then this code will be very useful and very short."
      Height          =   1215
      Left            =   5040
      TabIndex        =   3
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":0000
      Height          =   1335
      Left            =   5040
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X1 As Single
Dim Y1 As Single
Dim start As Boolean
Private Sub command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
start = True
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo hell:
X2 = X
Y2 = Y
If (start) Then
If Button = 2 Then
 With Command1
 .Width = X2
 .Height = Y2
End With
ElseIf Button = 1 Then

With Command1
.Move Command1.Left - X1 + X2, Command1.Top - Y1 + Y2
End With
End If

End If
hell:
Exit Sub
End Sub

Private Sub command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
start = False
End Sub


Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

X1 = X      'This holds the previouse X value for the mouse..only for moving
Y1 = Y      'This holds the previous y value for the mouse..only for moving
start = True
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X2, Y2 As Single
On Error GoTo hell:
X2 = X
Y2 = Y
If (start) Then
If Button = 2 Then
 With Picture1(Index)
 .Width = X2
 .Height = Y2
End With
ElseIf Button = 1 Then

With Picture1(Index)
.Move Picture1(Index).Left - X1 + X2, Picture1(Index).Top - Y1 + Y2
End With
End If

End If
hell:
Exit Sub
End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
start = False
End Sub
