VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "VitruRope 1.0"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10755
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Clear"
      Height          =   255
      Left            =   1000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "Settings"
      ForeColor       =   &H00FFFF00&
      Height          =   4100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4700
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   360
         Top             =   1200
      End
      Begin VB.CheckBox AutoC 
         BackColor       =   &H00808000&
         Caption         =   "AutoColor"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3360
         TabIndex        =   26
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox GravityS 
         BackColor       =   &H00808000&
         Caption         =   "Gravity"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   2040
         Width           =   1095
      End
      Begin VB.HScrollBar Tlen 
         Height          =   255
         Left            =   960
         Max             =   200
         Min             =   1
         TabIndex        =   23
         Top             =   1680
         Value           =   20
         Width           =   3615
      End
      Begin VB.CheckBox CirclesD 
         BackColor       =   &H00808000&
         Caption         =   "Circles"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox LinesD 
         BackColor       =   &H00808000&
         Caption         =   "Lines"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   21
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   120
         Top             =   600
      End
      Begin VB.CommandButton HideSet 
         BackColor       =   &H00808000&
         Caption         =   "Hide"
         Height          =   255
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3840
         Width           =   4695
      End
      Begin VB.HScrollBar FPSset 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   18
         Top             =   1320
         Value           =   1
         Width           =   3615
      End
      Begin VB.CheckBox PointD 
         BackColor       =   &H00808000&
         Caption         =   "Points"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808000&
         Caption         =   "Color"
         ForeColor       =   &H00FFFF00&
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   4455
         Begin VB.CheckBox CselB 
            BackColor       =   &H00808000&
            Caption         =   "Blue"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   15
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox CselG 
            BackColor       =   &H00808000&
            Caption         =   "Green"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox CselR 
            BackColor       =   &H00808000&
            Caption         =   "Red"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label CselD 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   4215
         End
      End
      Begin VB.CheckBox FTrail 
         BackColor       =   &H00808000&
         Caption         =   "FadeTrail"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   2040
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox ClrRnd 
         BackColor       =   &H00808000&
         Caption         =   "ColorRnd"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.HScrollBar DrwWdth 
         Height          =   255
         Left            =   960
         Max             =   10
         Min             =   1
         TabIndex        =   7
         Top             =   960
         Value           =   1
         Width           =   3615
      End
      Begin VB.CheckBox AutoClr 
         BackColor       =   &H00808000&
         Caption         =   "Auto Clear"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.HScrollBar ScaterB 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   5
         Top             =   600
         Width           =   3615
      End
      Begin VB.HScrollBar ScaterA 
         Height          =   255
         Left            =   960
         Max             =   1000
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   120
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "Trail Lenth"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         Caption         =   "Frame Limit"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "DrawWidth"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Scater B"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Scater A"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton ShowSet 
      BackColor       =   &H00FFFF00&
      Caption         =   "Settings"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim dx(0 To 200), dy(0 To 200) As Long
Dim cx, cy, lx, ly As Long
Dim cliked As Boolean
Dim SelR, SelG, SelB As Integer
Dim HideOrShow As Boolean
Dim HideTemp As Integer
Dim ColCTemp As Integer
Dim TempSetShow As Boolean

'Code to run the user interface
Private Sub AutoC_Click()
If AutoC.Value = 1 Then
Timer3.Enabled = True
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Form1.Cls
End Sub

Private Sub Form_DblClick()
Form1.Cls
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cliked = True
If Button = vbLeftButton Then
If Timer1.Enabled = True Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
Else
For i = 0 To 200
dx(i) = cx
dy(i) = cy
Next i
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cx = X
cy = Y

If TempSetShow = True And Timer2.Enabled = False And Frame1.Visible = True Then
Timer2.Enabled = True
HideTemp = 10
HideOrShow = False
End If
End Sub

Private Sub DrwWdth_Change()
Form1.DrawWidth = DrwWdth.Value
End Sub

Private Sub FPSset_Change()
Timer1.Interval = FPSset
End Sub


Private Sub Frame1_Click()
Timer2.Enabled = True
HideTemp = 10
HideOrShow = False
End Sub

Private Sub HideSet_Click()
Timer2.Enabled = True
HideTemp = 10
HideOrShow = False
End Sub

Private Sub ShowSet_Click()
TempSetShow = False
End Sub

Private Sub ShowSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer2.Enabled = False And Frame1.Visible = False Then
Timer2.Enabled = True
HideTemp = 0
HideOrShow = True
Frame1.Visible = True
TempSetShow = True
End If
End Sub
'End of code to run the user interface



Private Sub Form_Load()
Frame1.Width = 0
Frame1.Height = 0
End Sub

Private Sub Timer1_Timer() 'The magic hapends here
UpdateColorSelection
'updates the first point to the mouse
dx(0) = cx
dy(0) = cy

For i = 1 To Tlen 'move the points towards each other
dx(i) = (dx(i) + dx(i - 1)) / 2
dy(i) = (dy(i) + dy(i - 1)) / 2
Next i
If AutoClr.Value = 1 Then Form1.Cls 'Automatic cleraing
For i = 1 To Tlen 'Rendering
If ClrRnd.Value = 1 Then 'If to name the colors shades random a bit
Form1.ForeColor = RGB((SelR / 2) + Rnd * 100, (SelG / 2) + Rnd * 100, (SelB / 2) + Rnd * 100)
Else
Form1.ForeColor = RGB(SelR, SelG, SelB)
End If

TrailFall = 255 \ Tlen.Value 'Later needed to figure out how fast the tail needs to fade

If GravityS.Value = 1 Then 'Simple gravity
For a = 0 To 200
dy(a) = dy(a) + 10
Next a
End If

If FTrail.Value = 1 Then Form1.ForeColor = RGB(BT0(SelR - i * TrailFall), BT0(SelG - i * TrailFall), BT0(SelB - i * TrailFall)) 'Fade the trail

If PointD.Value = 1 Then Form1.PSet (dx(i) + (Rnd * ScaterA), dy(i) + (Rnd * ScaterA)) 'Render points

If LinesD.Value = 1 Then Form1.Line (dx(i) + (Rnd * ScaterA), dy(i) + (Rnd * ScaterA))-(dx(i - 1) + (Rnd * ScaterB), dy(i - 1) + (Rnd * ScaterB)) 'Render lines

If CirclesD.Value = 1 Then Form1.Circle (dx(i), dy(i)), 100 - i * (100 \ Tlen.Value) ' Render circles


Next i
End Sub


Private Sub UpdateColorSelection() 'Validates the color selection check boxes
If CselR.Value = 1 Then
SelR = 255
Else
SelR = 0
End If
If CselG.Value = 1 Then
SelG = 255
Else
SelG = 0
End If
If CselB.Value = 1 Then
SelB = 255
Else
SelB = 0
End If
CselD.BackColor = RGB(SelR, SelG, SelB)
End Sub


Private Function BT0(Num)   'Bigger than 0
If Num >= 0 Then
BT0 = Num
Else
BT0 = 0
End If
End Function

Private Sub Timer2_Timer() 'Animates the menu
If HideOrShow = True Then
HideTemp = HideTemp + 1
Else
HideTemp = HideTemp - 1
End If
Frame1.Height = HideTemp * 410
Frame1.Width = HideTemp * 470
Command1.Left = 1000 + HideTemp * 380
If HideTemp = 10 Then Timer2.Enabled = False
If HideTemp = 0 Then
Timer2.Enabled = False
Frame1.Visible = False
End If
End Sub

Private Sub Timer3_Timer() 'Cyceling trough the colors
ColCTemp = ColCTemp + 1
Select Case ColCTemp
Case 1: CselR.Value = 1
Case 2: CselG.Value = 1
Case 3: CselR.Value = 0
Case 4: CselB.Value = 1
Case 5: CselG.Value = 0
Case 6: CselR.Value = 1
Case 7: CselB.Value = 0
End Select
If ColCTemp = 7 Then ColCTemp = 0
End Sub
