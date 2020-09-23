VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5505
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   195
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   930
      ScaleWidth      =   630
      TabIndex        =   27
      Top             =   4335
      Width           =   630
   End
   Begin VB.Frame Frame2 
      Caption         =   "Difficulty"
      Height          =   2355
      Left            =   990
      TabIndex        =   20
      Top             =   1680
      Width           =   1455
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   390
         Left            =   195
         TabIndex        =   19
         Top             =   1860
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Beginner"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Intermediate"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Advanced"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   810
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Expert"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1140
         Width           =   930
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Crazy!"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1470
         Width           =   870
      End
   End
   Begin VB.Timer time 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   3075
   End
   Begin VB.Timer start 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4515
      Top             =   2205
   End
   Begin VB.PictureBox Picture2 
      Height          =   990
      Left            =   4425
      ScaleHeight     =   930
      ScaleWidth      =   1125
      TabIndex        =   14
      Top             =   345
      Width           =   1185
      Begin VB.Image light 
         Height          =   870
         Index           =   2
         Left            =   765
         Picture         =   "Form1.frx":1F44
         Top             =   30
         Width           =   330
      End
      Begin VB.Image light 
         Height          =   870
         Index           =   1
         Left            =   390
         Picture         =   "Form1.frx":24BC
         Top             =   30
         Width           =   330
      End
      Begin VB.Image light 
         Height          =   870
         Index           =   0
         Left            =   15
         Picture         =   "Form1.frx":2A2D
         Top             =   30
         Width           =   330
      End
   End
   Begin VB.Timer randomAccel 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3150
      Top             =   3180
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2970
      Top             =   1470
   End
   Begin VB.Timer coneCollision6 
      Interval        =   1
      Left            =   1980
      Top             =   2520
   End
   Begin VB.Timer coneCollision5 
      Interval        =   1
      Left            =   1605
      Top             =   2505
   End
   Begin VB.Timer coneCollision4 
      Interval        =   1
      Left            =   1230
      Top             =   2520
   End
   Begin VB.Timer coneCollision3 
      Interval        =   1
      Left            =   2475
      Top             =   2220
   End
   Begin VB.Timer coneCollision2 
      Interval        =   1
      Left            =   1980
      Top             =   2130
   End
   Begin VB.Timer coneCollision1 
      Interval        =   1
      Left            =   1605
      Top             =   2115
   End
   Begin VB.Timer coneCollision0 
      Interval        =   1
      Left            =   1245
      Top             =   2130
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3750
      Begin VB.PictureBox lightMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   870
         Left            =   105
         Picture         =   "Form1.frx":2FC6
         ScaleHeight     =   870
         ScaleWidth      =   330
         TabIndex        =   13
         Top             =   150
         Width           =   330
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Top             =   615
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Max             =   350
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Time:"
         Height          =   285
         Left            =   1425
         TabIndex        =   26
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label8 
         Caption         =   "MPH"
         Height          =   165
         Left            =   3135
         TabIndex        =   18
         Top             =   405
         Width           =   345
      End
      Begin VB.Label Label6 
         Caption         =   "350"
         Height          =   225
         Left            =   2835
         TabIndex        =   17
         Top             =   405
         Width           =   330
      End
      Begin VB.Label Label4 
         Caption         =   "Target:"
         Height          =   225
         Left            =   2295
         TabIndex        =   16
         Top             =   390
         Width           =   510
      End
      Begin VB.Label secLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1845
         TabIndex        =   15
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   315
         Left            =   1335
         TabIndex        =   12
         Top             =   435
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   525
         TabIndex        =   4
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "MPH"
         Height          =   165
         Left            =   2865
         TabIndex        =   3
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2325
         TabIndex        =   2
         Top             =   630
         Width           =   510
      End
   End
   Begin VB.Timer cone6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1980
      Top             =   4920
   End
   Begin VB.Timer cone5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1590
      Top             =   4920
   End
   Begin VB.Timer cone4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1170
      Top             =   4920
   End
   Begin VB.Timer cone3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2385
      Top             =   4530
   End
   Begin VB.Timer cone2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1980
      Top             =   4515
   End
   Begin VB.Timer cone1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1575
      Top             =   4515
   End
   Begin VB.Timer cone0 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1170
      Top             =   4515
   End
   Begin VB.Timer random 
      Interval        =   1000
      Left            =   1425
      Top             =   1350
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   6
      Left            =   1800
      Picture         =   "Form1.frx":355F
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   11
      Top             =   75
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   1335
      Picture         =   "Form1.frx":380B
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   10
      Top             =   90
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   270
      Picture         =   "Form1.frx":3AB7
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   9
      Top             =   90
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   2295
      Picture         =   "Form1.frx":3D63
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   8
      Top             =   60
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   3420
      Picture         =   "Form1.frx":400F
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   7
      Top             =   30
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   750
      Picture         =   "Form1.frx":42BB
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   6
      Top             =   105
      Width           =   210
   End
   Begin VB.PictureBox cone 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   2865
      Picture         =   "Form1.frx":4567
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   5
      Top             =   45
      Width           =   210
   End
   Begin VB.Timer border 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   450
      Top             =   1410
   End
   Begin VB.Timer movetim 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   630
      Top             =   1020
   End
   Begin VB.Timer tmrLane 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   645
      Top             =   1830
   End
   Begin VB.Timer tmrAccel 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   870
      Top             =   1425
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   11
      X1              =   2325
      X2              =   2325
      Y1              =   5955
      Y2              =   6525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   10
      X1              =   2325
      X2              =   2325
      Y1              =   5040
      Y2              =   5610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   9
      X1              =   2325
      X2              =   2325
      Y1              =   4155
      Y2              =   4725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   8
      X1              =   2325
      X2              =   2325
      Y1              =   3240
      Y2              =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   7
      X1              =   2325
      X2              =   2325
      Y1              =   2385
      Y2              =   2955
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   5
      X1              =   1065
      X2              =   1065
      Y1              =   5940
      Y2              =   6510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   4
      X1              =   1065
      X2              =   1065
      Y1              =   5025
      Y2              =   5595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   3
      X1              =   1065
      X2              =   1065
      Y1              =   4140
      Y2              =   4710
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   2
      X1              =   1065
      X2              =   1065
      Y1              =   3225
      Y2              =   3795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   1065
      X2              =   1065
      Y1              =   2370
      Y2              =   2940
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   6
      X1              =   2325
      X2              =   2325
      Y1              =   165
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   1065
      X2              =   1065
      Y1              =   180
      Y2              =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row1 As Integer
Dim Row2 As Integer
Dim Row3 As Integer
Dim Row4 As Integer
Dim Row5 As Integer
Dim Row6 As Integer
Dim accel As Integer
Dim lblSpeed As Integer
Dim moveCar As Integer
Dim mousepos As Integer
Dim c As Integer
Dim randCar As Integer
Dim maxVar As Integer
Dim maxSpeed As Integer
Dim i As Integer
Dim ms As Integer
Dim sec As Integer
Dim wsec As String
Dim min As Integer
Dim times As String
Dim erro As Integer
Dim Crazy$
Dim Expert$
Dim Advanced$
Dim Intermediate$
Dim Beginner$

Private Sub border_Timer()
If Image1.Left > 3000 Then
movetim.Enabled = False
End If
If Image1.Left > mousepos Then
Image1.Left = Image1.Left - 100
End If
If Image1.Left < 3000 Then
movetim.Enabled = True
End If
End Sub

Public Sub start_move()
movetim.Enabled = True
tmrLane.Enabled = True
tmrAccel.Enabled = True
border.Enabled = True
randomAccel.Enabled = True
Time.Enabled = True
End Sub


Private Sub Command2_Click()
start.Enabled = True
If Option1(0).Value = True Then
maxSpeed = 200
ElseIf Option1(1).Value = True Then
maxSpeed = 350
ElseIf Option1(2).Value = True Then
maxSpeed = 400
ElseIf Option1(3).Value = True Then
maxSpeed = 500
ElseIf Option1(4).Value = True Then
maxSpeed = 700
End If
ProgressBar1.Max = maxSpeed
Label6.Caption = maxSpeed
Frame2.Visible = False
End Sub

Private Sub cone0_Timer()
If cone(0).Top < 5500 Then
cone(0).Top = cone(0).Top + (accel / 6)
Else
cone(0).Top = 0
cone0.Enabled = False
End If
End Sub

Private Sub cone1_Timer()
If cone(1).Top < 5500 Then
cone(1).Top = cone(1).Top + (accel / 6)
Else
cone(1).Top = 0
cone1.Enabled = False
End If
End Sub

Private Sub cone2_Timer()
If cone(2).Top < 5500 Then
cone(2).Top = cone(2).Top + (accel / 6)
Else
cone(2).Top = 0
cone2.Enabled = False
End If
End Sub

Private Sub cone3_Timer()
If cone(3).Top < 5500 Then
cone(3).Top = cone(3).Top + (accel / 6)
Else
cone(3).Top = 0
cone3.Enabled = False
End If
End Sub

Private Sub cone4_Timer()
If cone(4).Top < 5500 Then
cone(4).Top = cone(4).Top + (accel / 6)
Else
cone(4).Top = 0
cone4.Enabled = False
End If
End Sub

Private Sub cone5_Timer()
If cone(5).Top < 5500 Then
cone(5).Top = cone(5).Top + (accel / 6)
Else
cone(5).Top = 0
cone5.Enabled = False
End If
End Sub

Private Sub cone6_Timer()
If cone(6).Top < 5500 Then
cone(6).Top = cone(6).Top + (accel / 6)
Else
cone(6).Top = 0
cone6.Enabled = False
End If
End Sub


Private Sub coneCollision0_Timer()
     If Image1.Left > cone(0).Left - (630 / 2) Then
          If Image1.Left < cone(0).Left + (630 / 2) Then
               If Image1.Top > cone(0).Top - (210 / 2) Then
                    If Image1.Top < cone(0).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision1_Timer()
     If Image1.Left > cone(1).Left - (630 / 2) Then
          If Image1.Left < cone(1).Left + (630 / 2) Then
               If Image1.Top > cone(1).Top - (210 / 2) Then
                    If Image1.Top < cone(1).Top + (210 / 2) Then
                         accel = accel / 2
                        lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision2_Timer()
     If Image1.Left > cone(2).Left - (630 / 2) Then
          If Image1.Left < cone(2).Left + (630 / 2) Then
               If Image1.Top > cone(2).Top - (210 / 2) Then
                    If Image1.Top < cone(2).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision3_Timer()
     If Image1.Left > cone(3).Left - (630 / 2) Then
          If Image1.Left < cone(3).Left + (630 / 2) Then
               If Image1.Top > cone(3).Top - (210 / 2) Then
                    If Image1.Top < cone(3).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision4_Timer()
     If Image1.Left > cone(4).Left - (630 / 2) Then
          If Image1.Left < cone(4).Left + (630 / 2) Then
               If Image1.Top > cone(4).Top - (210 / 2) Then
                    If Image1.Top < cone(4).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision5_Timer()
     If Image1.Left > cone(5).Left - (630 / 2) Then
          If Image1.Left < cone(5).Left + (630 / 2) Then
               If Image1.Top > cone(5).Top - (210 / 2) Then
                    If Image1.Top < cone(5).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub coneCollision6_Timer()
     If Image1.Left > cone(6).Left - (630 / 2) Then
          If Image1.Left < cone(6).Left + (630 / 2) Then
               If Image1.Top > cone(6).Top - (210 / 2) Then
                    If Image1.Top < cone(6).Top + (210 / 2) Then
                         accel = accel / 2
                         lblSpeed = lblSpeed / 2
                    End If
               End If
          End If
     End If
End Sub

Private Sub Form_Load()
i = 0
lightMain.Picture = light(0).Picture
Row1 = 0
Row2 = 1095
Row3 = 1950
Row4 = 2865
Row5 = 3750
Row6 = 4665
cone(0).Top = 0
cone(1).Top = 0
cone(2).Top = 0
cone(3).Top = 0
cone(4).Top = 0
cone(5).Top = 0
cone(6).Top = 0

End Sub

Private Sub movetim_Timer()
If Image1.Left < mousepos Then
Image1.Left = Image1.Left + 100
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousepos = X ' + 100
End Sub



Private Sub random_Timer()
    Randomize
    Select Case Int((Rnd * 7) + 1)
        Case 1
        Label3.Caption = 0
        cone0.Enabled = True
        
        Case 2
        Label3.Caption = 1
        cone1.Enabled = True
        
        Case 3
        Label3.Caption = 2
        cone2.Enabled = True
        
        Case 4
        Label3.Caption = 3
        cone3.Enabled = True
        
        Case 5
        Label3.Caption = 4
        cone4.Enabled = True
        
        Case 6
        Label3.Caption = 5
        cone5.Enabled = True
        
        Case 7
        Label3.Caption = 6
        cone6.Enabled = True
    End Select

End Sub

Private Sub randomAccel_Timer()
random.Interval = 100
randomAccel.Enabled = False
End Sub

Private Sub start_Timer()
lightMain.Picture = light(i).Picture
i = i + 1
If i = 3 Then
Command1.Caption = "GO!"
start.Enabled = False
start_move
End If
End Sub

Private Sub time_Timer()
sec = sec + 1
secLbl.Caption = sec
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = accel
Label1.Caption = lblSpeed
End Sub

Private Sub tmrAccel_Timer()
If accel < maxSpeed Then
accel = accel + 1
lblSpeed = lblSpeed + 1

End If
If ProgressBar1.Value = maxSpeed Then
tmrAccel.Enabled = False
MsgBox ("You Won!")
check_bs
End If
End Sub

Private Sub tmrLane_Timer()
If Row1 > 5000 Then
Row1 = 0
ElseIf Row2 > 5000 Then
Row2 = 0
ElseIf Row3 > 5000 Then
Row3 = 0
ElseIf Row4 > 5000 Then
Row4 = 0
ElseIf Row5 > 5000 Then
Row5 = 0
ElseIf Row6 > 5000 Then
Row6 = 0
End If
Row1 = Row1 + accel
Row2 = Row2 + accel
Row3 = Row3 + accel
Row4 = Row4 + accel
Row5 = Row5 + accel
Row6 = Row6 + accel
'----------------------------
Line1(0).Y1 = Row1
Line1(0).Y2 = Row1 + 600
Line1(6).Y1 = Row1
Line1(6).Y2 = Row1 + 600
'----------------------------
Line1(1).Y1 = Row2
Line1(1).Y2 = Row2 + 600
Line1(7).Y1 = Row2
Line1(7).Y2 = Row2 + 600
'----------------------------
Line1(2).Y1 = Row3
Line1(2).Y2 = Row3 + 600
Line1(8).Y1 = Row3
Line1(8).Y2 = Row3 + 600
'----------------------------
Line1(3).Y1 = Row4
Line1(3).Y2 = Row4 + 600
Line1(9).Y1 = Row4
Line1(9).Y2 = Row4 + 600
'----------------------------
Line1(4).Y1 = Row5
Line1(4).Y2 = Row5 + 600
Line1(10).Y1 = Row5
Line1(10).Y2 = Row5 + 600
'----------------------------
Line1(5).Y1 = Row6
Line1(5).Y2 = Row6 + 600
Line1(11).Y1 = Row6
Line1(11).Y2 = Row6 + 600
End Sub

Private Sub check_bs()
wsec = sec
Crazy$ = GetFromINI("BestScores", "Crazy", App.Path & "\data.ini")
Expert$ = GetFromINI("BestScores", "Expert", App.Path & "\data.ini")
Advanced$ = GetFromINI("BestScores", "Advanced", App.Path & "\data.ini")
Intermediate$ = GetFromINI("BestScores", "Intermediate", App.Path & "\data.ini")
Beginner$ = GetFromINI("BestScores", "Beginner", App.Path & "\data.ini")
If Option1(4).Value = True Then
    If wsec > Crazy$ Then
        Call WriteToINI("BestScores", "Crazy", wsec, App.Path & "\data.ini")
    End If
ElseIf Option1(3).Value = True Then
    If wsec > Expert$ Then
        Call WriteToINI("BestScores", "Expert", wsec, App.Path & "\data.ini")
    End If
ElseIf Option1(2).Value = True Then
    If wsec > Advanced$ Then
        Call WriteToINI("BestScores", "Advanced", wsec, App.Path & "\data.ini")
    End If
ElseIf Option1(1).Value = True Then
    If wsec > Intermediate$ Then
        Call WriteToINI("BestScores", "Intermediate", wsec, App.Path & "\data.ini")
    End If
ElseIf Option1(0).Value = True Then
    If wsec > Beginner$ Then
        Call WriteToINI("BestScores", "Beginner", wsec, App.Path & "\data.ini")
    End If
End If
stop_move
frmBestTimes.Show
End Sub

Public Sub stop_move()
random.Enabled = False
movetim.Enabled = False
tmrLane.Enabled = False
tmrAccel.Enabled = False
border.Enabled = False
randomAccel.Enabled = False
Time.Enabled = False
Frame2.Visible = True
lightMain.Picture = light(0).Picture
i = 0
sec = 0
cone0.Enabled = False
cone1.Enabled = False
cone2.Enabled = False
cone3.Enabled = False
cone4.Enabled = False
cone5.Enabled = False
cone6.Enabled = False
cone(0).Top = 0
cone(1).Top = 0
cone(2).Top = 0
cone(3).Top = 0
cone(4).Top = 0
cone(5).Top = 0
cone(6).Top = 0
End Sub
