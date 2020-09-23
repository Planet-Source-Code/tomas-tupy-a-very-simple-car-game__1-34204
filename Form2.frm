VERSION 5.00
Begin VB.Form frmBestTimes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Best Times"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2055
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   9
      Top             =   150
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   375
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   6
      Top             =   810
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Index           =   4
      Left            =   210
      TabIndex        =   5
      Top             =   1035
      Width           =   555
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   810
      X2              =   810
      Y1              =   180
      Y2              =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   795
      X2              =   795
      Y1              =   180
      Y2              =   1290
   End
   Begin VB.Label Label7 
      Caption         =   "(Beginner)"
      Height          =   225
      Left            =   1125
      TabIndex        =   4
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label6 
      Caption         =   "(Intermediate)"
      Height          =   225
      Left            =   885
      TabIndex        =   3
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label5 
      Caption         =   "(Expert)"
      Height          =   210
      Left            =   1305
      TabIndex        =   2
      Top             =   405
      Width           =   585
   End
   Begin VB.Label Label4 
      Caption         =   "(Advanced)"
      Height          =   210
      Left            =   1020
      TabIndex        =   1
      Top             =   630
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "(Crazy)"
      Height          =   210
      Left            =   1365
      TabIndex        =   0
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmBestTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim Crazy$
    Dim Expert$
    Dim Advanced$
    Dim Intermediate$
    Dim Beginner$
    Crazy$ = GetFromINI("BestScores", "Crazy", App.Path & "\data.ini")
    Expert$ = GetFromINI("BestScores", "Expert", App.Path & "\data.ini")
    Advanced$ = GetFromINI("BestScores", "Advanced", App.Path & "\data.ini")
    Intermediate$ = GetFromINI("BestScores", "Intermediate", App.Path & "\data.ini")
    Beginner$ = GetFromINI("BestScores", "Beginner", App.Path & "\data.ini")
   Label1(0).Caption = Crazy$
   Label1(1).Caption = Expert$
   Label1(2).Caption = Advanced$
   Label1(3).Caption = Intermediate$
   Label1(4).Caption = Beginner$
End Sub
