VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "K15Puzzle Help"
   ClientHeight    =   6060
   ClientLeft      =   -405
   ClientTop       =   1110
   ClientWidth     =   4995
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblRule 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   360
      TabIndex        =   5
      Top             =   2205
      Width           =   4200
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   " How to play"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   4
      Top             =   1845
      Width           =   1110
   End
   Begin VB.Shape Shape2 
      Height          =   3300
      Left            =   225
      Top             =   1935
      Width           =   4605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   2295
      MouseIcon       =   "Form3.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5490
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "http://DevGod.blogspot.com"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   330
      MouseIcon       =   "Form3.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label LabelAblut 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Form3.frx":02A4
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   360
      TabIndex        =   1
      Top             =   585
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " About"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Top             =   225
      Width           =   690
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   225
      Top             =   315
      Width           =   4605
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   0
      Left            =   1755
      MouseIcon       =   "Form3.frx":034F
      MousePointer    =   99  'Custom
      Picture         =   "Form3.frx":04A1
      Stretch         =   -1  'True
      Top             =   5355
      Width           =   1410
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
      Me.Icon = Form1.Icon
      lblRule.Caption = "           When you begin game,This game have 15 cells and each  cell is  swap" & vbCrLf _
      & " Goal of game is you must sort all of 15 cells to be a start game" & vbCrLf _
      & " You can click cell to can move or you just you arrow key to move " & vbCrLf _
      

End Sub

Private Sub Image3_Click(Index As Integer)
      Label2_Click Index
End Sub

Private Sub Label2_Click(Index As Integer)
      Unload Me
      
End Sub

Private Sub Label3_Click()
        
    Shell ("c:\Program Files\Internet Explorer\IEXPLORE.EXE http://DevGod.blogspot.com")
    
    
    
    
End Sub
