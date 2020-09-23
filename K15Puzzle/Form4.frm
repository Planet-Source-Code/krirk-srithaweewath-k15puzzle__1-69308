VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Score"
   ClientHeight    =   5970
   ClientLeft      =   3270
   ClientTop       =   1215
   ClientWidth     =   7080
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   90
      ScaleHeight     =   4245
      ScaleWidth      =   6945
      TabIndex        =   2
      Top             =   855
      Width           =   6945
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   42
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   0
         Left            =   1125
         TabIndex        =   41
         Top             =   135
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   0
         Left            =   4275
         TabIndex        =   40
         Top             =   135
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   0
         Left            =   5580
         TabIndex        =   39
         Top             =   135
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   38
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   1
         Left            =   1125
         TabIndex        =   37
         Top             =   540
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   1
         Left            =   4275
         TabIndex        =   36
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   1
         Left            =   5580
         TabIndex        =   35
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   2
         Left            =   90
         TabIndex        =   34
         Top             =   945
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   2
         Left            =   1125
         TabIndex        =   33
         Top             =   945
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   2
         Left            =   4275
         TabIndex        =   32
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   2
         Left            =   5580
         TabIndex        =   31
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   3
         Left            =   1125
         TabIndex        =   29
         Top             =   1350
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   3
         Left            =   4275
         TabIndex        =   28
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   3
         Left            =   5580
         TabIndex        =   27
         Top             =   1350
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   4
         Left            =   90
         TabIndex        =   26
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   4
         Left            =   1125
         TabIndex        =   25
         Top             =   1755
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   4
         Left            =   4275
         TabIndex        =   24
         Top             =   1755
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   4
         Left            =   5580
         TabIndex        =   23
         Top             =   1755
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   5
         Left            =   90
         TabIndex        =   22
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   5
         Left            =   1125
         TabIndex        =   21
         Top             =   2160
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   5
         Left            =   4275
         TabIndex        =   20
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   5
         Left            =   5580
         TabIndex        =   19
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   6
         Left            =   90
         TabIndex        =   18
         Top             =   2565
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   6
         Left            =   1125
         TabIndex        =   17
         Top             =   2565
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   6
         Left            =   4275
         TabIndex        =   16
         Top             =   2565
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   6
         Left            =   5580
         TabIndex        =   15
         Top             =   2565
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   7
         Left            =   90
         TabIndex        =   14
         Top             =   2970
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   7
         Left            =   1125
         TabIndex        =   13
         Top             =   2970
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   7
         Left            =   4275
         TabIndex        =   12
         Top             =   2970
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   7
         Left            =   5580
         TabIndex        =   11
         Top             =   2970
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   8
         Left            =   90
         TabIndex        =   10
         Top             =   3375
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   8
         Left            =   1125
         TabIndex        =   9
         Top             =   3375
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   8
         Left            =   4275
         TabIndex        =   8
         Top             =   3375
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   8
         Left            =   5580
         TabIndex        =   7
         Top             =   3375
         Width           =   1275
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   9
         Left            =   90
         TabIndex        =   6
         Top             =   3780
         Width           =   1005
      End
      Begin VB.Label lblName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   9
         Left            =   1125
         TabIndex        =   5
         Top             =   3780
         Width           =   3120
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   9
         Left            =   4275
         TabIndex        =   4
         Top             =   3780
         Width           =   1275
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   330
         Index           =   9
         Left            =   5580
         TabIndex        =   3
         Top             =   3780
         Width           =   1275
      End
   End
   Begin VB.Label lblMessageShow 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   510
      Left            =   270
      TabIndex        =   44
      Top             =   5310
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "K15Puzzle Hall of frame"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1575
      TabIndex        =   1
      Top             =   225
      Width           =   3930
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
      Left            =   5580
      MouseIcon       =   "Form4.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5445
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   0
      Left            =   5085
      MouseIcon       =   "Form4.frx":0152
      MousePointer    =   99  'Custom
      Picture         =   "Form4.frx":02A4
      Stretch         =   -1  'True
      Top             =   5265
      Width           =   1410
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type HallofFrameType
      UserName As String
      UserDate As String
      TimeToUse As Integer
End Type
Dim UserScore(1 To 10) As HallofFrameType
Dim NewPosition As Integer
Public NewScore As Integer
Private Sub NewData(TimeToUse As Integer)
      Dim i As Integer
      Dim j As Integer
      For i = 1 To 10
            If TimeToUse < UserScore(i).TimeToUse Then
                  lblMessageShow.Caption = " Congreatulation.Your score is no   " & i & "  please enter your name"
                  For j = 9 To i Step -1
                        UserScore(j + 1).UserName = UserScore(j).UserName
                        UserScore(j + 1).TimeToUse = UserScore(j).TimeToUse
                        UserScore(j + 1).UserDate = UserScore(j).UserDate
                  Next j
                  DisplayData
                  Text1.Visible = True
                  
                  lblTime(i - 1) = TimeToUse
                  lblDate(i - 1) = DateTime.Day(DateTime.Now) & "/" & DateTime.Month(DateTime.Now) & "/" & DateTime.Year(DateTime.Now)
                  NewPosition = i - 1
                  Text1.Move lblName(i - 1).Left, lblName(i - 1).Top, lblName(i - 1).Width, lblName(i - 1).Height
                  Text1.Text = " Please type your name and click enter"
                  'Text1.SetFocus
                  If Me.Visible Then
                        Text1.SetFocus
                        Text1.SelStart = 0
                        Text1.SelLength = Len(Text1.Text)
                  End If
                  Exit Sub
            End If
      Next i
            If TimeToUse = 30000 Then
                  lblMessageShow.Caption = "  This is a 10  best score "
            Else
                  lblMessageShow.Caption = "Your time is  " & TimeToUse & " so not 1 in 10 by the way,I think you are good to try"
            End If
End Sub
Private Sub AddScore(NewScore As HallofFrameType)
      Dim i As Integer
      Dim j As Integer
      For i = 1 To 10
            If NewScore.TimeToUse < UserScore(i).TimeToUse Then

                  UserScore(i).UserName = NewScore.UserName
                  UserScore(i).UserDate = NewScore.UserDate
                  UserScore(i).TimeToUse = NewScore.TimeToUse
                  Exit Sub
            End If
      Next i
End Sub
Private Sub DisplayData()
      Dim i As Integer
      For i = 1 To 10
            lblName(i - 1) = UserScore(i).UserName
            lblTime(i - 1) = UserScore(i).TimeToUse
            lblDate(i - 1) = UserScore(i).UserDate
      Next i
      
End Sub
Private Sub PrepareData()
Dim i As Integer
      For i = 1 To 10
            lblNo(i - 1) = i
            '+++ Create Default Data
            lblName(i - 1) = "K15Puzzle Staff"
            lblTime(i - 1) = 12 * Int(i + 10)
            lblDate(i - 1) = DateTime.Day(DateTime.Now) & "/" & DateTime.Month(DateTime.Now) & "/" & DateTime.Year(DateTime.Now)
            '+++
            '+++ Get Data from Registry
            '      SaveSetting "K15Puzzle", "HallofFrame", "Name" & i, lblName(i - 1).Caption
            'SaveSetting "K15Puzzle", "HallofFrame", "TimetoUse" & i, lblTime(i - 1).Caption
            'SaveSetting "K15Puzzle", "HallofFrame", "Date" & i, lblDate(i - 1).Caption
            lblName(i - 1) = GetSetting("K15Puzzle", "HallofFrame", "Name" & i, lblName(i - 1))
            lblTime(i - 1) = GetSetting("K15Puzzle", "HallofFrame", "TimetoUse" & i, lblTime(i - 1))
            lblDate(i - 1) = GetSetting("K15Puzzle", "HallofFrame", "Date" & i, lblDate(i - 1))
            '+++
            '+++ Contain to Variable
             UserScore(i).TimeToUse = lblTime(i - 1).Caption
             UserScore(i).UserDate = lblDate(i - 1)
             UserScore(i).UserName = lblName(i - 1)
             '+++
      Next i
End Sub

Private Sub Form_Activate()
  PrepareData
      NewData NewScore
End Sub

Private Sub Form_Click()
'            NewData InputBox("·´ÊÍº¡ÒÃáÊ´§¤Ðá¹¹", "ãÊè¤èÒà¾×èÍ·´ÊÍº", 100)
End Sub

Private Sub Form_Load()
      Me.Icon = Form1.Icon
      
    
      
End Sub
Private Sub SavetoRegistry()
      Dim i As Integer
      For i = 1 To 10
            SaveSetting "K15Puzzle", "HallofFrame", "Name" & i, lblName(i - 1).Caption
            SaveSetting "K15Puzzle", "HallofFrame", "TimetoUse" & i, lblTime(i - 1).Caption
            SaveSetting "K15Puzzle", "HallofFrame", "Date" & i, lblDate(i - 1).Caption
      Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
      SavetoRegistry
End Sub

Private Sub Image3_Click(Index As Integer)

      Text1_KeyDown vbKeyReturn, 0
      Unload Me
End Sub

Private Sub Label2_Click(Index As Integer)
    Image3_Click Index
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim NewScore As HallofFrameType
      If KeyCode = vbKeyReturn Then
            With NewScore
                  lblName(NewPosition) = Text1.Text
                  .UserName = lblName(NewPosition).Caption
                  .TimeToUse = lblTime(NewPosition).Caption
                  .UserDate = lblDate(NewPosition).Caption
                  AddScore NewScore
                  Text1.Visible = False
                  DisplayData
            End With
      End If
End Sub
