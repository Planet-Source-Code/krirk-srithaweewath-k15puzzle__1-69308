VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Image"
   ClientHeight    =   7620
   ClientLeft      =   1695
   ClientTop       =   495
   ClientWidth     =   8865
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Index           =   2
      Left            =   135
      Max             =   255
      TabIndex        =   13
      Top             =   6390
      Width           =   1230
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Index           =   1
      Left            =   135
      Max             =   255
      TabIndex        =   12
      Top             =   6030
      Width           =   1230
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Streach"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5355
      TabIndex        =   9
      Top             =   5445
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   1125
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1440
      Left            =   135
      TabIndex        =   6
      Top             =   1530
      Width           =   3300
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   135
      Pattern         =   "*.JPG"
      TabIndex        =   5
      Top             =   3015
      Width           =   3300
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   3690
      MouseIcon       =   "Form2.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   334
      TabIndex        =   4
      ToolTipText     =   "¤ÅÔ¡à¾×èÍà»ÅÕèÂ¹ÃÙ»"
      Top             =   225
      Width           =   5040
      Begin VB.Label lblDisplayFont 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sample number colour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1530
         TabIndex        =   14
         Top             =   2160
         Width           =   2025
      End
      Begin VB.Image Image1 
         Height          =   5010
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   180
      Left            =   8145
      Top             =   5355
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use image"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   180
      TabIndex        =   1
      Top             =   765
      Width           =   1725
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Not use image"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Index           =   0
      LargeChange     =   10
      Left            =   135
      Max             =   255
      TabIndex        =   11
      Top             =   5670
      Width           =   1230
   End
   Begin VB.Label lblMotherColor 
      BackColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   1485
      TabIndex        =   17
      Top             =   6390
      Width           =   420
   End
   Begin VB.Label lblMotherColor 
      BackColor       =   &H0000FF00&
      Height          =   285
      Index           =   1
      Left            =   1485
      TabIndex        =   16
      Top             =   6030
      Width           =   420
   End
   Begin VB.Label lblMotherColor 
      BackColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   1485
      TabIndex        =   15
      Top             =   5670
      Width           =   420
   End
   Begin VB.Shape ShapeShowColor 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1005
      Left            =   2070
      Top             =   5670
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Number colour"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   5310
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   1410
      Left            =   45
      Top             =   5400
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show sample"
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
      Left            =   5400
      TabIndex        =   8
      Top             =   5940
      Width           =   1185
   End
   Begin VB.Line Line4 
      X1              =   45
      X2              =   3555
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line Line3 
      X1              =   3555
      X2              =   3555
      Y1              =   900
      Y2              =   5175
   End
   Begin VB.Line Line2 
      X1              =   45
      X2              =   45
      Y1              =   900
      Y2              =   5175
   End
   Begin VB.Line Line1 
      X1              =   3555
      X2              =   45
      Y1              =   900
      Y2              =   900
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
      Left            =   855
      MouseIcon       =   "Form2.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7110
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancle"
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
      Index           =   1
      Left            =   2205
      MouseIcon       =   "Form2.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   7125
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   0
      Left            =   315
      MouseIcon       =   "Form2.frx":03F6
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":0548
      Stretch         =   -1  'True
      Top             =   6975
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   1
      Left            =   1800
      MouseIcon       =   "Form2.frx":223F
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":2391
      Stretch         =   -1  'True
      Top             =   6975
      Width           =   1410
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldFileName As String
Dim mvarFilename As String
Dim mvarFileNameOnly As String
Private Sub Check1_Click()
      Dim i As Integer
      Dim j As Integer
      Image1.Picture = Picture2.Picture
      Image1.Visible = Check1.Value
      Form1.Image1.Visible = Check1.Value
End Sub

Private Sub Dir1_Change()
      File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
      Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
                        If Right(Dir1.Path, 1) <> "\" Then
                             Picture2.Picture = LoadPicture(Dir1.Path & "\" & File1.List(File1.ListIndex))
                             Image1.Picture = Picture2.Picture
                             
                        Else
                              Picture2.Picture = LoadPicture(Dir1.Path & File1.List(File1.ListIndex))
                              Image1.Picture = Picture2.Picture
                        End If
                        mvarFileNameOnly = File1.List(File1.ListIndex)
End Sub

Private Sub File1_DblClick()
      Image3_Click 0
End Sub

Private Sub Form_Activate()
    Me.Timer4.Enabled = True
    
End Sub

Private Sub Form_Load()
      Me.Icon = Form1.Icon
      Form1.blnChangePicture = False

      'mvarFilename = GetSetting("K15Puzzle", "OldFilename", "001")
      GetSettingRegistry
    'Drive1.Enabled = False
    '  Drive1.BackColor = vbButtonFace
    '  Dir1.Enabled = False
    '  Dir1.BackColor = vbButtonFace
    '  File1.Enabled = False
    '  File1.BackColor = vbButtonFace
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      SaveSettingRegistry
      
      
End Sub
Private Sub GetSettingRegistry()
      Drive1.Drive = GetSetting("K15Puzzle", "Drive", "001", Drive1.Drive)
      Dir1.Path = GetSetting("K15Puzzle", "Dir", "001", Dir1.Path)
      File1.Path = GetSetting("K15Puzzle", "file", "001", File1.Path)
      
      Check1.Value = GetSetting("K15Puzzle", "STREACH", "001", Check1.Value)
      
      Option1.Value = GetSetting("K15Puzzle", "HavePic", "001", Option1.Value)
      mvarFileNameOnly = GetSetting("K15Puzzle", "FilenameOnly", "001", mvarFileNameOnly)
      HScroll1(0).Value = GetSetting("K15Puzzle", "Color", "Red", HScroll1(0).Value)
      HScroll1(1).Value = GetSetting("K15Puzzle", "Color", "Green", HScroll1(1).Value)
      HScroll1(2).Value = GetSetting("K15Puzzle", "Color", "Blue", HScroll1(2).Value)
                  
      Option2.Value = Not Option1.Value
      If Option1.Value Then
            Option1_Click
      Else
            Option2_Click
      End If
      
      OldFileName = GetSetting("K15Puzzle", "OldFilename", "001", mvarFilename)
      
      On Error Resume Next
      Picture2.Picture = LoadPicture(OldFileName)
      Check1_Click
      Dim i As Integer
      For i = 0 To File1.ListCount - 1
            If File1.List(i) = mvarFileNameOnly Then File1.ListIndex = i
      Next i
End Sub
Private Sub SaveSettingRegistry()
      SaveSetting "K15Puzzle", "Drive", "001", Drive1.Drive
      SaveSetting "K15Puzzle", "Dir", "001", Dir1.Path
      SaveSetting "K15Puzzle", "file", "001", File1.Path
      SaveSetting "K15Puzzle", "HavePic", "001", Option1.Value
      SaveSetting "K15Puzzle", "OldFilename", "001", mvarFilename
      SaveSetting "K15Puzzle", "STREACH", "001", Check1.Value
      SaveSetting "K15Puzzle", "FilenameOnly", "001", mvarFileNameOnly
      SaveSetting "K15Puzzle", "Color", "Red", HScroll1(0).Value
      SaveSetting "K15Puzzle", "Color", "Green", HScroll1(1).Value
      SaveSetting "K15Puzzle", "Color", "Blue", HScroll1(2).Value
      'SaveSetting "K15Puzzle", "", "001", Drive1.Drive
End Sub

Private Sub HScroll1_Change(Index As Integer)
      ShapeShowColor.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
      lblDisplayFont.ForeColor = ShapeShowColor.BackColor
End Sub

Private Sub Image3_Click(Index As Integer)
      Select Case Index
            Case 0
                  If Option1.Value = True Then
                        
                        mvarFilename = ""
                  Else
                        'Form1.blnChangePicture = True
                        If File1.ListIndex = -1 Then
                              MsgBox "You must choose picture ", vbOKOnly + vbInformation, "K15Puzzle"
                              Exit Sub
                        End If
                        If Right(Dir1.Path, 1) <> "\" Then
                              mvarFilename = Dir1.Path & "\" & File1.List(File1.ListIndex)
                        Else
                              mvarFilename = Dir1.Path & File1.List(File1.ListIndex)
                        End If
                  End If
                        
                        'If mvarFilename <> OldFileName Then
                              Form1.blnChangePicture = True
                        'End If
                        Form1.PicFileName = mvarFilename
                        
                        
                  Image3_Click 1
            Case 1
                  Unload Me
      End Select
End Sub

Private Sub Label2_Click(Index As Integer)
    Image3_Click Index
    
End Sub

Private Sub Option1_Click()
      Drive1.Enabled = False
      Drive1.BackColor = vbButtonFace
      Dir1.Enabled = False
      Dir1.BackColor = vbButtonFace
      File1.Enabled = False
      File1.BackColor = vbButtonFace
      Drive1.ForeColor = &H808080
      Dir1.ForeColor = &H808080
      File1.ForeColor = &H808080
      Picture2.Picture = LoadPicture()
      Image1.Picture = Picture2.Picture
      mvarFilename = ""
      mvarFileNameOnly = ""
      
End Sub

Private Sub Option2_Click()
      Drive1.Enabled = True
      Dir1.Enabled = True
      File1.Enabled = True
      Drive1.BackColor = vbWhite
      Dir1.BackColor = vbWhite
      File1.BackColor = vbWhite
      Drive1.ForeColor = vbBlack
      Dir1.ForeColor = vbBlack
      File1.ForeColor = vbBlack
End Sub

Private Sub Picture2_Resize()
      Image1.Move 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
End Sub

Private Sub Timer4_Timer()
      'EffectButton
     
      Timer4.Enabled = False
End Sub
