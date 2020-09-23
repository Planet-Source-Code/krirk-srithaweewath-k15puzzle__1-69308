VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "K15Puzzle"
   ClientHeight    =   5985
   ClientLeft      =   690
   ClientTop       =   855
   ClientWidth     =   8955
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8955
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Left            =   8070
      Top             =   5190
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hide image"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   4860
      Width           =   1995
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   5085
      Top             =   3735
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show number"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   4545
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5490
      Top             =   3285
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   8820
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   6
      Top             =   3285
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4980
      Left            =   4500
      MouseIcon       =   "Form1.frx":628A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":63DC
      ScaleHeight     =   330
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   5
      ToolTipText     =   "click for change image"
      Top             =   45
      Width           =   6345
      Begin VB.Image Image1 
         Height          =   5010
         Left            =   -30
         Picture         =   "Form1.frx":249D1
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   6315
      End
   End
   Begin VB.Timer Timer2 
      Left            =   6885
      Top             =   3285
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5085
      Top             =   4230
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   2250
      TabIndex        =   2
      Top             =   4500
      Width           =   1860
      Begin VB.Shape Shape1 
         Height          =   600
         Left            =   45
         Top             =   45
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hit"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3890
      Left            =   45
      ScaleHeight     =   3885
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   45
      Width           =   4155
      Begin VB.CommandButton Cmd 
         BackColor       =   &H00FFFFC0&
         Height          =   1050
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   1050
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
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
      Index           =   4
      Left            =   6840
      MouseIcon       =   "Form1.frx":42FC6
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5490
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Index           =   3
      Left            =   5355
      MouseIcon       =   "Form1.frx":43118
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5490
      Width           =   330
   End
   Begin VB.Image ImagePicClick 
      Height          =   570
      Left            =   5040
      MouseIcon       =   "Form1.frx":4326A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":433BC
      Stretch         =   -1  'True
      Top             =   2250
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Image ImagePicReal 
      Height          =   570
      Left            =   4680
      MouseIcon       =   "Form1.frx":45116
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":45268
      Stretch         =   -1  'True
      Top             =   1350
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Index           =   2
      Left            =   3375
      MouseIcon       =   "Form1.frx":46F5F
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5490
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   2
      Left            =   2925
      MouseIcon       =   "Form1.frx":470B1
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":47203
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change image"
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
      Left            =   1710
      MouseIcon       =   "Form1.frx":48EFA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5490
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Newgame"
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
      Left            =   300
      MouseIcon       =   "Form1.frx":4904C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5490
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   0
      Left            =   45
      MouseIcon       =   "Form1.frx":4919E
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":492F0
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   1
      Left            =   1485
      MouseIcon       =   "Form1.frx":4AFE7
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4B139
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   3
      Left            =   4995
      MouseIcon       =   "Form1.frx":4CE30
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4CF82
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   570
      Index           =   4
      Left            =   6435
      MouseIcon       =   "Form1.frx":4EC79
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4EDCB
      Stretch         =   -1  'True
      Top             =   5310
      Width           =   1410
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ScrCopy = &HCC0020
Const FormWidthNotShow = 4500
Const FormWidthShow = 8985
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim Pindex As Integer
Dim blnStatusBegin As Boolean
Dim Hit As Integer
Dim TiBe As Long

Public PicFileName As String
Public blnChangePicture As Boolean
Private Type POINTOak
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

Private Type R        ' For Row and Column
      Row As Integer
      Column As Integer
End Type
Private Type DetailAboutTable
       CWidth As Single
       CHeight As Single
       NumPerRow As Integer
       All As Integer
       rBlank As R
       rSim(1 To 4, 1 To 4) As R
       ModePicture As Boolean
       IsStart As Boolean
       mustHaveNumberonCell As Boolean
End Type
Private Detail As DetailAboutTable
Dim Ccount As Integer
Private C(1 To 4, 1 To 4) As Object ' For each command
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen
Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Const SM_CYCAPTION = 4 'Height of windows caption
Const SM_CXBORDER = 5 'Width of no-sizable borders
Const SM_CYBORDER = 6 'Height of non-sizable borders
Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Const SM_CYHTHUMB = 9 'Height of scroll box on horizontal scroll bar
Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Const SM_CXICON = 11 'Width of standard icon
Const SM_CYICON = 12 'Height of standard icon
Const SM_CXCURSOR = 13 'Width of standard cursor
Const SM_CYCURSOR = 14 'Height of standard cursor
Const SM_CYMENU = 15 'Height of menu
Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Const SM_DEBUG = 22 'True if deugging version of windows is running
Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Const SM_CXMIN = 28 'Minimum width of window
Const SM_CYMIN = 29 'Minimum height of window
Const SM_CXSIZE = 30 'Width of title bar bitmaps
Const SM_CYSIZE = 31 'height of title bar bitmaps
Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Const SM_CXDOUBLECLK = 36 'double click width
Const SM_CYDOUBLECLK = 37 'double click height
Const SM_CXICONSPACING = 38 'width between desktop icons
Const SM_CYICONSPACING = 39 'height between desktop icons
Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Const SM_CMETRICS = 44 'Number of system metrics
Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Const SM_CXMENUSIZE = 54 'width of button on menu bar
Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Const SM_CYMENUSIZE = 55 'height of button on menu bar
Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.
Dim ArrayPositionRandom(1 To 49) As Integer
Dim blnStatusNotFirstRun As Boolean
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Enum enumTimeDelay
      Ti3 = 1
      Ti4 = 2
      Ti5 = 3
      Ti6 = 4
      Ti8 = 5
End Enum
Private Type TimeDalayType
      TimerDelay As enumTimeDelay
      InitiVal As Integer
End Type
Dim IsCanClick As Boolean


Dim TimeDelay As TimeDalayType


Private Function ConVertRealtoSim(Index As Integer) As R
Dim TempR As R
Dim i As Integer
Dim j As Integer
      TempR = ConvertToArray(Index)
      For i = 1 To 4
            For j = 1 To 4
                              
            Next j
      Next i
End Function
Private Function ConvertToArray(Index As Integer) As R
'Purpose : For Convert 1 Dimension Array to 2 Dimension Array
'Accept : index of 1 Dimension
'Return :Column and Row of Varible type R

      Dim rTemp As R
      '      List1.AddItem "Real   = " & i \ 4 + 1 & "          " & i Mod 4 + 1
      '      List1.AddItem "Sim   = " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Row & "      " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Column
      rTemp.Row = Index \ Detail.NumPerRow + 1
      rTemp.Column = Index Mod Detail.NumPerRow + 1
      ConvertToArray = rTemp
End Function
Private Sub MovePosition(Row As Integer, Col As Integer, Ob As R, Bla As R)
'Purpose : For Change Blank Position and Move Position
'Accept  :Row,Col of BlankPosition  and Position of Cmd to move by pass Ob
'Return : None
      If Col = 0 And Row = 0 Then Exit Sub
      If Col <> 0 Then
            C(Ob.Row, Ob.Column).Left = (Col - 1) * C(1, 1).Width
      End If
      
      If Row <> 0 Then
            C(Ob.Row, Ob.Column).Top = (Row - 1) * C(1, 1).Height
      End If
      If Row <> 0 Or Col <> 0 Then
            With Detail
                  
                  .rSim(Ob.Row, Ob.Column).Column = .rBlank.Column
                  .rSim(Ob.Row, Ob.Column).Row = .rBlank.Row
'                  List1.AddItem "Blank " & Bla.Row & "       " & Bla.Column
                  
                  
                  .rBlank.Column = Bla.Column
                  .rBlank.Row = Bla.Row
            End With
      End If
      
End Sub





Private Sub Check1_Click()
      If Detail.mustHaveNumberonCell <> Check1.Value Then
            Detail.mustHaveNumberonCell = Check1.Value
            SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

            DrawPicture
            Dim i As Integer

      End If
End Sub

Private Sub Check2_Click()
      If Check2.Value = 0 Then
            Me.Width = FormWidthShow
      Else
            Me.Width = FormWidthNotShow
      End If
End Sub

Private Sub Cmd_GotFocus(Index As Integer)
      'Timer3.Enabled = True
      TimeDelay.TimerDelay = Ti3
      Timer3.Interval = 50
      Timer3.Enabled = True
      
End Sub

Private Sub Cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
                  If Button = vbRightButton Then Exit Sub
                  If Not IsCanClick Then
                            If MsgBox("Do you want to new game ?", vbOKCancel) = vbOK Then
                            Image3_Click (0)
                            
                            End If
                            Exit Sub
                  End If
                  
            Dim blnStatusSel As Boolean
            Dim rClick As R                          ' For Position of Cmd
            Dim rTemp As R
            Dim MoveRow As Integer          ' For Row To MOve
            Dim MoveCol As Integer            ' For Col to Move
            Hit = Hit + 1
            Label1(0).Caption = "Hit            : " & Hit
            
            blnStatusSel = False
            i = Index
                       
            'List1.Clear
            'List1.SetFocus
            rClick = ConvertToArray(i)
            rTemp.Row = rClick.Row
            rTemp.Column = rClick.Column
            rClick.Row = Detail.rSim(rTemp.Row, rTemp.Column).Row
            rClick.Column = Detail.rSim(rTemp.Row, rTemp.Column).Column
            
      
      'List1.AddItem "Real   = " & i \ 4 + 1 & "          " & i Mod 4 + 1
      'List1.AddItem "Sim   = " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Row & "      " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Column
     
      ' Check for BlankPosition ++++++++++++++++++++++
      
      With Detail
      'List1.AddItem "Blank    " & .rBlank.Row & "      " & .rBlank.Column
      'List1.AddItem "Click  " & rClick.Row & "       " & rClick.Column
      Select Case .rBlank.Row
            Case rClick.Row + 1 And rClick.Column = .rBlank.Column
                  MoveRow = rClick.Row + 1
                  blnStatusSel = True
            Case rClick.Row - 1 And rClick.Column = .rBlank.Column
                  MoveRow = rClick.Row - 1
                  blnStatusSel = True
            Case rClick.Row
                  MoveRow = 0
      End Select
      If Not blnStatusSel Then
            Select Case .rBlank.Column
                  Case rClick.Column + 1 And rClick.Row = .rBlank.Row
                        MoveCol = rClick.Column + 1
                  Case rClick.Column - 1 And rClick.Row = .rBlank.Row
                        MoveCol = rClick.Column - 1
                  Case rClick.Column
                        MoveCol = 0
            End Select
      End If
      End With

      '++++++++++++ End Check ++++++++++++++++++++
      MovePosition MoveRow, MoveCol, rTemp, rClick
      If blnStatusBegin Then
            If CheckWin Then

                  Timer1.Enabled = False
                  SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                  IsCanClick = False
                  
                  Form4.NewScore = Format(Timer - TiBe, "##.00")
                  Form4.Show vbModal
                  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                  
            End If
      
      End If
      
      'Text1.Text = Text1.Text & Index
      'If Val(Text2) <> 49 Then Text1.Text = Text1.Text & ", "
      'Text2.Text = Val(Text2.Text) + 1
      'If Text2.Text = 49 Then MsgBox 49
      'Clipboard.SetText Text1.Text
End Sub

Private Sub InitGame()
      Dim i As Integer
      Dim j As Integer
      Dim rTemp As R
   '   Text1.Text = ""
'      Image1.Picture = Picture2.Picture
'     Image2.Picture = LoadPicture("C:\WINDOWS\Desktop\Untitled-3Yub.jpg")
'     MsgBox CheckCanMove
      'Preare Data about Table
                 Detail.All = 16
                Detail.CHeight = Form1.Cmd(0).Height
                Detail.CWidth = Form1.Cmd(0).Width
                Detail.NumPerRow = 4
                Detail.rBlank.Row = Detail.NumPerRow
                Detail.rBlank.Column = Detail.NumPerRow
      '++++++ End Prepare +++++++++++++++++
      
      
      
      
            'Cmd(0).Caption = 1
            Set C(1, 1) = Cmd(0)
            Cmd(0).Top = 0
            Cmd(0).Left = 0
        '    Check2.Value = 1

      For i = 1 To 15
            If Not blnStatusNotFirstRun Then
                  Load Cmd(i)
            End If
            Cmd(i).Left = Cmd(i - 1).Left + Cmd(i - 1).Width
            Cmd(i).Top = Cmd(i - 1).Top
            If i Mod 4 = 0 Then
                  Cmd(i).Top = Cmd(i - 1).Top + Cmd(i - 1).Height
                  Cmd(i).Left = Cmd(0).Left
            End If
            Cmd(i).Visible = True
            'Cmd(i).Caption = i + 1
            Cmd(i).Caption = ""
            rTemp = ConvertToArray(i)
            Set C(rTemp.Row, rTemp.Column) = Cmd(i)
         '   List1.AddItem i \ 4 + 1 & "        " & i Mod 4 + 1
      Next i
      Cmd(15).Visible = False
      Picture1.Height = Cmd(0).Height * 4 + 50
      Picture1.Width = Cmd(0).Width * 4 + 50
      
      For i = 1 To 4
            For j = 1 To 4
                  Detail.rSim(i, j).Column = j
                  Detail.rSim(i, j).Row = i
            Next j
      Next i
      Detail.rBlank.Row = 4
      Detail.rBlank.Column = 4
      Me.Show
'      Picture2.Width = Picture1.Width
    '  MsgBox Picture1.Height & " " & Picture1.Width
    
     Picture2.Width = Picture1.Width
     Picture2.Height = Picture1.Height
     'Image1.Height = Picture2.Height
     'Image1.Width = Picture2.Width
'          MsgBox CheckCanMove
  '    For i = 1 To 200
  '          CheckCanMove
  '    Next i
  Timer4.Interval = 100
  
  Timer4.Enabled = True
End Sub





Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim blnStatusSel As Boolean
      Dim i As Integer
      If Detail.IsStart Then
      Dim rClick As R                          ' For Position of Cmd
      Dim rTemp As R
      Dim MoveRow As Integer          ' For Row To MOve
      Dim MoveCol As Integer            ' For Col to Move
      'Hit = Hit + 1
      'Label1(0).Caption = "Hit            : " & Hit
      
      blnStatusSel = False
      'i = index
                 
      'List1.Clear
      'List1.SetFocus
      For i = 0 To 15
            rClick = ConvertToArray(i)
            If i = 10 Then
                  Dim j  As Integer
                  j = 3
            End If
            rTemp.Row = rClick.Row
            rTemp.Column = rClick.Column
            rClick.Row = Detail.rSim(rTemp.Row, rTemp.Column).Row
            rClick.Column = Detail.rSim(rTemp.Row, rTemp.Column).Column
            
            With Detail
      'List1.AddItem "Blank    " & .rBlank.Row & "      " & .rBlank.Column
      'List1.AddItem "Click  " & rClick.Row & "       " & rClick.Column
            Select Case .rBlank.Row
                  Case rClick.Row + 1 And rClick.Column = .rBlank.Column And KeyCode = vbKeyDown
                        MoveRow = rClick.Row + 1
                        blnStatusSel = True
                  Case rClick.Row - 1 And rClick.Column = .rBlank.Column And KeyCode = vbKeyUp
                        MoveRow = rClick.Row - 1
                        blnStatusSel = True
                  Case rClick.Row
                        MoveRow = 0
            End Select
            
            If Not blnStatusSel Then
                  Select Case .rBlank.Column
                        Case rClick.Column + 1 And rClick.Row = .rBlank.Row And KeyCode = vbKeyRight
                              MoveCol = rClick.Column + 1
                              blnStatusSel = True
                        Case rClick.Column - 1 And rClick.Row = .rBlank.Row And KeyCode = vbKeyLeft
                              MoveCol = rClick.Column - 1
                              blnStatusSel = True
                        Case rClick.Column
                              MoveCol = 0
                  End Select
            End If
      End With
            
            If MoveRow <> 0 Or MoveCol <> 0 Then
                  Exit For
            End If
      Next i
      
      
      
      'List1.AddItem "Real   = " & i \ 4 + 1 & "          " & i Mod 4 + 1
      'List1.AddItem "Sim   = " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Row & "      " & Detail.rSim(i \ 4 + 1, i Mod 4 + 1).Column
     
      ' Check for BlankPosition ++++++++++++++++++++++
      
      

      '++++++++++++ End Check ++++++++++++++++++++
      'MovePosition MoveRow, MoveCol, rTemp, rClick
      'If blnStatusBegin Then
      '      If CheckWin Then
      '            MsgBox "¤Ø³à¡è§ÁÒ¡"
      '      End If
      
      'End If
      
            'Select Case KeyCode
            '      Case vbKeyLeft
                                   
            '      Case vbKeyRight
            
            '      Case vbKeyUp
            
            '      Case vbKeyDown
            
            'End Select
            Me.blnChangePicture = True
            
            Label2_Click (1)
            
            If blnStatusSel Then
                  Cmd_MouseDown i, vbLeftButton, 0, 1, 1
            End If
      End If
End Sub

Private Sub Form_Load()
      InitGame
      'MsgBox CheckWin
      blnStatusNotFirstRun = True
      Dim i As Integer
      Dim strTemp As String
     ' For i = 21 To 49
     '       'strTemp = strTemp & " Num" & i & " as Integer,"
     '       'ArrayPositionRandom(20) = num20
     '       strTemp = strTemp & "ArrayPositionRandom(" & i & ") = " & "Num" & i & vbCrLf
     ' Next i
     ' Clipboard.SetText strTemp
    Me.Image1.Visible = True
'    Label2_Click (0)
    
    
     
      
End Sub

Private Sub Form_Resize()
      Debug.Print Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
      Dim i As Integer
      Dim j As Integer
      ' Return Resource
      For i = 1 To Detail.All - 1
            Unload Cmd(i)
      Next i
      
      For i = 1 To 4
            For j = 1 To 4
      '            Unload C(i, j)
                  Set C(i, j) = Nothing
                
            Next j
      Next i
      
End Sub
Private Sub DrawPicture()
                        Dim nWidth As Long
                        Dim nHeight As Long
                        Dim Xthick As Long
                        Dim Ythick As Long
                        Xthick = GetSystemMetrics(7)
                        Ythick = GetSystemMetrics(8)
                  
                         nWidth = Cmd(0).Width
                         nHeight = Cmd(0).Height
                         
                  
                         Picture3.Width = nWidth
                         Picture3.Height = nHeight
                  
                         Dim i As Integer
                         
                        Dim A As R
                        
                         For i = 0 To 14
      
                                    A = ConvertToArray(i)
                                    Dim P As POINTOak
                                    Dim P2 As POINTOak
                              
                             With P
                                    .X = 0 '(A.Column - 1) * Nwidth
                                    .Y = 0 ' (A.Row - 1) * NHeight
                                    .Height = nHeight
                                    .Width = nWidth
                              End With
                              nWidth = Picture3.ScaleWidth
                              nHeight = Picture3.ScaleHeight
      
                              With P2
                                    .X = Xthick + (A.Column - 1) * nWidth
                                    .Y = Ythick + (A.Row - 1) * nWidth
                                    .Height = nHeight
                                    .Width = nWidth
                              End With
                        

                               BitBlt Picture3.hdc, P.X, P.Y, P.Width, P.Height, Picture2.hdc, P2.X, P2.Y, ScrCopy

                              If Detail.mustHaveNumberonCell Then
                                    Picture3.Print ""
                               
                                    Picture3.FontName = "Microsoft Sans serif"
                                    Picture3.FontSize = 8
                           
                                    Picture3.ForeColor = RGB(GetSetting("K15Puzzle", "Color", "Red", 0), GetSetting("K15Puzzle", "Color", "Green", 0), GetSetting("K15Puzzle", "Color", "Blue", 0))
                                    Picture3.Print
                                    Picture3.Print "        " & i + 1
                         End If
                                    Picture3.Picture = Picture3.Image
                              
                            Cmd(i).Picture = Picture3.Picture
                              Picture3.Cls
                              

                        Next i
End Sub

Private Sub Image1_DblClick()
      Picture2_DblClick
End Sub

Private Sub Image3_Click(Index As Integer)
                        IsCanClick = True

                        Pindex = Index
                        Label2(Pindex).Top = Label2(Pindex).Top + 10
                        Label2(Pindex).Left = Label2(Pindex).Left + 10
                        Image3(Index).Picture = ImagePicReal.Picture
                  '      Timer4.Enabled = True
                        TimeDelay.TimerDelay = Ti4
                        Timer3.Interval = 180
                        Timer3.Enabled = True
                        
                        
      Select Case Index
            Case 0
                        Check2.Value = 0
                        Check2_Click
                        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                        Timer7.Enabled = True
                        
            Case 1
                  SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                  
                  Form2.Show vbModal
                  If Me.blnChangePicture Then
                        Picture2.Picture = LoadPicture(Me.PicFileName)
                        Image1.Picture = Picture2.Picture
                        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                        If Me.PicFileName = "" Then
                              Check1.Value = 1
                              Check1_Click
                        End If
                        
                        'DrawPicture
                        'Timer8.Enabled = True
                        TimeDelay.TimerDelay = Ti8
                        Timer3.Interval = 220
                        Timer3.Enabled = True
                        'DrawPicture
                        
                  End If
            Image3_Click 0
            
            Case 2

                  If MsgBox("Do you want to exit ", vbOKOnly + vbInformation, "K15Puzzle") = vbOK Then
                        Unload Me
                  End If
            Case 3
                  SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                  Form3.Show vbModal
                  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            Case 4
                  SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                  Form4.NewScore = 30000
                  Form4.Show vbModal
                  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
      End Select
End Sub
Private Sub FillPositionRandom(Num1 As Integer, Num2 As Integer, Num3 As Integer, num4 As Integer, num5 As Integer, num6 As Integer, num7 As Integer, num8 As Integer, num9 As Integer, num10 As Integer, num11 As Integer, num12 As Integer, num13 As Integer, num14 As Integer, num15 As Integer, num16 As Integer, num17 As Integer, num18 As Integer, num19 As Integer, num20 As Integer, Num21 As Integer, Num22 As Integer, Num23 As Integer, Num24 As Integer, Num25 As Integer, Num26 As Integer, Num27 As Integer, Num28 As Integer, Num29 As Integer, Num30 As Integer, Num31 As Integer, Num32 As Integer, Num33 As Integer, Num34 As Integer, Num35 As Integer, Num36 As Integer, Num37 As Integer, Num38 As Integer, Num39 As Integer, Num40 As Integer, Num41 As Integer, Num42 As Integer, Num43 As Integer, Num44 As Integer, Num45 As Integer, Num46 As Integer, Num47 As Integer, Num48 As Integer, Num49 As Integer)
                  
                  
                  ArrayPositionRandom(1) = Num1
                  ArrayPositionRandom(2) = Num2
                  ArrayPositionRandom(3) = Num3
                  ArrayPositionRandom(4) = num4
                  ArrayPositionRandom(5) = num5
                  ArrayPositionRandom(6) = num6
                  ArrayPositionRandom(7) = num7
                  ArrayPositionRandom(8) = num8
                  ArrayPositionRandom(9) = num9
                  ArrayPositionRandom(10) = num10
                  ArrayPositionRandom(11) = num11
                  ArrayPositionRandom(12) = num12
                  ArrayPositionRandom(13) = num13
                  ArrayPositionRandom(14) = num14
                  ArrayPositionRandom(15) = num15
                  ArrayPositionRandom(16) = num16
                  ArrayPositionRandom(17) = num17
                  ArrayPositionRandom(18) = num18
                  ArrayPositionRandom(19) = num19
                  ArrayPositionRandom(20) = num20
                  ArrayPositionRandom(21) = Num21
                  ArrayPositionRandom(22) = Num22
                  ArrayPositionRandom(23) = Num23
                  ArrayPositionRandom(24) = Num24
                  ArrayPositionRandom(25) = Num25
                  ArrayPositionRandom(26) = Num26
                  ArrayPositionRandom(27) = Num27
                  ArrayPositionRandom(28) = Num28
                  ArrayPositionRandom(29) = Num29
                  ArrayPositionRandom(30) = Num30
                  ArrayPositionRandom(31) = Num31
                  ArrayPositionRandom(32) = Num32
                  ArrayPositionRandom(33) = Num33
                  ArrayPositionRandom(34) = Num34
                  ArrayPositionRandom(35) = Num35
                  ArrayPositionRandom(36) = Num36
                  ArrayPositionRandom(37) = Num37
                  ArrayPositionRandom(38) = Num38
                  ArrayPositionRandom(39) = Num39
                  ArrayPositionRandom(40) = Num40
                  ArrayPositionRandom(41) = Num41
                  ArrayPositionRandom(42) = Num42
                  ArrayPositionRandom(43) = Num43
                  ArrayPositionRandom(44) = Num44
                  ArrayPositionRandom(45) = Num45
                  ArrayPositionRandom(46) = Num46
                  ArrayPositionRandom(47) = Num47
                  ArrayPositionRandom(48) = Num48
                  ArrayPositionRandom(49) = Num49
                  
End Sub

Private Sub Label2_Click(Index As Integer)
      Label2(Index).Top = Label2(Index).Top + 10
      Label2(Index).Left = Label2(Index).Left + 10
      Image3_Click Index
      
End Sub

Private Sub Picture2_DblClick()
            Image3_Click 1
End Sub

Private Sub Picture2_Resize()
      Image1.Height = Picture2.ScaleHeight
      Image1.Width = Picture2.ScaleWidth
End Sub

Private Sub Timer1_Timer()
      Label1(1).Caption = "Time         : " & Format(Timer - TiBe, "##.00")
      
End Sub

Private Sub Timer2_Timer()
      Dim Col As Integer
      Dim Row As Integer
Dim rTemp As R
      If Ccount > 50 Then
            Timer2.Enabled = True
            Ccount = 0
            Exit Sub
      End If
      With Detail.rBlank
            If .Column = 4 Then
                  Col = 3
            Else
                  Col = .Column + 1
            End If
            
            If .Row = 4 Then
                  Row = 3
            Else
                  Row = .Row + 1
            End If
            
      End With
      'cmd_click(
End Sub

Private Sub Timer3_Timer()
      '3,4,5,6,8
      'Dim TempenumTimeDelay As enumTimeDelay
      Static staticChoosePosition As Integer
      Dim ButtonSlide As Integer
      With TimeDelay
            Select Case True
                  Case .TimerDelay = Ti3
                        'Picture1.SetFocus
                        Timer3.Enabled = False
                  Case .TimerDelay = Ti4
                        EffectButton
                        Timer3.Enabled = False
                  Case .TimerDelay = Ti5
                              
                              staticChoosePosition = staticChoosePosition + 1
                              If staticChoosePosition = 50 Then
                                    staticChoosePosition = 0
                                    Timer3.Enabled = False
                                    Hit = 0
                                                Label1(0).Caption = "Hit            : " & Hit
                                    blnStatusBegin = True
                                    Detail.IsStart = True
                                    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
                                    Timer1.Enabled = True
                                    Exit Sub
                              End If
                              'ButtonSlide = CheckCanMove
                              '15 * Rnd
                                    Cmd_MouseDown ArrayPositionRandom(staticChoosePosition), 0, 0, 0, 0
                                    Timer3.Enabled = False
                                    'Timer6.Enabled = True
                                    TimeDelay.TimerDelay = Ti6
                         
                                    Timer3.Interval = 10
                                    Timer3.Enabled = True
                                    
                  Case .TimerDelay = Ti6
                              'Timer6.Enabled = False
                              TimeDelay.TimerDelay = Ti5
                              Timer3.Interval = 50
                              Timer3.Enabled = True
                              
                  Case .TimerDelay = Ti8
                              DrawPicture
                              Timer3.Enabled = False
            End Select
      End With
      
      'Picture1.SetFocus
      'Timer3.Enabled = False
End Sub
Private Sub EffectButton()
      Image3(Pindex).Picture = ImagePicClick.Picture 'LoadPicture("C:\WINDOWS\Desktop\testbuttonfork15.jpg")
      Label2(Pindex).Top = Label2(Pindex).Top - 10
      Label2(Pindex).Left = Label2(Pindex).Left - 10
      
End Sub

Private Function CheckWin() As Boolean
      Dim Result As Boolean
      Result = False
      Dim i As Integer
      Dim j As Integer
      For i = 1 To 4
            For j = 1 To 4
            If Detail.rSim(i, j).Row <> i Then GoTo PROCESS_NOTWIN
            
            If Detail.rSim(i, j).Column <> j Then GoTo PROCESS_NOTWIN
            
            Next j
      Next i
      'MsgBox Detail.rSim(1, 1).Row
      'MsgBox Detail.rSim(1, 1).Column
      CheckWin = True
      Exit Function
PROCESS_NOTWIN:
      CheckWin = False
End Function






Private Sub Timer4_Timer()
    Me.Timer4.Enabled = False
    Check1_Click
    
End Sub

Private Sub Timer7_Timer()
                        InitGame
                              blnStatusBegin = False
                       
                        TiBe = Timer
                         
                            
                        DrawPicture
                  
                        Randomize
                        Dim TempChoose As Integer
                        Randomize
                        TempChoose = 10 * Rnd
                              Select Case TempChoose
                                    Case 0
                                          'FillPositionRandom 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12
                                          'FillPositionRandom 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15
                                          'FillPositionRandom 14, 10, 6, 5, 9, 13, 12, 8, 13, 9, 4, 0, 1, 4, 5, 2, 3, 7, 11, 6
                                        '  FillPositionRandom 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14, 14
                                          FillPositionRandom 14, 10, 6, 5, 9, 8, 4, 0, 1, 9, 5, 2, 3, 7, 11, 6, 10, 13, 12, 4, 8, 5, 0, 1, 9, 3, 2, 11, 6, 14, 13, 10, 5, 12, 10, 5, 11, 0, 12, 11, 0, 6, 7, 2, 3, 12, 1, 9, 12
                                    Case 1
                                          FillPositionRandom 14, 10, 6, 5, 9, 8, 4, 0, 1, 9, 5, 2, 3, 7, 11, 6, 10, 13, 8, 4, 12, 8, 4, 10, 6, 11, 2, 5, 9, 3, 5, 6, 13, 14, 11, 13, 10, 9, 3, 1, 0, 3, 1, 5, 6, 10, 9, 1, 3
                                    Case 2
                                          FillPositionRandom 14, 13, 12, 8, 4, 0, 1, 2, 3, 7, 11, 14, 13, 10, 6, 3, 2, 1, 0, 4, 9, 6, 14, 11, 7, 2, 1, 0, 4, 9, 8, 12, 10, 14, 3, 1, 0, 6, 5, 6, 10, 14, 13, 11, 3, 10, 8, 9, 6
                                    Case 3
                                           FillPositionRandom 14, 10, 6, 7, 3, 2, 7, 7, 2, 3, 11, 6, 10, 14, 6, 10, 14, 6, 6, 13, 9, 5, 7, 14, 5, 8, 0, 4, 7, 8, 5, 14, 11, 10, 14, 13, 6, 14, 13, 5, 9, 6, 5, 9, 4, 12, 6, 4, 9
                                    Case 4
                                           FillPositionRandom 14, 13, 9, 5, 2, 1, 2, 6, 10, 11, 7, 10, 1, 4, 0, 2, 4, 5, 11, 1, 10, 7, 14, 13, 1, 11, 9, 12, 8, 9, 5, 0, 2, 4, 6, 10, 7, 3, 3, 7, 10, 0, 6, 0, 2, 4, 0, 2, 5
                                    Case 5
                                          FillPositionRandom 11, 7, 3, 2, 1, 0, 4, 8, 12, 13, 14, 11, 7, 3, 2, 1, 0, 4, 8, 12, 13, 14, 11, 7, 3, 10, 9, 13, 12, 8, 4, 0, 1, 2, 10, 3, 7, 11, 14, 12, 8, 5, 6, 10, 3, 7, 11, 14, 12
                                    Case 6
                                          FillPositionRandom 14, 13, 12, 8, 9, 5, 1, 2, 6, 10, 11, 7, 10, 1, 4, 0, 2, 4, 5, 11, 1, 10, 3, 6, 4, 5, 11, 1, 10, 3, 7, 14, 13, 10, 1, 12, 8, 9, 12, 11, 5, 4, 3, 1, 14, 7, 1, 5, 0
                                    Case 7
                                          FillPositionRandom 14, 10, 9, 5, 4, 0, 1, 4, 6, 9, 11, 14, 10, 13, 12, 8, 5, 6, 9, 2, 3, 7, 14, 10, 13, 11, 6, 9, 0, 1, 4, 3, 7, 14, 2, 6, 9, 12, 8, 5, 12, 8, 11, 9, 6, 2, 10, 6, 2
                                    Case 8
                                          FillPositionRandom 14, 14, 11, 11, 14, 10, 11, 11, 10, 14, 14, 13, 13, 10, 6, 6, 11, 7, 3, 2, 2, 2, 1, 5, 5, 5, 6, 6, 6, 11, 9, 13, 10, 14, 7, 9, 11, 6, 4, 0, 5, 1, 6, 4, 13, 8, 12, 10, 8
                                          
                                    Case 9
                                          FillPositionRandom 14, 10, 6, 5, 4, 8, 12, 13, 10, 6, 5, 7, 3, 2, 1, 4, 9, 5, 11, 3, 2, 1, 4, 9, 5, 12, 13, 10, 6, 11, 7, 5, 8, 13, 10, 6, 11, 7, 5, 8, 13, 0, 9, 4, 1, 2, 3, 14, 7
                                    Case Else
                                          FillPositionRandom 11, 10, 9, 5, 6, 2, 3, 7, 2, 9, 5, 13, 14, 5, 10, 2, 9, 3, 1, 6, 13, 10, 3, 9, 2, 3, 5, 14, 10, 8, 12, 10, 8, 5, 14, 11, 3, 14, 9, 2, 7, 1, 2, 13, 6, 2, 13, 7, 1

                              End Select
                        'Timer5.Enabled = True
                        Timer3.Interval = 1
                        TimeDelay.TimerDelay = Ti5
                        Timer3.Enabled = True
                        Timer7.Enabled = False
End Sub



