VERSION 5.00
Begin VB.Form baterdattleweele 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   2130
   ClientTop       =   1275
   ClientWidth     =   11100
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "frmFlame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   740
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer6 
      Interval        =   10000
      Left            =   10560
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      Picture         =   "frmFlame.frx":27102
      ScaleHeight     =   2265
      ScaleWidth      =   11055
      TabIndex        =   12
      Top             =   -120
      Width           =   11085
   End
   Begin VB.FileListBox filWindows 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Hidden          =   -1  'True
      Left            =   1440
      System          =   -1  'True
      TabIndex        =   6
      Top             =   2895
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   240
   End
   Begin VB.ListBox lstPaths 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmFlame.frx":2C1F3
      Left            =   3120
      List            =   "frmFlame.frx":2C1F5
      TabIndex        =   5
      Top             =   2895
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer tmrBar 
      Interval        =   1
      Left            =   8520
      Top             =   240
   End
   Begin VB.DirListBox dirDirs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11055
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         Caption         =   "I'm the Baxterdattelweezel Virus."
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   2160
         Width           =   11055
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   0
         Top             =   3120
         Width           =   10995
      End
      Begin VB.Label lblDir 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   2880
         Width           =   7215
      End
      Begin VB.Shape shpProgress 
         BorderColor     =   &H00004080&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   3120
         Width           =   15
      End
      Begin VB.Label txtDelete 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   7215
      End
      Begin VB.Label txtLong2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4200
         Width           =   7215
      End
      Begin VB.Label txtLong 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   7215
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   35000
      Left            =   6960
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Interval        =   50
      Left            =   9720
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   28500
      Left            =   6480
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   25000
      Left            =   6000
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   5520
      Top             =   240
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   120
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   188
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.PictureBox picFlame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   -1080
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   1
      Top             =   960
      Width           =   9735
   End
   Begin VB.Timer tmrFlame 
      Interval        =   1
      Left            =   4560
      Top             =   240
   End
   Begin VB.PictureBox picPAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "frmFlame.frx":2C1F7
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   871
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   13065
   End
End
Attribute VB_Name = "baterdattleweele"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'Process A Beep Through The PC's Internal Speaker
'dwFreq Can Be 37 To 37767
'dwDuration Is Any Amount Of Time In Milliseconds: 1000ms = 1sec



Private ontop As New clsOnTop

Dim Console As New clsConsole
Dim lFreq As Long
Dim Plot(320, 200) As Integer, BlkSize As Integer, XOff As Long, YOff As Long, StopRandFlame As Boolean, DragBrightness As Single

Private Sub cmdMakeTopMost_Click()
   ontop.MakeNormal hWnd
End Sub

Private Sub Command2_Click()
CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "0"

End Sub

Private Sub Form_Load()
Set ontop = New clsOnTop
    hidemouse
    'make on top.
        ontop.MakeTopMost hWnd
 CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "1"
 
  

    Dim CurrentX As Integer, CurrentY As Integer
    picPAL.Width = 255 * 6
    picPAL.Height = 10
    
    'Try setting the "DragBrightness" parameter between 0 to 60
    DragBrightness = 1
        
    BlkSize = 5
    picFlame.Width = 200 * BlkSize + 3
    picFlame.Height = 100 * BlkSize - 2
  



    Call InitFlame
    StopRandFlame = True

     
      filWindows.ListIndex = 0
    Timer6.Enabled = False
  tmrBar.Enabled = False

  
Picture1.Visible = False

Picture1.Left = 200
 

Label1.Width = 9000
Label1.Visible = False
Label1.Left = 900

Frame1.Width = 1600
Shape1.Visible = False
End Sub






Private Sub InitFlame()
    'This subroutine initializes the flames array
    Dim XPos As Integer, YPos As Integer
    For YPos = 10 To 200
        For XPos = 0 To 320
            Plot(XPos, YPos) = 0
        Next XPos
    Next YPos
End Sub

Private Sub DrawFlame(StopGen As Boolean)
    Dim XPos As Integer, YPos As Integer, Sum As Integer
    Randomize Timer
    'This generate random patterns for flame effect
    For YPos = 100 To 98 Step -1
        For XPos = 10 To 180
            If (StopGen = True) Then Sum = Int(Rnd * 256) + 1 Else Sum = 1
            Plot(XPos, YPos) = Sum - 1
        Next XPos
    Next YPos
    'This plots the flame on the screen and adding blurry effect
    For YPos = 98 To 1 Step -1
        For XPos = 10 To 199
            Sum = Plot(XPos - 1, YPos + 1) + Plot(XPos, YPos + 1) + Plot(XPos + 1, YPos + 1) + Plot(XPos, YPos)
            Sum = Sum / 4
            Plot(XPos, YPos) = Sum + Int(Rnd * DragBrightness)
            If (YPos <= 97) Then BitBlt picFlame.hdc, XPos * BlkSize, YPos * BlkSize, BlkSize, BlkSize, picPAL.hdc, Plot(XPos, YPos) * 4, 1, SRCCOPY
        Next XPos
    Next YPos
    'This prints the caption texts
    If (StopGen = False) Then
        picFlame.FontSize = 20: picFlame.CurrentX = 130: picFlame.CurrentY = 20
        picFlame.Print "Thanks for viewing..."
    Else
        picFlame.FontSize = 48: picFlame.CurrentX = 75: picFlame.CurrentY = 420
        picFlame.Print "BAXTERDATTELWEEZEL VIRUS"
  
    End If
    picFlame.Refresh
End Sub




Private Sub Timer1_Timer()
  

Picture1.Visible = True
 lFreq = (10750 / 46) + 10
        Beep lFreq, 100
         lFreq = (5100 / 52) + 49
        Beep lFreq, 50
         lFreq = (21750 / 32) + 79
        Beep lFreq, 100
              lFreq = ((-64) * (7000 / 26)) + 37
        Beep lFreq, 100
         lFreq = (27000 / 82) + 115
        Beep lFreq, 150
              lFreq = ((-64) * (7000 / 26)) + 37
        Beep lFreq, 100
         lFreq = (37000 / 90) + 189
        Beep lFreq, 200
             lFreq = (22000 / 36) + 110
        Beep lFreq, 250
             lFreq = (30100 / 56) + 40
       Beep lFreq, 100
         lFreq = (15100 / 52) + 80
        Beep lFreq, 50
         lFreq = (21750 / 72) + 119

        Label1.Visible = True
        Timer1.Enabled = False
End Sub


Private Sub Timer2_Timer()
   ontop.MakeNormal hWnd
 
 lFreq = (10750 / 16) + 10
        Beep lFreq, 100
         lFreq = (11050 / 33) + 29
        Beep lFreq, 50
         lFreq = (31750 / 66) + 49
        Beep lFreq, 100
  Console.ConsoleWindowTitle = "-baxterdattelweezel.exe"
  
  Console.LoadConsole

  ShowWelcome
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
ontop.MakeTopMost hWnd

Label1.Caption = "I have infected your computer!"
     DoEvents
Console.CloseConsole
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Static Col1, Col2, Col3 As Integer
Static C1, C2, C3 As Integer
If (Col1 = 0 Or Col1 = 250) And (Col2 = 0 Or Col2 = 250) And (Col3 = 0 Or Col3 = 250) Then
C1 = Int(Rnd * 3)
C2 = Int(Rnd * 3)
C3 = Int(Rnd * 3)
End If
If C1 = 1 And Col1 <> 0 Then Col1 = Col1 - 10
If C2 = 1 And Col2 <> 0 Then Col2 = Col2 - 10
If C3 = 1 And Col3 <> 0 Then Col3 = Col3 - 10
If C1 = 2 And Col1 <> 250 Then Col1 = Col1 + 10
If C2 = 2 And Col2 <> 250 Then Col2 = Col2 + 10
If C3 = 2 And Col3 <> 250 Then Col3 = Col3 + 10
Label1.ForeColor = RGB(Col1, Col2, Col3)
End Sub





Private Sub Timer5_Timer()
 
 lFreq = (10750 / 16) + 10
        Beep lFreq, 100
         lFreq = (11050 / 33) + 29
        Beep lFreq, 50
         lFreq = (31750 / 66) + 49
        Beep lFreq, 100
 tmrBar.Enabled = True
Label1.Caption = "Initializing..."
Label1.Left = 1100
Timer5.Enabled = False

End Sub

Private Sub Timer6_Timer()
 'enable mouse,task manager, also not on top.
    
    showmouse
    ontop.MakeNormal hWnd
    CreateIntegerKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskmgr", "0"
         StopRandFlame = False
    Picture1.Visible = False
Frame1.Visible = False
 lFreq = (10750 / 46) + 10
        Beep lFreq, 100
         lFreq = (5100 / 52) + 49
        Beep lFreq, 50
         lFreq = (21750 / 32) + 79
        Beep lFreq, 100
              lFreq = ((-64) * (7000 / 26)) + 37
        Beep lFreq, 100
         lFreq = (27000 / 82) + 115
        Beep lFreq, 150
              lFreq = ((-64) * (7000 / 26)) + 37
        Beep lFreq, 100
         lFreq = (37000 / 90) + 189
        Beep lFreq, 200
             lFreq = (22000 / 36) + 110
        Beep lFreq, 250
             lFreq = (30100 / 56) + 40
       Beep lFreq, 100
         lFreq = (15100 / 52) + 80
        Beep lFreq, 50
         lFreq = (21750 / 72) + 119
MsgBox "                         Just kidding" & vbCrLf & vbCrLf & _
       "The baxterdattleweezle virus was only a joke." & vbCrLf & _
       "     Your computer will now return to normal." & vbCrLf & _
        "           Thank you, and have a nice day.", vbOKOnly
End
End Sub

Private Sub tmrBar_Timer()

    Dim Temp As String
    Open App.Path & "\Dirs.txt" For Input As 1
        Do Until EOF(1)
            Line Input #1, Temp
            lstPaths.AddItem Temp
        Loop
    Close #1
    lstPaths.AddItem "end"
    If shpProgress.Width < Shape1.Width Then
        DoEvents
        shpProgress.Width = shpProgress.Width + 100
        DoEvents
    End If
    
    If shpProgress.Width >= Shape1.Width Then
        tmrBar.Enabled = False
        tmrRun.Enabled = True
         shpProgress.Width = 15
        lstPaths.ListIndex = lstPaths.ListIndex + 1
        shpProgress.Width = 15
        Shape1.Width = filWindows.ListCount * 10

    End If
 
End Sub

Private Sub tmrRun_Timer()

    If filWindows.ListIndex = filWindows.ListCount - 1 Then 'While there are files left to do
        lstPaths.ListIndex = lstPaths.ListIndex + 1
        If Not lstPaths.Text = "end" Then 'Start on a new directory
            On Error Resume Next
            filWindows.Path = lstPaths.Text
            shpProgress.Width = 15
            Shape1.Width = filWindows.ListCount * 10
  Else
   
          Label1.Caption = "Your computer has been erased."

        
           tmrRun.Enabled = False
           Timer6.Enabled = True
         
        End If
    Else
        filWindows.ListIndex = filWindows.ListIndex + 1 'Go to next file, print its name and update the progress bar
        DoEvents
        shpProgress.Width = shpProgress.Width + 10
        lblDir.Caption = "Deleting contents in:  " & lstPaths.Text & "... "
       Label1.Caption = "Deleting: " & filWindows.FileName
       Label1.Left = 300
       DoEvents
        txtDelete.Caption = "Verifying file integrity...  " & shpProgress.Width / Shape1.Width * 100 & "%"
    Dim K As Long
        Randomize
        K = Int(Rnd * 1000000) + 1
        txtLong.Caption = "Sizing to zero...  " & filWindows.FileName & " \%cw " & K
        txtLong2.Caption = "Formatting...  " & lstPaths.Text & "\" & filWindows.FileName

    End If
    
End Sub


Private Sub tmrFlame_Timer()
    'This is the animation timer
    Static Count As Integer
    If (StopRandFlame = False) Then Count = Count + 1
    Call DrawFlame(StopRandFlame)
  
End Sub
Sub ShowWelcome()

' created using http://st-www.cs.uiuc.edu/users/chai/figlet.html

Console.WriteOut String(79, "=")
Console.WriteOut String(79, "=")
Console.WriteOut String(79, "=")
Console.WriteOut " Installing baxterdattelweezel virus."
Console.WriteOut " Installing..."
Console.WriteOut " C:\"
Console.WriteOut " C:\Windows\"
Console.WriteOut " C:\Windows\System\"
Console.WriteOut " C:\Windows\System32\"
Console.WriteOut " C:\Windows\System32\autoexec.bat"
Console.WriteOut " C:\Windows\System32\system32.exe"
Console.WriteOut " C:\Windows\System32\kernal32.exe"
Console.WriteOut " baxterdattelweezel.exe Completed!"
Console.WriteOut String(79, "=")
Console.Important " Infecting..."
Console.WriteOut String(79, "=")


End Sub

