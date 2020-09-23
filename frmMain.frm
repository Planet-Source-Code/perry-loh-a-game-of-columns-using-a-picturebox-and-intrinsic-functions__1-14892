VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VeeBee Columns"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3645
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2745
      Top             =   3060
   End
   Begin VB.Frame fraMain 
      Height          =   3525
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   3345
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         Height          =   3075
         Left            =   135
         ScaleHeight     =   3015
         ScaleWidth      =   1635
         TabIndex        =   5
         Top             =   270
         Width           =   1695
      End
      Begin VB.Frame fraScore 
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2025
         TabIndex        =   3
         Top             =   180
         Width           =   1095
         Begin VB.Label lblScore 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   4
            Top             =   270
            Width           =   780
         End
      End
      Begin VB.Frame fraLevel 
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   1
         Top             =   900
         Width           =   735
         Begin VB.Label lblLevel 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   90
            TabIndex        =   2
            Top             =   315
            Width           =   510
         End
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A clone of COLUMNS - Keeping It Short and Simple
' Perry Loh (skeevs@hotmail.com,skeevs@crosswinds.net)
' Completed version 1.0 on 01/31/2001
'
' Started working on this about 6 months ago, and left it on my computer till the last
' few days where I actually got to solve some major/minor bugs and actually get it to a playable
' state. This is the first "game" that I've worked on. Decided to use intrisic functions
' to keep it simple. Creating this created more logical and code issues rather than
' technical issues such as using bitblt/directx. Which is good for a beginner.. :)
'
' Anyhow, it was good to learn from and good to complete some code that was lying around
' unfinished for the last 5 months or so.
'
' I'm still tinkering around with the code and adding new functions to add some new things
' into the gameplay. Right now, it is still a very basic COLUMNS clone.
'
' I can be reached on MSN Messenger, and my ID is skeevs@hotmail.com

Dim PlayField(FIELD_WIDTH, FIELD_HEIGHT) As udtFieldBlock
Dim SysTime As SYSTEMTIME
Dim bolRunning As Boolean

Dim UserScore As Long
Dim UserLevel As Integer
Dim CurrColumn(1) As udtBlock
Dim CurrColumnDir As Integer
Dim FirstColumnColor As Integer

Private Sub StartGame()
    ' Inits all properties to start the game.
    ' Also creates a new column and puts it on the playfield
    Dim i As Integer, j As Integer
    
    bolRunning = True
    UserScore = 0
    UserLevel = 1
    lblScore = 0
    
    ' Reset play field
    For i = 1 To FIELD_HEIGHT
        For j = 1 To FIELD_WIDTH
            PlayField(j, i).Color = 0
            PlayField(j, i).Block = False
        Next j
    Next i
    
    'Clear pic
    picMain.Cls
    
    Timer1.Interval = TIMER_GAME
    Timer1.Enabled = True
    FirstColumnColor = GetColor()
    Call CreateColumn
    Call PutColumn
End Sub

Private Sub GameOver()
    ' Ends the game and shows a message and score to the user
    bolRunning = False
    Timer1.Enabled = False
    MsgBox "Game Over!" & vbCrLf & vbCrLf & "Your Score : " & UserScore, vbExclamation
End Sub

Private Sub ChangeLevel()
    ' Changes the level of the game, thereby shortening the timer interval to make it faster
    UserLevel = UserLevel + 1
    lblLevel = UserLevel
    Timer1.Interval = Timer1.Interval - 90
    Call PlaySnd(enmLevelUp)
End Sub

Private Sub DrawBlock(ByVal X As Integer, ByVal Y As Integer, ByVal BFillColor As Integer, Optional ByVal OutLineColor As Integer = -1)
    ' Draws a block to the picturebox based on coordinates
    picMain.Line (X * BLOCK_SIZE, Y * 200)-((X * BLOCK_SIZE) - BLOCK_SIZE, Y * 200 - BLOCK_SIZE), QBColor(BFillColor), BF
    If OutLineColor <> -1 Then
        picMain.Line (X * BLOCK_SIZE, Y * 200)-((X * BLOCK_SIZE) - BLOCK_SIZE, Y * 200 - BLOCK_SIZE), QBColor(OutLineColor), B
    End If
    picMain.Refresh
End Sub

Private Sub UnDrawBlock(ByVal X As Integer, ByVal Y As Integer)
    ' Undraws a block to the picturebox based on coordinates
    picMain.Line (X * BLOCK_SIZE, Y * 200)-((X * BLOCK_SIZE) - BLOCK_SIZE, Y * 200 - BLOCK_SIZE), picMain.BackColor, BF
    picMain.Refresh
End Sub

Private Sub CreateColumn()
    ' Creates a new random column. We take from the FirstColumnColor variables so that
    ' we get a different or rather more random color generated by the random functions
    CurrColumn(0).Color = FirstColumnColor
    CurrColumn(1).Color = GetColor()
    
    
    ' Sets the coordinates for the column to start on
    CurrColumn(0).X = 4
    CurrColumn(0).Y = 1
    CurrColumn(1).X = 5
    CurrColumn(1).Y = 1
    
    ' This column is falling
    CurrColumn(0).Falling = True
    CurrColumn(1).Falling = True
    
    ' The direction it's facing is to the right
    CurrColumnDir = DIR_RIGHT
    
    ' Tag the field to show that this block currently resides there
    PlayField(CurrColumn(0).X, CurrColumn(0).Y).Color = CurrColumn(0).Color
    PlayField(CurrColumn(1).X, CurrColumn(1).Y).Color = CurrColumn(1).Color
    
End Sub

Private Sub PutColumn()
    ' Puts the column there, this is more like a refresh function
    Call DrawBlock(CurrColumn(0).X, CurrColumn(0).Y, CurrColumn(0).Color, BLACK)
    Call DrawBlock(CurrColumn(1).X, CurrColumn(1).Y, CurrColumn(1).Color, BLACK)
End Sub

Private Sub MoveColumn(MoveX As Integer, MoveY As Integer)
    ' Moves column based on the parameters sent in, and then checks for collision with
    ' the IsCanMove function
    
    If IsCanMove(CurrColumn(0).X + MoveX, CurrColumn(0).Y + MoveY, CurrColumn(1).X + MoveX, CurrColumn(1).Y + MoveY) = CLD_NONE Then
       
                ' Clear the current position in the field because we are moving
                Call ClearBlock(CurrColumn(0).X, CurrColumn(0).Y)
                Call ClearBlock(CurrColumn(1).X, CurrColumn(1).Y)
                
                ' Assign the Current Column properties with the new position
                CurrColumn(0).X = CurrColumn(0).X + MoveX
                CurrColumn(0).Y = CurrColumn(0).Y + MoveY
                CurrColumn(1).X = CurrColumn(1).X + MoveX
                CurrColumn(1).Y = CurrColumn(1).Y + MoveY
                
                ' Tag the field to show that this block currently resides there
                PlayField(CurrColumn(0).X, CurrColumn(0).Y).Color = CurrColumn(0).Color
                PlayField(CurrColumn(1).X, CurrColumn(1).Y).Color = CurrColumn(1).Color
                
                ' Put the column on where the position properties are
                Call PutColumn
                Exit Sub
    'ElseIf IsCanMove(CurrColumn(0).X + MoveX, CurrColumn(0).Y + MoveY, CurrColumn(1).X + MoveX, CurrColumn(1).Y + MoveY) = CLD_BLOCK Then
    ' 30/01/00  Commented line above, because when we check the ".X + MoveX", it returns a block collision
    '           even tho the block can still fall down
    
    ElseIf IsCanMove(CurrColumn(0).X, CurrColumn(0).Y + MoveY, CurrColumn(1).X, CurrColumn(1).Y + MoveY) = CLD_BLOCK Then
            CurrColumn(0).Falling = False
            CurrColumn(1).Falling = False
            PlayField(CurrColumn(0).X, CurrColumn(0).Y).Block = True
            PlayField(CurrColumn(1).X, CurrColumn(1).Y).Block = True
            PlayField(CurrColumn(0).X, CurrColumn(0).Y).Color = CurrColumn(0).Color
            PlayField(CurrColumn(1).X, CurrColumn(1).Y).Color = CurrColumn(1).Color
            
            Exit Sub
            
    ElseIf IsCanMove(CurrColumn(0).X + MoveX, CurrColumn(0).Y + MoveY, CurrColumn(1).X + MoveX, CurrColumn(1).Y + MoveY) = CLD_WALL Then
       
            ' Clear the current position in the field because we are moving
            Call ClearBlock(CurrColumn(0).X, CurrColumn(0).Y)
            Call ClearBlock(CurrColumn(1).X, CurrColumn(1).Y)
            
            ' Assign the Current Column properties with the new position
            CurrColumn(0).Y = CurrColumn(0).Y + MoveY
            CurrColumn(1).Y = CurrColumn(1).Y + MoveY
            
            ' Tag the field to show that this block currently resides there
            PlayField(CurrColumn(0).X, CurrColumn(0).Y).Color = CurrColumn(0).Color
            PlayField(CurrColumn(1).X, CurrColumn(1).Y).Color = CurrColumn(1).Color
            
            ' Put the column on where the position properties are
            Call PutColumn
            Exit Sub
            
    ElseIf IsCanMove(CurrColumn(0).X + MoveX, CurrColumn(0).Y + MoveY, CurrColumn(1).X + MoveX, CurrColumn(1).Y + MoveY) = CLD_FLOOR Then
            CurrColumn(0).Falling = False
            CurrColumn(1).Falling = False
            PlayField(CurrColumn(0).X, CurrColumn(0).Y).Block = True
            PlayField(CurrColumn(1).X, CurrColumn(1).Y).Block = True
            PlayField(CurrColumn(0).X, CurrColumn(0).Y).Color = CurrColumn(0).Color
            PlayField(CurrColumn(1).X, CurrColumn(1).Y).Color = CurrColumn(1).Color
            
            Exit Sub
            
    End If
    DoEvents
End Sub

Private Sub MoveBlock(X As Integer, Y As Integer)
    ' The second block is the one that always rotates around the first block
    ' So whatever it is, the first block maintains its positions whereas the second
    ' block moves around the first block
    
    If IsCanMove(CurrColumn(0).X, CurrColumn(0).Y, CurrColumn(1).X + X, CurrColumn(1).Y + Y) = CLD_WALL Then
        Exit Sub
    ElseIf IsCanMove(CurrColumn(0).X + X, CurrColumn(0).Y + Y, CurrColumn(1).X + X, CurrColumn(1).Y + Y) = CLD_FLOOR Then
        Exit Sub
    ElseIf IsCanMove(CurrColumn(0).X + X, CurrColumn(0).Y + Y, CurrColumn(1).X + X, CurrColumn(1).Y + Y) = CLD_BLOCK Then
        Exit Sub
    End If
    
    If IsCanMove(CurrColumn(0).X, CurrColumn(0).Y, CurrColumn(1).X, CurrColumn(1).Y) = CLD_NONE Then
    
        ' If we can spin the block, then update the columns directions first
        Select Case CurrColumnDir
            Case DIR_RIGHT
                CurrColumnDir = DIR_UP
                
            Case DIR_UP
                CurrColumnDir = DIR_LEFT
                
            Case DIR_LEFT
                CurrColumnDir = DIR_DOWN
            
            Case DIR_DOWN
                CurrColumnDir = DIR_RIGHT
        End Select
        
        ' Play the Spin Block Sound
        Call PlaySnd(enmSpin)
        
        ' Clear the current position in the field because we are moving
        Call ClearBlock(CurrColumn(0).X, CurrColumn(0).Y)
        Call ClearBlock(CurrColumn(1).X, CurrColumn(1).Y)
        
        ' Clear position of field b4 updating new position, because we are rotating
        PlayField(CurrColumn(1).X, CurrColumn(1).Y).Block = False
                        
        ' Set the new position of the block to the X and Y properties
        CurrColumn(1).X = CurrColumn(1).X + X
        CurrColumn(1).Y = CurrColumn(1).Y + Y
        
        Call PutColumn
    End If
    
     
    picMain.Refresh
End Sub

Private Sub ClearBlock(PostX As Integer, PostY As Integer)
    ' Tag this position empty in the play field
    PlayField(PostX, PostY).Color = 0
    PlayField(PostX, PostY).Block = False
    
    ' Clear the picturebox of the graphic
    Call UnDrawBlock(PostX, PostY)
End Sub

Private Function IsCanMove(PostX1 As Integer, PostY1 As Integer, PostX2 As Integer, PostY2 As Integer) As Integer
    ' This is the main collision detection function, checks whether it collided with
    ' the wall , floor or another block
    
    ' Check whether it is hitting the side borders, left and right
    If PostX1 > FIELD_WIDTH Or PostX1 < 1 Or _
       PostX2 > FIELD_WIDTH Or PostX2 < 1 Then
        IsCanMove = CLD_WALL
        Exit Function
    End If
    
    ' Check whether it is hitting the floor
    If PostY1 > FIELD_HEIGHT Or PostY1 < 1 Or _
       PostY2 > FIELD_HEIGHT Or PostY2 < 1 Then
        IsCanMove = CLD_FLOOR
        Exit Function
    End If
    
    ' Check whether it is hitting a block
    If (PlayField(PostX1, PostY1).Block Or PlayField(PostX2, PostY2).Block) Then
        IsCanMove = CLD_BLOCK
        
    Else
        IsCanMove = CLD_NONE
        
    End If

End Function

Private Sub SpinColumn()
    ' Spin the block around the column, this rotates the blocks based on the direction.
    ' Note that rotating causes the 2nd block to rotate around the 1st block
    Select Case CurrColumnDir
        Case DIR_RIGHT
            Call MoveBlock(-1, -1)
            
        Case DIR_UP
            Call MoveBlock(-1, 1)
            
        Case DIR_LEFT
            Call MoveBlock(1, 1)
            
        Case DIR_DOWN
            Call MoveBlock(1, -1)
            
    End Select
End Sub

Private Sub ChkConnect()
    ' Function chks to see if there are columns that are connected
    ' If there are, get rid of it from the screen and update score
    
    Dim i As Integer, j As Integer, CurrentColor As Integer, HorzCounter As Integer, VertCounter As Integer
    Dim FlashCount As Integer, HorzCounter2 As Integer, VertCounter2 As Integer
    
    For i = 1 To FIELD_HEIGHT
        For j = 1 To FIELD_WIDTH
            ' Get the color of the current block first
            CurrentColor = PlayField(j, i).Color
            
            ' Set the connect counters to 1
            HorzCounter = 1
            VertCounter = 1
            
            ' Check horizontal connections
            If (j + 1) < FIELD_WIDTH Then
                If PlayField(j, i).Block And PlayField(j + 1, i).Block And PlayField(j + 2, i).Block Then
                    
                    ' Checks how many blocks of the same color there are in a row,
                    ' and stores them in the HorzCounter variable
                    Do Until PlayField(j + HorzCounter, i).Color <> CurrentColor
                            HorzCounter = HorzCounter + 1
                            ' If the next addition is greater than field width,
                            ' EXIT so we don't get a Out of range error for the array
                            If (j + HorzCounter) > FIELD_WIDTH Then
                                Exit Do
                            End If
                    Loop
    
                End If
            End If
            
            ' Check vertical connections
            If (i + 1) < FIELD_HEIGHT Then
                If PlayField(j, i).Block And PlayField(j, i + 1).Block And PlayField(j, i + 2).Block Then
                    ' Checks how many blocks of the same color there are in a row,
                    ' and stores them in the HorzCounter variable
                    Do Until PlayField(j, i + VertCounter).Color <> CurrentColor
                            VertCounter = VertCounter + 1
                            ' If the next addition is greater than field width,
                            ' EXIT so we don't get a Out of range error for the array
                            If (i + VertCounter) > FIELD_HEIGHT Then
                                Exit Do
                            End If
                    Loop
                End If
            End If
            
            ' If there are more than 3 blocks of the same color in a row
            ' Flash the blocks and remove the blocks from the play field
            If HorzCounter >= 3 Then
                For HorzCounter2 = 0 To (HorzCounter - 1)
                    Call DrawBlock(j + HorzCounter2, i, 15)
                Next HorzCounter2
                
                Sleep FLASH_TIME
                
                ' Play the Clear blocks sound
                Call PlaySnd(enmClear)
                
                For HorzCounter2 = 0 To (HorzCounter - 1)
                    Call DrawBlock(j + HorzCounter2, i, PlayField(j + HorzCounter2, i).Color, BLACK)
                Next HorzCounter2
                
                Sleep FLASH_TIME
                
                ' Play the Clear blocks sound
                Call PlaySnd(enmClear)
                
                For HorzCounter2 = 0 To (HorzCounter - 1)
                    Call ClearBlock(j + HorzCounter2, i)
                Next HorzCounter2
                
                ' Update the users score
                UserScore = UserScore + (HorzCounter * 50)
                lblScore = UserScore
                
                
            End If
            
            ' Do the same for vertical rows
            If VertCounter >= 3 Then
                For VertCounter2 = 0 To (VertCounter - 1)
                    Call DrawBlock(j, i + VertCounter2, 15)
                Next VertCounter2
                
                Sleep FLASH_TIME
                
                ' Play the Clear blocks sound
                Call PlaySnd(enmClear)
                
                For VertCounter2 = 0 To (VertCounter - 1)
                    Call DrawBlock(j, i + VertCounter2, PlayField(j, i + VertCounter2).Color, BLACK)
                Next VertCounter2
                
                Sleep FLASH_TIME
                
                ' Play the Clear blocks sound
                Call PlaySnd(enmClear)
                
                For VertCounter2 = 0 To (VertCounter - 1)
                    Call ClearBlock(j, i + VertCounter2)
                Next VertCounter2
                
                ' Update the users score
                UserScore = UserScore + (VertCounter * 50)
                lblScore = UserScore
                                
                       
            End If
                 
        Next j
    Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Select Case KeyCode
        Case vbKeyLeft
            Call MoveColumn(-1, 0)
            
        Case vbKeyRight
            Call MoveColumn(1, 0)
            
        Case vbKeyDown
            Call MoveColumn(0, 1)
        
        Case vbKeySpace
            Call SpinColumn
        Case vbKeyShift
            If Timer1.Enabled = True Then Timer1.Enabled = False Else Timer1.Enabled = True
    End Select
    DoEvents
    Call RedrawField
End Sub


Private Sub ChkAddColumn()
    ' If the blocks are not falling anymore, then create a new column and put it on
    ' the play field
    
    If Not (CurrColumn(0).Falling And CurrColumn(1).Falling) Then
        If Not ChkFieldFull Then
            Call CreateColumn
            Call PutColumn
        End If
    End If
End Sub

Private Function ChkFieldFull() As Boolean
    ' Checks to see if the blocks are all on the top, if it is, then it's game over!
    If PlayField(4, 1).Block Or PlayField(5, 1).Block Then
        Call GameOver
        ChkFieldFull = True
    Else
        ChkFieldFull = False
    End If
End Function

Private Sub ChkScore()
    ' Checks the score and changes the level
    If UserScore >= 2500 And UserLevel = 1 Then
        Call ChangeLevel
    ElseIf UserScore >= 5000 And UserLevel = 2 Then
        Call ChangeLevel
    ElseIf UserScore >= 7500 And UserLevel = 3 Then
        Call ChangeLevel
    ElseIf UserScore >= 10000 And UserLevel = 4 Then
        Call ChangeLevel
    End If
End Sub

Private Sub RedrawField()
    ' This sub loops through the PlayField array and calls the draw function to
    ' draw all the blocks residing in the playfield
    
    Dim i As Integer, j As Integer
    
    For i = 1 To FIELD_HEIGHT
        For j = 1 To FIELD_WIDTH
            If PlayField(j, i).Block Then
                Call DrawBlock(j, i, PlayField(j, i).Color, BLACK)
                DoEvents
            End If
        Next j
    Next i
End Sub

Private Function GenRandom() As Integer
    ' Generate a random number based on the system time
    Call GetSystemTime(SysTime)
    Randomize (SysTime.wMilliseconds)
    GenRandom = (3 * Rnd()) + 1
End Function

Private Function GetColor() As Integer
    ' Selects a color based on the random number
    Select Case GenRandom()
        Case 1
            GetColor = MAGENTA
        Case 2
            GetColor = LIGHT_GREEN
        Case 3
            GetColor = LIGHT_RED
        Case 4
            GetColor = LIGHT_YELLOW
    End Select

End Function

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuReset_Click()
    Call StartGame
End Sub

Private Sub mnuStart_Click()
    Call StartGame
End Sub
Private Sub PlaySnd(Soundtype As enmSoundType)
    ' Plays the selected sound
    Select Case Soundtype
        Case enmSpin
            Call PlaySound(App.Path + "\Sounds\spin.wav", 0, SND_ASYNC)
        Case enmClear
            Call PlaySound(App.Path + "\Sounds\clear.wav", 0, SND_ASYNC)
        Case enmLevelUp
            Call PlaySound(App.Path + "\Sounds\levelup.wav", 0, SND_ASYNC)
    End Select
End Sub

Private Sub Timer1_Timer()
    ' Main Game loop
    
    FirstColumnColor = GetColor
    Call MoveColumn(0, 1)
    Call ChkConnect
    Call ChkScore
    Call ChkAddColumn
    Call RedrawField
    
    DoEvents
    picMain.Refresh
End Sub
