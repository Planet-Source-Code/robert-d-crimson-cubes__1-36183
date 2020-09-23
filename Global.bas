Attribute VB_Name = "Global"
Option Explicit
'Option Explicit will save you more pain than it creates!

'This routine holds the definintions for the global variables and has a few misc general game play routines
'(I like to use the "g" to designate global variables but that's about as far as I go with Hungarian notation.)
Global gDiceRoll As Integer 'Total of the two dice for the current roll
Global gGamePoint As Integer 'Value of the point for the current game
Global gLastBet As Integer 'Holds the value of the last bet to be placed.  Used for Undo.

Global gComeOutRoll As Boolean 'Come out roll flag - set to true when it is the come out roll
Global gRolledHard As Boolean 'Set to true if 4,6,8 or 10 was rolled the hard way

'This array holds the amounts of the bets
Global gPlacedBet(49) As Integer

'This list is all the possible bets that can be made
'Enumerated types can make code much more readable.
'For example, instead of the following:
'   Select case Var%
'       Case 0
'       Case 1
'       Case 2
'         .
'         .
'         .
' You can use:
'   Select case Var%
'       Case Bet.IsNotValid
'       Case Bet.IsPass
'       Case Bet.IsPassOdds
'         .
'         .
'         .
'The first example is pretty meaningless.  You could add comments to explain what the case values represent
'but, as I hope you can see, the second example is self documenting.  You don't need to read anything more than
'then code to know what bet each case represents.

'Because of the way some of the tests are structered, groups of bets need to be in sequential order but
'for the most part, the actual numbering is just how I happened to write them down as I was recalling
'what all the bets are.
Enum Bet
    IsNotValid = 0
    IsPass = 1
    IsPassOdds = 2
    'Come Numbers have to be sequential
    IsCome4 = 3
    IsCome5 = 4
    IsCome6 = 5
    IsCome8 = 6
    IsCome9 = 7
    IsCome10 = 8
    IsCome4Odds = 9
    IsCome5Odds = 10
    IsCome6Odds = 11
    IsCome8Odds = 12
    IsCome9Odds = 13
    IsCome10Odds = 14
    '(The Come bet needs to be assigned a higher number in the list than the actual Come Numbers so it will get tested
    ' for winning or losing last.  That way any Number it may get reassigned to will already be cleared.)
    IsCome = 15
    'Dont Numbers have to be sequential
    IsDont4 = 16
    IsDont5 = 17
    IsDont6 = 18
    IsDont8 = 19
    IsDont9 = 20
    IsDont10 = 21
    IsDont4Odds = 22
    IsDont5Odds = 23
    IsDont6Odds = 24
    IsDont8Odds = 25
    IsDont9Odds = 26
    IsDont10Odds = 27
    '(Dont is similar to Come - See note above)
    IsDont = 28
    IsBig6 = 29
    IsBig8 = 30
    IsField = 31
    'Place bets need to be sequential
    IsPlace4 = 32
    IsPlace5 = 33
    IsPlace6 = 34
    IsPlace8 = 35
    IsPlace9 = 36
    IsPlace10 = 37
    IsHard = 38
    'The Hardways have to be sequential
    IsHard4 = 39
    IsHard6 = 40
    IsHard8 = 41
    IsHard10 = 42
    IsAny7 = 43
    IsCraps = 44
    IsHorn = 45
    'The Horn bets have to be sequenctial
    IsHorn2 = 46
    IsHorn3 = 47
    IsHorn11 = 48
    IsHorn12 = 49
End Enum

Public Sub Main()
    'The first roll of the game is always a come out roll so set the come out roll flag
    gComeOutRoll = True
    'Seed the random number generator
    Randomize
    'Start the game
    Layout.Show
End Sub

'The next two routines are used to create the illusion of the custom buttons getting pressed and released.
'In the mousedown event of a button image control, Press_Button gets called and in the mouseup event,
'Release_Button is called.  The advantage to doing it this way is that the code makes a little more sense
'(the mousedown event calls Press_Button which is much more descriptive of what is going on visually
'than setting the image control's visible property to false) and if I decide to add a sound effect I would
'only need to add it here instead of every mousedown event for every image control.

Public Sub Press_Button(ByRef ButtonX As Image)
    'Turn off the button image to create the illusion it has been pressed in
    ButtonX.Visible = False
End Sub

Public Sub Release_Button(ByRef ButtonX As Image)
    'Turn on the button image to create the illusion that it has been released
    ButtonX.Visible = True
End Sub

Public Function Roll_The_Dice() As Integer
    'This function rolls the dice, returns the total and if the number was rolled the hard way,
    'as a pair, then it sets the "Rolled Hard" flag to true
    
    '(Two dice need to be rolled so pairs, "hard ways", can be tested for)
    Dim Die_1 As Integer
    Dim Die_2 As Integer
    
    'Roll the dice
    Die_1% = Inc(Int(Rnd * 6))
    Die_2% = Inc(Int(Rnd * 6))
        
    'See if this number was rolled the hard way
    '(There is no such a thing as a hard 2 or hard 12 bet but setting the flag doesn't
    ' cause any problems so there is no need to have a test to exclude them)
    If Die_1% = Die_2% Then
        gRolledHard = True
    Else
        gRolledHard = False
    End If
    
    'Reference the layout form and set the images for the dice
    With Layout
        .img_D1.Picture = .img_Dice.ListImages(Die_1%).Picture
        .img_D2.Picture = .img_Dice.ListImages(Die_2%).Picture
    End With
    
    'Once the dice are rolled the player can no longer recall the last bet
    gLastBet% = 0
    
    'Only need to return the total because the "Rolled Hard" flag will take care of
    'hard way testing when the time comes.
    Roll_The_Dice = Die_1% + Die_2%
    
End Function

Public Sub Delay(ByVal TimeDelay As Single)
    'This is just a loop that will cause the program to pause for however long a time is passed in.
    '(In this game, it is used as a delay between messages)
    Dim BeginDelay As Single
    'Get the current timer value
    BeginDelay! = Timer
    'Loop until the timer value exceeds that of the start time plus the delay time
    Do While Timer < BeginDelay! + TimeDelay!
    Loop
End Sub

Public Sub Button_Control()
    'This routine controls the button image that marks the point
    'See if a new point needs to be marked
    If gComeOutRoll Then
        'We are on a come out roll so determine if a 4,5,6,8,9 or 10 has been rolled
        If gDiceRoll% > 3 And gDiceRoll% < 11 And gDiceRoll% <> 7 Then
            'Place the button image on the number and establish the "point" value
            gGamePoint% = gDiceRoll%
            'Reference the layout form
            With Layout
                '4 is on the left edge of the layout.  This calculation will determine how many units to the right of
                'the 4 to place the button.  To place the button on the 4, for example, the button needs to be placed
                'zero units to the right.  The 5 is one unit to the right of the 4 and 6 is two units.  8 is three units but
                'in value it is, in essence, 4 units larger.  That is where the test to see if the roll is larger than 7 comes
                'in to play.  The number 4, 5 and 6 are less than 7 so the test comes back false which has a value of
                'zero and the numbers 8, 9 and 10 are true which has a value (in VB anyway) of -1.  So when the
                '8,9 or 10 is rolled, one unit is effectively subtracted from the units to move and button is placed on
                'the correct number.
                .img_Button.Move (450 + (gDiceRoll% - 4 + (gDiceRoll% > 7)) * 1095), 2670
                .img_Button.Visible = True
                'Disable the control so it won't block mouse moves on the form
                .img_Button.Enabled = False
            End With
            'A point has been established so this is no longer a come out roll so reset the flag
            gComeOutRoll = False
        End If
    Else
        'See if the button needs to be removed
        If gDiceRoll% = 7 Or gDiceRoll% = gGamePoint% Then
            'A 7 was rolled or the point was made so remove the button image
            Layout.img_Button.Visible = False
            'We are now on a come out roll so set the flag
            gComeOutRoll = True
        End If
    End If
End Sub

Public Function Inc(ByVal ThisNumber As Integer) As Integer
    'This function return increments "ThisNumber%" by 1 and returns the result.
    '(It seems that I always end up incrementing a counter at some point so I stick
    ' this function in somewhere so instead of "var% = var% + 1" I can just
    ' just type "Inc var%")
    Inc = ThisNumber% + 1
End Function

Public Sub End_The_Game()
    'Inform the player of the results
    If Val(Layout.lbl_BankRoll.Caption) > 250 Then
        MsgBox "Thanks for Playing Crimson Cubes." & Chr$(13) & "You won $" & CStr(Val(Layout.lbl_BankRoll.Caption) - 250), , "We Have A Winner!"
    Else
        If Val(Layout.lbl_BankRoll.Caption) < 250 Then
            MsgBox "Thanks for Playing Crimson Cubes." & Chr$(13) & "You lost $" & CStr(250 - Val(Layout.lbl_BankRoll.Caption)), , "Better Luck Next Time..."
        Else
             MsgBox "Thanks for Playing Crimson Cubes." & Chr$(13) & "Looks like you broke even.", , "Better Than Losing...."
        End If
    End If
    'Release the resources and end the program
    Set Layout = Nothing
    Set Main_Help = Nothing
    End
End Sub
