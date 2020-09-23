Attribute VB_Name = "Make_Bets"
Option Explicit

'This module takes care of placing a bet.
'(Generating the graphics for the bets, determining which bet the player wants to make, validating bets etc.)

'These are the starting top and left locations for the current bet image - used to replace the bet if the player
'tries to make an illegal bet.
Private Const Home_X% = 10020
Private Const Home_Y% = 4920

Public Function Find_The_Number(ByVal X_Coord As Integer) As Integer
    'This function will determine which number (4,5,6,8,9 or 10) the player has selected based on the horizontal value (X coordinate) passed in.
    'Start by determining which half we are dealing with - 4,5,6 or 8,9,10.
    'Then start at the outer most position (4 or 10) and work back toward the center (6 or 8) until the number has been found.
    '(used to determine which number for Come, Don't Come or Place bets)
    If X_Coord% < 3705 Then
        '4,5 or 6
        If X_Coord% < 1500 Then
            Find_The_Number = 4
        Else
            If X_Coord% < 2610 Then
                Find_The_Number = 5
            Else
                Find_The_Number = 6
            End If
        End If
    Else
        '8,9 or 10
        If X_Coord% > 5895 Then
            Find_The_Number = 10
        Else
            If X_Coord% > 4800 Then
                Find_The_Number = 9
            Else
                Find_The_Number = 8
            End If
        End If
    End If
End Function

Public Function Find_The_Mouse(ByVal X_Coord As Integer, ByVal Y_Coord As Integer) As Integer
    'This function determines which bet the mouse was clicked on (or if off the playing field) by dividing the field into sections and testing against
    'the coordinates (mouse X and Y values) passed.  Works in the same fashion as Find_The_Number but takes "Y" into account and returns
    'the number of the bet that was clicked in.  (Ex: Clicking on 4,5,6,8,9 or 10 returns a value of 13 for Place bet.)
    Dim NumberBet As Integer
    
    'Test for Center bets or not
    If X_Coord% > 6975 Then
        'Center bet section or off the field
        If X_Coord% > 9225 Or Y_Coord% > 8025 Or Y_Coord% < 2745 Then
            'Mouse is off the playing field so bet is not valid
            Find_The_Mouse = Bet.IsNotValid
            'Nothing else to do so exit the function
            Exit Function
        End If
        'Determine which Center bet
        If Y_Coord% < 5355 Then
            'Is either Hardway or Any 7
            If Y_Coord% > 4740 Then
                Find_The_Mouse = Bet.IsAny7
            Else
                'Hardway - may be individual number or covering all
                If Y_Coord% < 3525 Then
                    Find_The_Mouse = Bet.IsHard
                Else
                    'Determine if 4,6, 8 or 10
                    If Y_Coord% < 4095 Then
                        'Is either 6 or 8
                        If X_Coord% < 8145 Then
                            Find_The_Mouse = Bet.IsHard6
                        Else
                            Find_The_Mouse = Bet.IsHard8
                        End If
                    Else
                        'Is either 4 or 10
                        If X_Coord% < 8145 Then
                            Find_The_Mouse = Bet.IsHard4
                        Else
                            Find_The_Mouse = Bet.IsHard10
                        End If
                    End If
                End If
            End If
        Else
            'Is either Craps or Horn
            If Y_Coord% < 5940 Then
                Find_The_Mouse = Bet.IsCraps
            Else
                'Horn - may be individual number or covering all
                If Y_Coord% < 6840 Then
                    Find_The_Mouse = Bet.IsHorn
                Else
                    'Determine if 2,3,11 or 12
                    If Y_Coord% < 7410 Then
                        'Is either 2 or 12
                        If X_Coord% < 8145 Then
                            Find_The_Mouse = Bet.IsHorn2
                        Else
                            Find_The_Mouse = Bet.IsHorn12
                        End If
                    Else
                        'Is either 3 or 11
                        If X_Coord% < 8145 Then
                            Find_The_Mouse = Bet.IsHorn3
                        Else
                            Find_The_Mouse = Bet.IsHorn11
                        End If
                    End If
                End If
            End If
        End If
    Else
        'Could be Pass, Come, Odds, Place, Field, Big 6 or Big 8 or off the field
        If X_Coord% < 450 Or Y_Coord% > 8025 Or Y_Coord% < 930 Then
            'Off the playing field
            Find_The_Mouse = Bet.IsNotValid
            'Nothing else to do so OK to exit
            Exit Function
        End If
        If Y_Coord% < 3720 Then
            'In the Numbers section so determine which Number
            NumberBet% = Find_The_Number(X_Coord%)
            'Is either Place or an Odds bet
            If Y_Coord% > 2055 Then
                Select Case NumberBet%
                    Case 4
                        Find_The_Mouse = Bet.IsCome4Odds
                    Case 5
                        Find_The_Mouse = Bet.IsCome5Odds
                    Case 6
                        Find_The_Mouse = Bet.IsCome6Odds
                    Case 8
                        Find_The_Mouse = Bet.IsCome8Odds
                    Case 9
                        Find_The_Mouse = Bet.IsCome9Odds
                    Case 10
                        Find_The_Mouse = Bet.IsCome10Odds
                End Select
            Else
                'Come bet Odds eliminated but may be Odds on a Don't bet
                If Y_Coord% < 1455 Then
                    Select Case NumberBet%
                        Case 4
                            Find_The_Mouse = Bet.IsDont4Odds
                        Case 5
                            Find_The_Mouse = Bet.IsDont5Odds
                        Case 6
                            Find_The_Mouse = Bet.IsDont6Odds
                        Case 8
                            Find_The_Mouse = Bet.IsDont8Odds
                        Case 9
                            Find_The_Mouse = Bet.IsDont9Odds
                        Case 10
                            Find_The_Mouse = Bet.IsDont10Odds
                    End Select
                Else
                    'Both Odds bets eliminated so must be a Place bet
                    Select Case NumberBet%
                        Case 4
                            Find_The_Mouse = Bet.IsPlace4
                        Case 5
                            Find_The_Mouse = Bet.IsPlace5
                        Case 6
                            Find_The_Mouse = Bet.IsPlace6
                        Case 8
                            Find_The_Mouse = Bet.IsPlace8
                        Case 9
                            Find_The_Mouse = Bet.IsPlace9
                        Case 10
                            Find_The_Mouse = Bet.IsPlace10
                    End Select
                End If
            End If
        Else
            'Could be Come, Don't, Field, Big 6, Big  8 or Pass
            If Y_Coord% < 5865 Then
                'Is either Come or Don't
                If Y_Coord% < 4800 Then
                    Find_The_Mouse = Bet.IsCome
                Else
                    Find_The_Mouse = Bet.IsDont
                End If
            Else
                'Could be Big 6, Big 8, Field or Pass
                If Y_Coord% > 6960 Then
                    Find_The_Mouse = Bet.IsPass
                Else
                    'Could be Big 6, Big 8 or Field
                    If X_Coord% > 2355 Then
                        Find_The_Mouse = Bet.IsField
                    Else
                        'Is either Big 6 or Big 8
                        If X_Coord% < 1395 Then
                            Find_The_Mouse = Bet.IsBig6
                        Else
                            Find_The_Mouse = Bet.IsBig8
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Private Function Valid_Amount(ByVal BarIndex As Integer, Optional ByVal KillMessage As Variant) As Integer
    'Make sure there is enough money in the bank to cover the bet or the bet high enough to cover the recall.
    'If amount is valid for the action the return the amount.
    'If there is a problem then return zero as the amount
    
    With Layout
        'Determine the amount to test based on the bet amount image's index passed in
        Select Case BarIndex%
            Case 0
                Valid_Amount = 1
            Case 1
                Valid_Amount = -1
            Case 2
                Valid_Amount = 5
            Case 3
                Valid_Amount = -5
            Case 4
                Valid_Amount = 25
            Case 5
                Valid_Amount = -25
            Case 6
                Valid_Amount = 100
            Case 7
                Valid_Amount = -100
            Case Else
                'The player may just drag the existing bet without changing the value in which case the BarIndex value will be set to 99
                Valid_Amount = Val(.pic_GoldBar(0).Tag)
        End Select
    
        'Make sure the bet will not go over the limit of $200 (less than zero will tested below)
        If BarIndex% < 99 Then
            If Valid_Amount + Val(.pic_GoldBar(0).Tag) > 200 Then
                'Over the limit...
                .lbl_Message.Caption = "Adding $" & CStr(Valid_Amount) & "will exceed the maximum bet limit of $200."
                Beep
                'Return zero as the amount
                Valid_Amount = 0
            End If
        End If
    
        'If bet did not go over the limit then test the amount against either the bankroll or the current bet amount.
        'If the amount is greater than zero, the player is adding to the bet so make sure the bank has at least
        'The amount of the current bet.  If the amount is less than zero then the player is recalling from the
        'total bet so make sure that the bet is at least as much as the recalled amount.
        If Valid_Amount > 0 Then
            'Player is adding to the bet
            If Val(.lbl_BankRoll.Caption) < Valid_Amount Then
                'Amount to bet is larger than bankroll
                If IsMissing(KillMessage) Then
                    .lbl_Message.Caption = "You do not have enough money to bet $" & CStr(Valid_Amount) & "."
                    Beep
                End If
                'Return zero as the amount
                Valid_Amount = 0
            End If
        Else
            'Player is recalling from the bet
            If Val(.pic_GoldBar(0).Tag) < Abs(Valid_Amount) Then
                'Recall amount is larger than current bet so take bet down to zero
                .pic_GoldBar(0).Tag = vbNullString
                .pic_GoldBar(0).Visible = False
                'Return zero as the amount
                Valid_Amount = 0
            End If
        End If
        
    End With
    
End Function

Public Function Is_Legal_Bet(ByVal Mouse_X As Single, Mouse_Y As Single) As Integer
    'Determine which bet the player wants to place.
    'If the bet is legal to make then return the bet number otherwise return false.
    
    Dim BetPlaced As Integer
    
    BetPlaced% = Find_The_Mouse(CInt(Mouse_X!), CInt(Mouse_Y!))
    If BetPlaced% > 0 Then
        'Player has clicked somewhere within the boundries of the playing field so test if the bet is legal to make
        If Test_Legal_Bet(BetPlaced%) Then
            'This is a legal bet so return the bet number
            Is_Legal_Bet = BetPlaced%
        Else
            'This bet is not legal
            Is_Legal_Bet = False
        End If
    Else
        'Player's bet was dropped out of the playing field so not legal
        Is_Legal_Bet = False
    End If

End Function

Function Test_Legal_Bet(ByRef BetNumber As Integer) As Boolean
    'Determine if it is legal to place the current bet.
    'If the bet is legal to make then return the bet number otherwise return false
    '(This function is called from Is_Legal_Bet which has already determined what
    ' bet the player wants to make and passed it to this function in the BetNumber parameter)
    
    Dim KillMessage As Boolean
    '(KillMessage serves somewhat of a dual purpose in this function.  If we get to the end and Msg$
    ' is null and KillMessage is false then no problems where found.  If there was a problem found then
    ' we look at KillMessage's value to determine if we need to display a message or not.)
    Dim Msg As String
    
    '(Pass bet can only be made on a Come Out Roll, can only lay Odds if Come bet already exists etc.)
    Select Case BetNumber%
        Case Bet.IsPass
            'Must be Come Out roll to make a Pass bet but player may be making an Odds bet also
            If Not gComeOutRoll Then
                'See if there is an existing Pass bet
                If gPlacedBet%(Bet.IsPass) = 0 Then
                    Msg$ = "You can only make a Pass bet on a Come Out Roll."
                Else
                    'OK to make an Odds bet so change the BetNumber
                    BetNumber% = Bet.IsPassOdds
                End If
            End If
        Case Bet.IsCome
            'Can't be a Come Out Roll
            If gComeOutRoll Then
                Msg$ = "You can't make a Come bet on a Come Out Roll."
            End If
        Case Bet.IsCome4Odds
            'Must have a working Come bet on the 4 before taking Odds
            If gPlacedBet%(Bet.IsCome4) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 4 to take Odds on."
            End If
        Case Bet.IsCome5Odds
            'Must have a working Come bet on the 5 before taking Odds
            If gPlacedBet%(Bet.IsCome5) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 5 to take Odds on."
            End If
        Case Bet.IsCome6Odds
            'Must have a working Come bet on the 6 before taking Odds
            If gPlacedBet%(Bet.IsCome6) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 6 to take Odds on."
            End If
        Case Bet.IsCome8Odds
            'Must have a working Come bet on the 8 before taking Odds
            If gPlacedBet%(Bet.IsCome8) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 8 to take Odds on."
            End If
        Case Bet.IsCome9Odds
            'Must have a working Come bet on the 9 before taking Odds
            If gPlacedBet%(Bet.IsCome9) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 9 to take Odds on."
            End If
        Case Bet.IsCome10Odds
            'Must have a working Come bet on the 10 before taking Odds
            If gPlacedBet%(Bet.IsCome10) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Come bet on the 10 to take Odds on."
            End If
        Case Bet.IsDont4Odds
            'Must have a working Don't bet on the 4 before taking Odds
            If gPlacedBet%(Bet.IsDont4) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 4 to take Odds on."
            End If
        Case Bet.IsDont5Odds
            'Must have a working Don't bet on the 5 before taking Odds
            If gPlacedBet%(Bet.IsDont5) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 5 to take Odds on."
            End If
        Case Bet.IsDont6Odds
            'Must have a working Don't bet on the 6 before taking Odds
            If gPlacedBet%(Bet.IsDont6) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 6 to take Odds on."
            End If
        Case Bet.IsDont8Odds
            'Must have a working Don't bet on the 8 before taking Odds
            If gPlacedBet%(Bet.IsDont8) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 8 to take Odds on."
            End If
        Case Bet.IsDont9Odds
            'Must have a working Don't bet on the 9 before taking Odds
            If gPlacedBet%(Bet.IsDont9) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 9 to take Odds on."
            End If
        Case Bet.IsDont10Odds
            'Must have a working Don't bet on the 10 before taking Odds
            If gPlacedBet%(Bet.IsDont10) > 0 Then
                'Validate the Odds bet
                If Not Valid_Odds(BetNumber%) Then
                    KillMessage = True
                End If
            Else
                Msg$ = "You do not have a Don't bet on the 10 to take Odds on."
            End If
        '(Commenting the Case Else like this is a note to myself.  Two months from now I won't remember if there should
        ' have been some exception code here or not so the commenting helps.)
        'Case Else
            'All other bets can be placed at any time so no exception code is needed here
    End Select
    
    'If there is no message assigned and the KillMessage flag is False then all went well so return the bet number.
    'Otherwise, there was problem so show the message if needed and return False
    If LenB(Msg$) = 0 And Not KillMessage Then
        Test_Legal_Bet = BetNumber%
    Else
        'If Valid_Odds already dispalyed a message then KillMessage will be true, otherwise we
        'need to display the problem we found in this function.
        If Not KillMessage Then
            Layout.lbl_Message.Caption = Msg$
            Beep
        End If
        'There was a problem so return that this is not a valid bet
        Test_Legal_Bet = Bet.IsNotValid
    End If

End Function

Public Sub Create_The_Bet_Amount(ByVal Index As Integer)
    'Player has clicked on a Bet Image so determine if adding to recalling from the current bet and update the bet amount and bankroll
    Dim TestAmount As Integer
    
    'Validate the amount and action
    TestAmount% = Valid_Amount(Index%)
    'See if a bet amount was assigned
    If TestAmount% <> 0 Then
        With Layout
            'The bet amount and action were valid so increment/decrement the bet
            '(TestAmount% may be + or - so actual addition or subtraction to bet amount will be deterimed by that factor)
            .pic_GoldBar(0).Tag = CStr(Val(.pic_GoldBar(0).Tag) + TestAmount%)
            '(The bankroll will get deducted when the bet is placed in Place_New_Bet)
            'Update the current bet image
            If Val(.pic_GoldBar(0).Tag) > 0 Then
                .pic_GoldBar(0).Picture = .ImageList1.ListImages(Val(.pic_GoldBar(0).Tag)).Picture
                .pic_GoldBar(0).Visible = True
            Else
                'Bet is zero so make the image invisible
                .pic_GoldBar(0).Visible = False
            End If
        End With
    End If
    
End Sub

Public Sub Assign_Come_Bet(ByVal NewComeBet As Integer)
    'This routine reassigns a Come bet to a specific Number
    Dim ComeNumber As Integer
    
    'Determine the value of the bet for the Number
    Select Case NewComeBet
        Case 4
            ComeNumber% = Bet.IsCome4
        Case 5
            ComeNumber% = Bet.IsCome5
        Case 6
            ComeNumber% = Bet.IsCome6
        Case 8
            ComeNumber% = Bet.IsCome8
        Case 9
            ComeNumber% = Bet.IsCome9
        Case 10
            ComeNumber% = Bet.IsCome10
    End Select
    
    'Create a new bet for the Number
    With Layout
        Load .pic_GoldBar(ComeNumber%)
        'Set up the new picture to be the same as the Come Bet picture
        .pic_GoldBar(ComeNumber%).Picture = .pic_GoldBar(Bet.IsCome).Picture
        'Move it in to position
        .pic_GoldBar(ComeNumber%).Move 570 + (gDiceRoll% - 4 + (gDiceRoll% > 7)) * 1095, 2175
        .pic_GoldBar(ComeNumber%).Visible = True
        'Disable the control so other bets can be dragged over the top of it
        .pic_GoldBar(ComeNumber%).Enabled = False
        'Record the amount of the bet
        gPlacedBet%(ComeNumber%) = gPlacedBet%(Bet.IsCome)
    End With
    
End Sub

Public Sub Assign_Dont_Bet(ByVal NewDontBet As Integer)
    'This routine reassigns a Dont bet to a specific Number
    Dim DontNumber As Integer
    
    'Determine the value of the bet for the Number
    Select Case NewDontBet
        Case 4
            DontNumber% = Bet.IsDont4
        Case 5
            DontNumber% = Bet.IsDont5
        Case 6
            DontNumber% = Bet.IsDont6
        Case 8
            DontNumber% = Bet.IsDont8
        Case 9
            DontNumber% = Bet.IsDont9
        Case 10
            DontNumber% = Bet.IsDont10
    End Select
    
    'Create a new bet for the Number
    With Layout
        Load .pic_GoldBar(DontNumber%)
        'Set up the new picture to be the same as the Dont Bet picture
        .pic_GoldBar(DontNumber%).Picture = .pic_GoldBar(Bet.IsDont).Picture
        'Move it in to position
        .pic_GoldBar(DontNumber%).Move 570 + (gDiceRoll% - 4 + (gDiceRoll% > 7)) * 1095, 1000
        .pic_GoldBar(DontNumber%).Visible = True
        'Disable the control so other bets can be dragged over the top of it
        .pic_GoldBar(DontNumber%).Enabled = False
        'Record the amount of the bet
        gPlacedBet%(DontNumber%) = gPlacedBet%(Bet.IsDont)
    End With
    
End Sub

Public Sub Place_New_Bet(ByVal BetNumber As Integer)
    'This routine places the gold bar in the proper location and records the amount of the bet.
    'Set the Top and left for the Bet image based on the Bet number.
    'The basic idea is to center the bar on the smaller areas such as Hardway or Horn bets and
    'on the larger areas such as Pass and Come, just keep the bar inside the boundry lines.  If it is in side then
    'just leave it where it was dropped
    Dim New_X As Integer
    Dim New_Y As Integer
    Dim KillMessage As Boolean
    
    With Layout.pic_GoldBar(0)
        Select Case BetNumber%
            Case Bet.IsNotValid
                'This bet is not a legal bet so reset it
                '("All things in moderation...." -
                ' I don't have a problem with an occasional GOTO.  So long as it is not used to jump all over
                ' the place and make the code hard to follow, GOTO comes in pretty handy now and then.)
                GoTo ResetBetZero
            Case Bet.IsPass
                'Keep image inside the box
                If .Left < 495 Then New_X% = 495
                If .Left > 6120 Then New_X% = 6120
                If .Top < 7020 Then New_Y% = 7020
                If .Top > 7590 Then New_Y% = 7590
            Case Bet.IsPassOdds
                'Place the Odds bet so it overlaps the Pass bet
                New_X% = Layout.pic_GoldBar(Bet.IsPass).Left + 150
                New_Y% = Layout.pic_GoldBar(Bet.IsPass).Top + 150
            Case Bet.IsCome
                If .Left < 495 Then New_X% = 495
                If .Left > 6120 Then New_X% = 6120
                If .Top < 3780 Then New_Y% = 3780
                If .Top > 4350 Then New_Y% = 4350
            Case Bet.IsCome4Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome4).Move Layout.pic_GoldBar(Bet.IsCome4).Left, Layout.pic_GoldBar(Bet.IsCome4).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 675
                New_Y% = 2100
            Case Bet.IsCome5Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome5).Move Layout.pic_GoldBar(Bet.IsCome5).Left, Layout.pic_GoldBar(Bet.IsCome5).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 1770
                New_Y% = 2100
            Case Bet.IsCome6Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome6).Move Layout.pic_GoldBar(Bet.IsCome6).Left, Layout.pic_GoldBar(Bet.IsCome6).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 2865
                New_Y% = 2100
            Case Bet.IsCome8Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome8).Move Layout.pic_GoldBar(Bet.IsCome8).Left, Layout.pic_GoldBar(Bet.IsCome8).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 3960
                New_Y% = 2100
            Case Bet.IsCome9Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome9).Move Layout.pic_GoldBar(Bet.IsCome9).Left, Layout.pic_GoldBar(Bet.IsCome9).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 5055
                New_Y% = 2100
            Case Bet.IsCome10Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsCome10).Move Layout.pic_GoldBar(Bet.IsCome10).Left, Layout.pic_GoldBar(Bet.IsCome10).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 6150
                New_Y% = 2100
            Case Bet.IsDont
                    If .Left < 495 Then New_X% = 495
                    If .Left > 6120 Then New_X% = 6120
                    If .Top < 4860 Then New_Y% = 4860
                    If .Top > 5430 Then New_Y% = 5430
            Case Bet.IsDont4Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont4).Move Layout.pic_GoldBar(Bet.IsDont4).Left, Layout.pic_GoldBar(Bet.IsDont4).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 675
                New_Y% = 930
            Case Bet.IsDont5Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont5).Move Layout.pic_GoldBar(Bet.IsDont5).Left, Layout.pic_GoldBar(Bet.IsDont5).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 1770
                New_Y% = 930
            Case Bet.IsDont6Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont6).Move Layout.pic_GoldBar(Bet.IsDont6).Left, Layout.pic_GoldBar(Bet.IsDont6).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 2865
                New_Y% = 930
            Case Bet.IsDont8Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont8).Move Layout.pic_GoldBar(Bet.IsDont8).Left, Layout.pic_GoldBar(Bet.IsDont8).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 3960
                New_Y% = 930
            Case Bet.IsDont9Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont9).Move Layout.pic_GoldBar(Bet.IsDont9).Left, Layout.pic_GoldBar(Bet.IsDont9).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 5055
                New_Y% = 930
            Case Bet.IsDont10Odds
                'Move the original bet up and to the left
                Layout.pic_GoldBar(Bet.IsDont10).Move Layout.pic_GoldBar(Bet.IsDont10).Left, Layout.pic_GoldBar(Bet.IsDont10).Top
                'Place the Odds bet on top of the original but offset
                New_X% = 6150
                New_Y% = 930
            Case Bet.IsBig6
                New_X% = 495
                New_Y% = 6210
            Case Bet.IsBig8
                New_X% = 1470
                New_Y% = 6210
            Case Bet.IsField
                If .Left < 2460 Then New_X% = 2460
                If .Left > 6120 Then New_X% = 6120
                If .Top < 5955 Then New_Y% = 5955
                If .Top > 6510 Then New_Y% = 6510
            Case Bet.IsPlace4
                New_X% = 570
                New_Y% = 1560
            Case Bet.IsPlace5
                New_X% = 1665
                New_Y% = 1560
            Case Bet.IsPlace6
                New_X% = 2760
                New_Y% = 1560
            Case Bet.IsPlace8
                New_X% = 3855
                New_Y% = 1560
            Case Bet.IsPlace9
                New_X% = 4950
                New_Y% = 1560
            Case Bet.IsPlace10
                New_X% = 6045
                New_Y% = 1560
            Case Bet.IsHard
                'Player wants to cover the hardway bets so need to test if there is enough money in the bank
                If Val(Layout.lbl_BankRoll.Caption) < Val(.Tag) * 4 Then
                    Layout.lbl_Message.Caption = "You do not enough money to place " & .Tag & " on each Hardway."
                    Beep
                    Exit Sub
                Else
                    'Place a bet on each of the hardways
                    Call Cover_Hardways
                    Exit Sub
                End If
            Case Bet.IsHard4
                New_X% = 7185
                New_Y% = 4200
            Case Bet.IsHard6
                New_X% = 7185
                New_Y% = 3630
            Case Bet.IsHard8
                New_X% = 8280
                New_Y% = 3630
            Case Bet.IsHard10
                New_X% = 8280
                New_Y% = 4200
            Case Bet.IsAny7
                If .Left < 7070 Then New_X% = 7070
                If .Left > 8385 Then New_X% = 8385
                If .Top < 4785 Then New_Y% = 4785
                If .Top > 4890 Then New_Y% = 4890
            Case Bet.IsCraps
                If .Left < 7070 Then New_X% = 7070
                If .Left > 8385 Then New_X% = 8385
                If .Top < 5415 Then New_Y% = 5415
                If .Top > 5550 Then New_Y% = 5550
            Case Bet.IsHorn
                'Player wants to cover the hardway bets so need to test if there is enough money in the bank
                If Val(Layout.lbl_BankRoll.Caption) < Val(.Tag) * 4 Then
                    Layout.lbl_Message.Caption = "You do not enough money to place " & .Tag & " on each Horn bet."
                    Beep
                    Exit Sub
                Else
                    'Place a bet on each of the hardways
                    Call Cover_Horn_Bets
                    Exit Sub
                End If
            Case Bet.IsHorn2
                New_X% = 7185
                New_Y% = 6930
            Case Bet.IsHorn3
                New_X% = 7185
                New_Y% = 7500
            Case Bet.IsHorn11
                New_X% = 8280
                New_Y% = 7500
            Case Bet.IsHorn12
                New_X% = 8280
                New_Y% = 6930
        End Select
        
        'If New_X% or New_Y% have not been assigned a new value then Ok to drop the Bet where it is
        If New_X% = 0 Then New_X% = .Left
        If New_Y% = 0 Then New_Y% = .Top
    
    End With
    
    'This is the amount of the current bet so we need to deduct the amount from the bankroll.
    'But first we need to see if the bankroll can cover the amount.
    If Valid_Amount(99, KillMessage) Then
        'OK to deduct this amount from the bankroll
        With Layout.lbl_BankRoll
            .Caption = CStr(Val(.Caption) - Val(Layout.pic_GoldBar(0).Tag))
        End With
    Else
        'There is not enough in the bank to cover another bet of this amount
        Layout.lbl_Message.Caption = "There is not enough money in the bank to cover this bet."
        Beep
        'Reset - Do not generate a new bet
        GoTo ResetBetZero
    End If
    
    'Create a new instance of the bet picture for this bet
    Call Generate_The_Bet(BetNumber%, New_X%, New_Y%)

ResetBetZero:
    'Always move the reference bet image back to the Home position
    '(In the case that the bet was not legal then this is all that happens - no new bet is generated)
    With Layout.pic_GoldBar(0)
        .Move Home_X%, Home_Y%
        .Enabled = True
        .DragMode = 0
    End With
    
End Sub

Private Function Valid_Odds(ByVal BetNumber As Integer) As Boolean
    'This function validates that the amount of the Odds bet does not exceed the limits based on the type and amount of the original bet.
    'May be pressing an odds bet so add the original - if it has not been placed yet then its value is zero so has no effect,
    'if it has been placed then the test for predding is validated properly.
    'If the bet is valid it will return True.
    Dim Msg As String
    
    Select Case BetNumber%
        Case Bet.IsPassOdds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsPassOdds) > gPlacedBet%(Bet.IsPass) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsPass) * 2)
            End If
        Case Bet.IsCome4Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome4Odds) > gPlacedBet%(Bet.IsCome4) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome4) * 2)
            End If
        Case Bet.IsCome5Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome5Odds) > gPlacedBet%(Bet.IsCome5) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome5) * 2)
            End If
        Case Bet.IsCome6Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome6Odds) > gPlacedBet%(Bet.IsCome6) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome6) * 2)
            End If
        Case Bet.IsCome8Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome8Odds) > gPlacedBet%(Bet.IsCome8) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome8) * 2)
            End If
        Case Bet.IsCome9Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome9Odds) > gPlacedBet%(Bet.IsCome9) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome9) * 2)
            End If
        Case Bet.IsCome10Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsCome10Odds) > gPlacedBet%(Bet.IsCome10) * 2 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsCome10) * 2)
            End If
        Case Bet.IsDont4Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont4Odds) > gPlacedBet%(Bet.IsDont4) * 4 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont4) * 4)
            End If
        Case Bet.IsDont5Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont5Odds) > gPlacedBet%(Bet.IsDont5) * 3 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont5) * 3)
            End If
        Case Bet.IsDont6Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont6Odds) > gPlacedBet%(Bet.IsDont6) * 2.5 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont6) * 2.5)
            End If
        Case Bet.IsDont8Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont8Odds) > gPlacedBet%(Bet.IsDont8) * 2.5 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont8) * 2.5)
            End If
        Case Bet.IsDont9Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont9Odds) > gPlacedBet%(Bet.IsDont9) * 3 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont9) * 3)
            End If
        Case Bet.IsDont10Odds
            If Val(Layout.pic_GoldBar(0).Tag) + gPlacedBet%(Bet.IsDont10Odds) > gPlacedBet%(Bet.IsDont10) * 4 Then
                Msg$ = CStr(gPlacedBet%(Bet.IsDont10) * 4)
            End If
    End Select
    
    'If there was no error message assigned to Msg$ then all went well so return true, otherwise display the message
    If LenB(Msg$) = 0 Then
        Valid_Odds = True
    Else
        Layout.lbl_Message.Caption = "You Odds bet may not exceed " & Msg$
        Beep
    End If
    
End Function

Private Sub Generate_The_Bet(ByVal BetNumber As Integer, ByVal Bet_Left As Integer, ByVal Bet_Top As Integer)
    'This routine creates a new instance of the GoldBar picture to be placed on teh layout for this new bet
    With Layout
        'Determine if creating a new bet or resetting an illegal bet attempt
        If BetNumber% > 0 Then
            'This is a legal bet
            'See if the bet already exisits
            If gPlacedBet%(BetNumber) = 0 Then
                Load .pic_GoldBar(BetNumber%)
                .pic_GoldBar(BetNumber%).ZOrder
                'Set up the new picture to be the same as the Bet picture
                .pic_GoldBar(BetNumber%).Picture = .pic_GoldBar(0).Picture
                'Place the new picture on the layout and show it
                .pic_GoldBar(BetNumber%).Move Bet_Left%, Bet_Top%
                .pic_GoldBar(BetNumber%).Visible = True
                'Disable the control so other bets can be dragged over the top of it
                .pic_GoldBar(BetNumber%).Enabled = False
                'Record the amount of the bet
                gPlacedBet%(BetNumber) = Val(.pic_GoldBar(0).Tag)
                'Record this as the last bet placed
                gLastBet% = BetNumber%
            Else
                'This bet already exisits so Press it if it does not exceed any limits
                If gPlacedBet%(BetNumber) + Val(.pic_GoldBar(0).Tag) < 201 Then
                    'This may be an Odds bet so run it through the test for valid Odds bets
                    '(If this is not an Odds bet, then Valid_Odds will default to True)
                    If Valid_Odds(BetNumber%) Then
                        'The bet does not exceed the $200 limit or Odds limit so OK to press it
                        gPlacedBet%(BetNumber) = gPlacedBet%(BetNumber) + Val(.pic_GoldBar(0).Tag)
                        .pic_GoldBar(BetNumber%).Picture = .ImageList1.ListImages(gPlacedBet%(BetNumber)).Picture
                    End If
                Else
                    'Pressing the bet will exceed the limit
                    .lbl_Message.Caption = "Adding $" & .pic_GoldBar(0).Tag & " to the current bet of $" & CStr(gPlacedBet%(BetNumber)) & " will exceed $200."
                    Beep
                    'Reset the bet to the home posistion
                    '(The top and left have been set to the Home positions)
                    .pic_GoldBar(BetNumber%).Move Bet_Left%, Bet_Top%
                End If
            End If
        Else
            'This bet is illegal so just need to move it
            '(The top and left have been set to the Home positions)
            .pic_GoldBar(BetNumber%).Move Bet_Left%, Bet_Top%
        End If
    End With
End Sub

Private Sub Cover_Hardways()
    'Create a new bet for each of the Hardway bets
    Call Place_New_Bet(Bet.IsHard4)
    Call Place_New_Bet(Bet.IsHard6)
    Call Place_New_Bet(Bet.IsHard8)
    Call Place_New_Bet(Bet.IsHard10)
End Sub

Private Sub Cover_Horn_Bets()
    'Create a new new for each of the Horn bets
    Call Place_New_Bet(Bet.IsHorn2)
    Call Place_New_Bet(Bet.IsHorn3)
    Call Place_New_Bet(Bet.IsHorn11)
    Call Place_New_Bet(Bet.IsHorn12)
End Sub

Public Sub Clear_This_Bet(ByVal BetNumber As Integer)
    'This routine resets the bet amount and unloads the image
    'May be passing in an Odds bet that may or may not have been placed so test for an existing value before unloading
    If gPlacedBet%(BetNumber%) > 0 Then
        gPlacedBet%(BetNumber%) = 0
        Unload Layout.pic_GoldBar(BetNumber%)
    End If
End Sub


