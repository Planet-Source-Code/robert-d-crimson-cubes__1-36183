Attribute VB_Name = "Test_Bets"
Option Explicit

'This module does all of the testing to see if a bet was won or lost and displays the results

Public Sub Test_Win_Or_Lose()
    'This routine looks at all the bets to determine if any have been placed.  If a bet has been placed, the proper test routine is called
    'To determine if that bet was won or lost (or there was no effect) on this roll of the dice.
    Dim Index As Integer
    
    'Iterate through the bets to see if any have been placed
    For Index% = 1 To 49
        If gPlacedBet%(Index%) > 0 Then
            'Found a bet that has an amount assigned to it so see if it won or lost
            '(Odds bets are tested for when the parent bet gets tested so they are not in this Case statement)
            Select Case Index%
                Case Bet.IsPass
                    Call Test_Pass_Bet
                Case Bet.IsCome4, Bet.IsCome5, Bet.IsCome6, Bet.IsCome8, Bet.IsCome9, Bet.IsCome10
                    Call Test_Come_Numbers(Index%)
                Case Bet.IsCome
                    Call Test_Come_Bet
                Case Bet.IsDont
                    Call Test_Dont_Bet
                Case Bet.IsDont4, Bet.IsDont5, Bet.IsDont6, Bet.IsDont8, Bet.IsDont9, Bet.IsDont10
                    Call Test_Dont_Numbers(Index%)
                Case Bet.IsBig6
                    Call Test_Big_6_Bet
                Case Bet.IsBig8
                    Call Test_Big_8_Bet
                Case Bet.IsField
                    Call Test_Field_Bet
                Case Bet.IsPlace4, Bet.IsPlace5, Bet.IsPlace6, Bet.IsPlace8, Bet.IsPlace9, Bet.IsPlace10
                    Call Test_Place_Bet(Index%)
                Case Bet.IsHard4, Bet.IsHard6, Bet.IsHard8, Bet.IsHard10
                    Call Test_Hard_Bet(Index%)
                Case Bet.IsAny7
                    Call Test_Any7_Bet
                Case Bet.IsCraps
                    Call Test_Craps_Bet
                Case Bet.IsHorn2, Bet.IsHorn3, Bet.IsHorn11, Bet.IsHorn12
                    Call Test_Horn_Bet(Index%)
            End Select
        End If
    Next

End Sub

Private Sub Test_Dont_Numbers(ByVal BetNumber As Integer)
    'A roll of 7 will win all working Don't bets
    Dim Index As Integer
    
    If gDiceRoll% = 7 Then
        'Iterate through the Don't bets
        For Index% = Bet.IsDont4 To Bet.IsDont10
            'See if the Don't bet has been made
            If gPlacedBet%(Index%) > 0 Then
                'Get the Number for the bet
                Select Case Index%
                    Case Bet.IsDont4
                        'Calculate how much was won, call Won_The_Bet to inform the player and then clear the bet
                        Call Won_The_Bet("Dont: 4", gPlacedBet%(Bet.IsDont4) * 2 + Dont_Odds_Payoff(Bet.IsDont4Odds))
                        Call Clear_This_Bet(Bet.IsDont4Odds)
                    Case Bet.IsDont5
                        Call Won_The_Bet("Dont: 5", gPlacedBet%(Bet.IsDont5) * 2 + Dont_Odds_Payoff(Bet.IsDont5Odds))
                        Call Clear_This_Bet(Bet.IsDont5Odds)
                    Case Bet.IsDont6
                        Call Won_The_Bet("Dont: 6", gPlacedBet%(Bet.IsDont6) * 2 + Dont_Odds_Payoff(Bet.IsDont6Odds))
                        Call Clear_This_Bet(Bet.IsDont6Odds)
                    Case Bet.IsDont8
                        Call Won_The_Bet("Dont: 8", gPlacedBet%(Bet.IsDont8) * 2 + Dont_Odds_Payoff(Bet.IsDont8Odds))
                        Call Clear_This_Bet(Bet.IsDont8Odds)
                    Case Bet.IsDont9
                        Call Won_The_Bet("Dont: 9", gPlacedBet%(Bet.IsDont9) * 2 + Dont_Odds_Payoff(Bet.IsDont9Odds))
                        Call Clear_This_Bet(Bet.IsDont9Odds)
                    Case Bet.IsDont10
                        Call Won_The_Bet("Dont: 10", gPlacedBet%(Bet.IsDont10) * 2 + Dont_Odds_Payoff(Bet.IsDont10Odds))
                        Call Clear_This_Bet(Bet.IsDont10Odds)
                End Select
                Call Clear_This_Bet(Index%)
            End If
        Next
    Else
        'See if the number rolled has a Don't bet against it
         Select Case BetNumber%
            Case Bet.IsDont4
                If gDiceRoll% = 4 Then
                    'Calculate how much was lost, call Lost_The_Bet to inform the player and then clear the bet
                    Call Lost_The_Bet("Dont: 4", gPlacedBet%(Bet.IsDont4) + gPlacedBet%(Bet.IsDont4Odds))
                    Call Clear_This_Bet(Bet.IsDont4)
                    Call Clear_This_Bet(Bet.IsDont4Odds)
                End If
            Case Bet.IsDont5
                If gDiceRoll% = 5 Then
                    Call Lost_The_Bet("Dont: 5", gPlacedBet%(Bet.IsDont5) + gPlacedBet%(Bet.IsDont5Odds))
                    Call Clear_This_Bet(Bet.IsDont5)
                    Call Clear_This_Bet(Bet.IsDont5Odds)
                End If
            Case Bet.IsDont6
                If gDiceRoll% = 6 Then
                    Call Lost_The_Bet("Dont: 6", gPlacedBet%(Bet.IsDont6) + gPlacedBet%(Bet.IsDont6Odds))
                    Call Clear_This_Bet(Bet.IsDont6)
                    Call Clear_This_Bet(Bet.IsDont6Odds)
                End If
            Case Bet.IsDont8
                If gDiceRoll% = 8 Then
                    Call Lost_The_Bet("Dont: 8", gPlacedBet%(Bet.IsDont8) + gPlacedBet%(Bet.IsDont8Odds))
                    Call Clear_This_Bet(Bet.IsDont8)
                    Call Clear_This_Bet(Bet.IsDont8Odds)
                End If
            Case Bet.IsDont9
                If gDiceRoll% = 9 Then
                    Call Lost_The_Bet("Dont: 9", gPlacedBet%(Bet.IsDont9) + gPlacedBet%(Bet.IsDont9Odds))
                    Call Clear_This_Bet(Bet.IsDont9)
                    Call Clear_This_Bet(Bet.IsDont9Odds)
                End If
            Case Bet.IsDont10
                If gDiceRoll% = 10 Then
                    Call Lost_The_Bet("Dont: 10", gPlacedBet%(Bet.IsDont10) + gPlacedBet%(Bet.IsDont10Odds))
                    Call Clear_This_Bet(Bet.IsDont10)
                    Call Clear_This_Bet(Bet.IsDont10Odds)
                End If
        End Select
    End If
    
End Sub

Private Sub Test_Come_Bet()
    'All numbers affect a Come bet so determine the proper action for this roll
    
    Select Case gDiceRoll%
        Case 2, 3, 12
            'We have a losing Come bet
            Call Lost_The_Bet("Come", gPlacedBet%(Bet.IsCome))
        Case 7, 11
            'We have a winning Come bet
            Call Won_The_Bet("Come", gPlacedBet%(Bet.IsCome) * 2)
        Case 4, 5, 6, 8, 9, 10
            'A Number has been rolled so this Come bet needs to get reassigned to the Number
            Call Assign_Come_Bet(gDiceRoll%)
    End Select
    
    'The Come bet either won, lost or got reassigned so it always gets cleared
    Call Clear_This_Bet(Bet.IsCome)
    
End Sub

Public Sub Test_Pass_Bet()
    'Test if a Pass bet wins or loses - may on the come out roll or shooting for a point
    'Determine if we are on a Come Out roll or not
    If gComeOutRoll Then
        'Test for Natural
        If gDiceRoll% = 7 Or gDiceRoll% = 11 Then
            'Winner on the line
            Call Won_The_Bet("Pass", gPlacedBet%(Bet.IsPass) * 2)
            Call Clear_This_Bet(Bet.IsPass)
        End If
        'Test for Craps
        If gDiceRoll% = 12 Or gDiceRoll < 4 Then
            'Loser on the line
            Call Lost_The_Bet("Pass", gPlacedBet%(Bet.IsPass))
            Call Clear_This_Bet(Bet.IsPass)
        End If
    Else
        'A Point has been established so determine if shooter rolled a 7 or made the Point
        If gDiceRoll% = 7 Then
            'Sevened out
            Call Lost_The_Bet("Pass", gPlacedBet%(Bet.IsPass))
            Call Clear_This_Bet(Bet.IsPass)
            Call Clear_This_Bet(Bet.IsPassOdds)
        Else
            'See if the current roll is the same as the Point
            If gDiceRoll% = gGamePoint% Then
                'Shooter made the Point - Payoff the bet plus any Odds bet winnings
                Call Won_The_Bet("Pass", gPlacedBet%(Bet.IsPass) * 2 + Pass_Come_Odds_Payoff(Bet.IsPassOdds))
                Call Clear_This_Bet(Bet.IsPass)
                Call Clear_This_Bet(Bet.IsPassOdds)
            End If
        End If
    End If
End Sub


Public Sub Test_Horn_Bet(ByVal BetNumber As Integer)
   'This routine tests the Hardway bets
    'The only winning rolls would b 2,3,11 or 12 - anything else would be a loser for any or all Horn bets
    Dim Index As Integer
    
    'Determine the Horn Number for the bet
    Select Case BetNumber%
        Case Bet.IsHorn2
            If gDiceRoll% = 2 Then
                Call Won_The_Bet("Horn: 2", gPlacedBet%(Bet.IsHorn2) * 31)
            Else
                Call Lost_The_Bet("Horn: 2", gPlacedBet%(Bet.IsHorn2))
                Call Clear_This_Bet(BetNumber%)
            End If
        Case Bet.IsHorn3
            If gDiceRoll% = 3 Then
                Call Won_The_Bet("Horn: 3", gPlacedBet%(Bet.IsHorn3) * 16)
            Else
                Call Lost_The_Bet("Horn: 3", gPlacedBet%(Bet.IsHorn3))
                Call Clear_This_Bet(BetNumber%)
            End If
        Case Bet.IsHorn11
            If gDiceRoll% = 11 Then
                Call Won_The_Bet("Horn: 11", gPlacedBet%(Bet.IsHorn11) * 16)
            Else
                Call Lost_The_Bet("Horn: 11", gPlacedBet%(Bet.IsHorn11))
                Call Clear_This_Bet(BetNumber%)
            End If
        Case Bet.IsHorn12
            If gDiceRoll% = 12 Then
                Call Won_The_Bet("Horn: 12", gPlacedBet%(Bet.IsHorn12) * 31)
            Else
                Call Lost_The_Bet("Horn: 12", gPlacedBet%(Bet.IsHorn12))
                Call Clear_This_Bet(BetNumber%)
            End If
    End Select

End Sub


Public Sub Test_Hard_Bet(ByVal BetNumber As Integer)
   'This routine tests the Hardway bets
    'A 7 will lose all the current Hardway bets, a Number rolled Easy that has a working Hardway bet will lose.
    'A Number rolled Hard that has a matching Hard bet will win
    Dim Index As Integer
    
    If gDiceRoll% = 7 Then
        'Lost the Hardway bet(s)
        '(This loop will clear all of the current Hardways so 7 will only be dealt with once no matter how many Hardway bets there are.)
        For Index% = Bet.IsHard4 To Bet.IsHard10
            'See if the Hardway has been bet
            If gPlacedBet%(Index%) > 0 Then
                'Get the Number for the bet
                Select Case Index%
                    Case Bet.IsHard4
                        Call Lost_The_Bet("Hard: 4", gPlacedBet%(Bet.IsHard4))
                    Case Bet.IsHard6
                        Call Lost_The_Bet("Hard: 6", gPlacedBet%(Bet.IsHard6))
                    Case Bet.IsHard8
                        Call Lost_The_Bet("Hard: 8", gPlacedBet%(Bet.IsHard8))
                    Case Bet.IsHard10
                        Call Lost_The_Bet("Hard: 10", gPlacedBet%(Bet.IsHard10))
                End Select
                Call Clear_This_Bet(Index%)
            End If
        Next
    Else
        'Test for a winning Place bet
        'Determine the Number value for the bet number and see if that Number was rolled
        'If the Number was rolled then test is rolled the hard way or not by checking the value of thegRolledHard flag
        Select Case BetNumber%
            Case Bet.IsHard4
                If gDiceRoll% = 4 Then
                    If gRolledHard Then
                        Call Won_The_Bet("Hard: 4", gPlacedBet%(Bet.IsHard4) * 8)
                    Else
                        Call Lost_The_Bet("Hard: 4", gPlacedBet%(Bet.IsHard4))
                        Call Clear_This_Bet(BetNumber%)
                    End If
                End If
            Case Bet.IsHard6
                If gDiceRoll% = 6 Then
                    If gRolledHard Then
                        Call Won_The_Bet("Hard: 6", gPlacedBet%(Bet.IsHard6) * 6)
                    Else
                        Call Lost_The_Bet("Hard: 6", gPlacedBet%(Bet.IsHard6))
                        Call Clear_This_Bet(BetNumber%)
                    End If
                End If
            Case Bet.IsHard8
                If gDiceRoll% = 8 Then
                    If gRolledHard Then
                        Call Won_The_Bet("Hard: 8", gPlacedBet%(Bet.IsHard8) * 7)
                    Else
                        Call Lost_The_Bet("Hard: 8", gPlacedBet%(Bet.IsHard8))
                        Call Clear_This_Bet(BetNumber%)
                    End If
                End If
            Case Bet.IsHard10
                If gDiceRoll% = 10 Then
                    If gRolledHard Then
                        Call Won_The_Bet("Hard: 10", gPlacedBet%(Bet.IsHard10) * 9)
                    Else
                        Call Lost_The_Bet("Hard: 10", gPlacedBet%(Bet.IsHard10))
                        Call Clear_This_Bet(BetNumber%)
                    End If
                End If
        End Select
    End If
End Sub

Public Sub Test_Big_6_Bet()
    'Only need to test for rolling a 6 or 7
    Select Case gDiceRoll%
        Case 6 'Winner
            Call Won_The_Bet("Big 6", gPlacedBet%(Bet.IsBig6) * 2)
            Call Clear_This_Bet(Bet.IsBig6)
        Case 7 'Loser
            Call Lost_The_Bet("Big 6", gPlacedBet%(Bet.IsBig6))
            Call Clear_This_Bet(Bet.IsBig6)
    End Select
End Sub

Public Sub Test_Big_8_Bet()
    'Only need to test for rolling an 8 or 7
    Select Case gDiceRoll%
        Case 8 'Winner
            Call Won_The_Bet("Big 8", gPlacedBet%(Bet.IsBig8) * 2)
            Call Clear_This_Bet(Bet.IsBig8)
        Case 7 'Loser
            Call Lost_The_Bet("Big 8", gPlacedBet%(Bet.IsBig8))
            Call Clear_This_Bet(Bet.IsBig8)
    End Select
End Sub

Public Sub Test_Field_Bet()
    'Winning rolls are 2,3,4,9,10,11,12 - anything else loses
    Dim PayOff As Integer
    
    Select Case gDiceRoll%
        Case 2 'Pays 2:1
            PayOff% = gPlacedBet%(Bet.IsField) * 3
        Case 3, 4, 9, 10, 11 'Pays (1:1)
            PayOff% = gPlacedBet%(Bet.IsField) * 2
        Case 5, 6, 7, 8 'Lost
            PayOff% = 0
        Case 12 'Pays (3:1)
            PayOff% = gPlacedBet%(Bet.IsField) * 4
    End Select
    
    'If Payoff% contains a value then a winner was rolled
    If PayOff% > 0 Then
        Call Won_The_Bet("Field", PayOff%)
        Call Clear_This_Bet(Bet.IsField)
    Else
        Call Lost_The_Bet("Field", gPlacedBet%(Bet.IsField))
        Call Clear_This_Bet(Bet.IsField)
    End If
    
End Sub

Public Sub Test_Place_Bet(ByVal BetNumber As Integer)
    'This routine tests the Place bets
    'A 7 will lose all the current Place bets, a Number rolled that has a working Place bet will win
    Dim Index As Integer
    
    If gDiceRoll% = 7 Then
        'Lost the Place bet(s)
        '(This loop will clear all of the current Place bets so 7 will only be dealt with once no matter how many Place bets there are.)
        For Index% = Bet.IsPlace4 To Bet.IsPlace10
            'See if the Place bet has been made
            If gPlacedBet%(Index%) > 0 Then
                'Get the Number for the bet
                Select Case Index%
                    Case Bet.IsPlace4
                        Call Lost_The_Bet("Place: 4", gPlacedBet%(Bet.IsPlace4))
                    Case Bet.IsPlace5
                        Call Lost_The_Bet("Place: 5", gPlacedBet%(Bet.IsPlace5))
                    Case Bet.IsPlace6
                        Call Lost_The_Bet("Place: 6", gPlacedBet%(Bet.IsPlace6))
                    Case Bet.IsPlace8
                        Call Lost_The_Bet("Place: 8", gPlacedBet%(Bet.IsPlace8))
                    Case Bet.IsPlace9
                        Call Lost_The_Bet("Place: 9", gPlacedBet%(Bet.IsPlace9))
                    Case Bet.IsPlace10
                        Call Lost_The_Bet("Place: 10", gPlacedBet%(Bet.IsPlace10))
                End Select
                Call Clear_This_Bet(Index%)
            End If
        Next
    Else
        'Test for a winning Place bet
        'Determine the Number value for the bet number and see if that Number was rolled
        Select Case BetNumber%
            Case Bet.IsPlace4
                If gDiceRoll% = 4 Then
                    Call Won_The_Bet("Place: 4", Int(gPlacedBet%(Bet.IsPlace4) * 7 / 5))
                End If
            Case Bet.IsPlace5
                If gDiceRoll% = 5 Then
                    Call Won_The_Bet("Place: 5", Int(gPlacedBet%(Bet.IsPlace5) * 7 / 5))
                End If
            Case Bet.IsPlace6
                If gDiceRoll% = 6 Then
                    Call Won_The_Bet("Place: 6", Int(gPlacedBet%(Bet.IsPlace6) * 7 / 6))
                End If
            Case Bet.IsPlace8
                If gDiceRoll% = 8 Then
                    Call Won_The_Bet("Place: 8", Int(gPlacedBet%(Bet.IsPlace8) * 7 / 6))
                End If
            Case Bet.IsPlace9
                If gDiceRoll% = 9 Then
                    Call Won_The_Bet("Place: 9", Int(gPlacedBet%(Bet.IsPlace9) * 7 / 5))
                End If
            Case Bet.IsPlace10
                If gDiceRoll% = 10 Then
                    Call Won_The_Bet("Place: 10", Int(gPlacedBet%(Bet.IsPlace10) * 7 / 5))
                End If
        End Select
    End If
    
End Sub

Private Sub Test_Dont_Bet()
    'This routine tests the Don't Line bet
    '2 or 3 is a win, 7 & 11 are losers, 12 is barred.
    'Any Number gets reassigned to that Number
    
    Select Case gDiceRoll%
        Case 2, 3
            Call Won_The_Bet("Don't", gPlacedBet%(Bet.IsDont))
        Case 7, 11
            Call Lost_The_Bet("Don't", gPlacedBet%(Bet.IsDont))
        Case 4, 5, 6, 8, 9, 10
            Call Assign_Dont_Bet(gDiceRoll%)
    End Select
    
    'Clear the bet for every roll but 12
    If gDiceRoll% < 12 Then
        Call Clear_This_Bet(Bet.IsDont)
    End If

End Sub

Private Sub Test_Come_Numbers(ByVal BetNumber As Integer)
    'This routine tests if the individual Come bets have won or lost
    'A 7 will lose all current Come bets otherwise if the number rolled matches a Number with a Come bet it wins
    Dim Index As Integer
    
    If gDiceRoll% = 7 Then
        'Lost the Place bet(s)
        '(This loop will clear all of the current Place bets so 7 will only be dealt with once no matter how many Place bets there are.)
        For Index% = Bet.IsCome4 To Bet.IsCome10
            'See if the Come bet has been made
            If gPlacedBet%(Index%) > 0 Then
                'Get the Number for the bet
                Select Case Index%
                    Case Bet.IsCome4
                        Call Lost_The_Bet("Come: 4", gPlacedBet%(Bet.IsCome4) + gPlacedBet%(Bet.IsCome4Odds))
                        Call Clear_This_Bet(Bet.IsCome4Odds)
                    Case Bet.IsCome5
                        Call Lost_The_Bet("Come: 5", gPlacedBet%(Bet.IsCome5) + gPlacedBet%(Bet.IsCome5Odds))
                        Call Clear_This_Bet(Bet.IsCome5Odds)
                    Case Bet.IsCome6
                        Call Lost_The_Bet("Come: 6", gPlacedBet%(Bet.IsCome6) + gPlacedBet%(Bet.IsCome6Odds))
                        Call Clear_This_Bet(Bet.IsCome6Odds)
                    Case Bet.IsCome8
                        Call Lost_The_Bet("Come: 8", gPlacedBet%(Bet.IsCome8) + gPlacedBet%(Bet.IsCome8Odds))
                        Call Clear_This_Bet(Bet.IsCome8Odds)
                    Case Bet.IsCome9
                        Call Lost_The_Bet("Come: 9", gPlacedBet%(Bet.IsCome9) + gPlacedBet%(Bet.IsCome9Odds))
                        Call Clear_This_Bet(Bet.IsCome9Odds)
                    Case Bet.IsCome10
                        Call Lost_The_Bet("Come: 10", gPlacedBet%(Bet.IsCome10) + gPlacedBet%(Bet.IsCome10Odds))
                        Call Clear_This_Bet(Bet.IsCome10Odds)
                End Select
                Call Clear_This_Bet(Index%)
            End If
        Next
    Else
        'Test for a winning Come bet
        'Determine the Number value for the bet number and see if that Number was rolled
        Select Case BetNumber%
            Case Bet.IsCome4
                If gDiceRoll% = 4 Then
                    Call Won_The_Bet("Come: 4", gPlacedBet%(Bet.IsCome4) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome4Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome4Odds)
                End If
            Case Bet.IsCome5
                If gDiceRoll% = 5 Then
                    Call Won_The_Bet("Come: 5", gPlacedBet%(Bet.IsCome5) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome5Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome5Odds)
                End If
            Case Bet.IsCome6
                If gDiceRoll% = 6 Then
                    Call Won_The_Bet("Come: 6", gPlacedBet%(Bet.IsCome6) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome6Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome6Odds)
                End If
            Case Bet.IsCome8
                If gDiceRoll% = 8 Then
                    Call Won_The_Bet("Come: 8", gPlacedBet%(Bet.IsCome8) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome8Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome8Odds)
                End If
            Case Bet.IsCome9
                If gDiceRoll% = 9 Then
                    Call Won_The_Bet("Come: 9", gPlacedBet%(Bet.IsCome9) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome9Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome9Odds)
                End If
            Case Bet.IsCome10
                If gDiceRoll% = 10 Then
                    Call Won_The_Bet("Come: 10", gPlacedBet%(Bet.IsCome10) * 2 + Pass_Come_Odds_Payoff(Bet.IsCome10Odds))
                    Call Clear_This_Bet(BetNumber%)
                    Call Clear_This_Bet(Bet.IsCome10Odds)
                End If
        End Select
    End If
End Sub

Public Sub Test_Any7_Bet()
    'A roll of 7 wins, anything else loses
    If gDiceRoll% = 7 Then
        Call Won_The_Bet("Any 7", gPlacedBet%(Bet.IsAny7) * 4)
    Else
        Call Lost_The_Bet("Any 7", gPlacedBet%(Bet.IsAny7))
        Call Clear_This_Bet(Bet.IsAny7)
    End If
End Sub

Public Sub Test_Craps_Bet()
    'A roll of 2, 3 or 12 wins, anything else loses
    If gDiceRoll% = 2 Or gDiceRoll% = 3 Or gDiceRoll% = 12 Then
        Call Won_The_Bet("Craps", gPlacedBet%(Bet.IsCraps) * 7)
    Else
        Call Lost_The_Bet("Craps", gPlacedBet%(Bet.IsCraps))
        Call Clear_This_Bet(Bet.IsCraps)
    End If
End Sub

Public Function Pass_Come_Odds_Payoff(ByVal BetNumber As Integer) As Integer
    'This function calculates and returns the amount won for the Odds bet on a Pass or Come bet
    Select Case gDiceRoll%
        Case 2, 4
            Pass_Come_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 3)
        Case 5, 9
            Pass_Come_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 2.5)
        Case 6, 8
            Pass_Come_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 2.2)
    End Select
End Function

Public Function Dont_Odds_Payoff(ByVal BetNumber As Integer) As Integer
    'This function calculates and returns the amount won for the Odds bet on a Don't Pass or Don't Come bet
    Select Case gDiceRoll%
        Case 2, 4
            Dont_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 1.5)
        Case 5, 9
            Dont_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 1.67)
        Case 6, 8
            Dont_Odds_Payoff = Int(gPlacedBet%(BetNumber%) * 1.83)
    End Select
End Function

Public Sub Won_The_Bet(ByVal BetName As String, ByVal PayOff As Integer)
    'Player has won this bet
    With Layout
        .lbl_Message.Caption = "Won $" + CStr(PayOff%) + " on " + BetName
        .lbl_Message.Refresh
        'Add winnings to bank
        .lbl_BankRoll.Caption = CStr(Val(.lbl_BankRoll.Caption) + PayOff%)
        .lbl_BankRoll.Refresh
    End With
    'Bring attention to the message
    Beep
    'Call for the delay
    Call Delay(2.25)
    'Clear the message
    Layout.lbl_Message.Caption = vbNullString
End Sub

Public Sub Lost_The_Bet(ByVal BetName As String, ByVal BetLost As Integer)
    'Player lost this bet.  The amount of the bet was deducted from the bankroll when it was placed so
    'there is no need to deduct the amount lost here.
    With Layout
        .lbl_Message.Caption = "Lost $" + CStr(BetLost%) + " on " + BetName
        .lbl_Message.Refresh
    End With
    'Bring attention to the message
    Beep
    'Call for the delay
    Call Delay(2.25)
    'Clear the message
    Layout.lbl_Message.Caption = vbNullString
End Sub
