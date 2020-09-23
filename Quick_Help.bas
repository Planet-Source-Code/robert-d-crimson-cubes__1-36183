Attribute VB_Name = "Quick_Help"
Option Explicit

'This module is the "What Is This" help.
'The player has pressed the ? and then clicked on a bet.
'This routine will pop up a message box with a brief description of the bet.

Sub What_Is_This_Bet(ByVal BetNumber As Integer)
    Dim Title As String
    Dim Msg As String
    
    Select Case BetNumber%
        Case 0
            Msg$ = "Outside the playing field."
        Case Bet.IsDont4Odds To Bet.isdont10odds
            Title$ = "Odds on Don't Pass or Don't Come"
            Msg$ = Msg$ & "This is where you lay odds against an exisitng Don't Pass or Don't Come bet." & Chr$(13)
            Msg$ = Msg$ & "The Odds will be won or lost in conjuction with the Don't Pass or Don't Come bet." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 4 & 10 = 1:2, 5 & 9  = 2:3, 6 & 8  = 5:6"
        Case Bet.IsCome4Odds To Bet.IsCome10Odds
            Title$ = "Odds on Come Bet"
            Msg$ = Msg$ & "This is where you take odds on an existing Come bet." & Chr$(13)
            Msg$ = Msg$ & "The Odds will be won or lost in conjuction with the Come bet." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 4 & 10 = 2:1, 5 & 9  = 3:2, 6 & 8  = 6:5"
        Case Bet.IsPlace4 To Bet.IsPlace10
            Title$ = "Place"
            Msg$ = Msg$ & "This is where you make a Place bet." & Chr$(13)
            Msg$ = Msg$ & "The number (4,5,6,8,9 or 10) must be rolled before a 7 to win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 4,5,9,10 = 5:7, 6 & 8 = 6:7"
        Case Bet.IsCome
            Title$ = "Come"
            Msg$ = Msg$ & "If the next roll is 7 or 11 you win - 2, 3 or 12 you lose." & Chr$(13)
            Msg$ = Msg$ & "4,5,6,8,9 or 10 becomes your point.  Then, if your point is rolled again before a 7 you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 1:1"
        Case Bet.IsDont
            Title$ = "Don't Pass or Don't Come"
            Msg$ = Msg$ & "If the next roll is 7 or 11 you lose - 2 or 3 you win - 12 has no effect." & Chr$(13)
            Msg$ = Msg$ & "4,5,6,8,9 or 10 becomes your point.  Then, if a 7 is rolled before your point you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 1:1"
        Case Bet.IsBig6
            Title$ = "Big 6"
            Msg$ = Msg$ & "If a 6 is rolled before a 7 you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 1:1"
        Case Bet.IsBig8
            Title$ = "Big 8"
            Msg$ = Msg$ & "If an 8 is rolled before a 7 you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 1:1"
        Case Bet.IsField
            Title$ = "Field"
            Msg$ = Msg$ & "If the next number rolled is in the Field (2,3,4,9,10,11 or 12) you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 2 pays 2:1, 12 pays 3:1 all other numbers pay 1:1"
        Case Bet.IsPass
            Title$ = "Pass"
            Msg$ = Msg$ & "If the next roll is 7 or 11 you win - 2, 3 or 12 you lose." & Chr$(13)
            Msg$ = Msg$ & "4,5,6,8,9 or 10 and that number becomes the Point.  If the Point number is rolled again before a 7 you win." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 1:1"
        Case Bet.IsHard
            Title$ = "Cover Hardways"
            Msg$ = Msg$ & "Dropping a bet here causes that amount to be bet on each hardway bet." & Chr$(13)
            Msg$ = Msg$ & "For example: a $5 bet will result in $5 getting placed on Hard 4, Hard 6, Hard 8 and Hard 10 for a total bet of $20." & Chr$(13)
            Msg$ = Msg$ & "(See each individual Hard bet for details.)"
        Case Bet.IsHard4
            Title$ = "Hard 4"
            Msg$ = Msg$ & "If a pair of 2's is rolled you win.  If a 3-1 or 1-3 combination is rolled or a 7 come up first you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 8:1"
        Case Bet.IsHard6
            Title$ = "Hard 6"
            Msg$ = Msg$ & "If a pair of 3's is rolled you win.  If a 4-2, 5-1, 1-5, 2-4 combination is rolled or a 7 come up first you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 6:1"
        Case Bet.IsHard8
            Title$ = "Hard 8"
            Msg$ = Msg$ & "If a pair of 4's is rolled you win.  If a 6-2, 5-3, 3-5, 2-6 combination is rolled or a 7 come up first you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 6:1"
        Case Bet.IsHard10
            Title$ = "Hard 10"
            Msg$ = Msg$ & "If a pair of 5's is rolled you win.  If a 4-6 or 6-4 combination is rolled or a 7 come up first you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 8:1"
        Case Bet.IsAny7
            Title$ = "Any 7"
            Msg$ = Msg$ & "If the next roll is a 7 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 4:1"
        Case Bet.IsCraps
            Title$ = "Craps"
            Msg$ = Msg$ & "If the next roll is 2,3 or 12 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 6:1"
        Case Bet.IsHorn
            Title$ = "Cover the Horn"
            Msg$ = Msg$ & "Dropping a bet here causes that amount to be bet on each Horn bet." & Chr$(13)
            Msg$ = Msg$ & "For example: a $5 bet will result in $5 getting placed on the 2, 3, 11 and 12 for a total bet of $20." & Chr$(13)
            Msg$ = Msg$ & "(See each individual Horn bet for details.)"
        Case Bet.IsHorn2
            Title$ = "2"
            Msg$ = Msg$ & "If the next roll is a 2 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 30:1"
        Case Bet.IsHorn3
            Title$ = "3"
            Msg$ = Msg$ & "If the next roll is a 3 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 15:1"
        Case Bet.IsHorn11
            Title$ = "11"
            Msg$ = Msg$ & "If the next roll is a 11 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 15:1"
        Case Bet.IsHorn12
            Title$ = "12"
            Msg$ = Msg$ & "If the next roll is a 12 you win, anything else and you lose." & Chr$(13)
            Msg$ = Msg$ & "Payoff: 30:1"
    End Select
    MsgBox Msg$, vbOKOnly, Title$
End Sub


