VERSION 5.00
Begin VB.Form Main_Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crimson Cubes Assistance"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "Main_Help.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_cont 
      Caption         =   "Continue..."
      Height          =   255
      Left            =   5895
      TabIndex        =   6
      Top             =   5370
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   6450
      ScaleHeight     =   75
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   15
      Width           =   240
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5025
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   645
      Width           =   7890
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   7875
      Begin VB.CommandButton Command1 
         Caption         =   "The End"
         Height          =   285
         Index           =   2
         Left            =   5880
         TabIndex        =   3
         Top             =   165
         Width           =   1920
      End
      Begin VB.CommandButton Command1 
         Caption         =   "The Bets"
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   2
         Top             =   180
         Width           =   1920
      End
      Begin VB.CommandButton Command1 
         Caption         =   "The Game"
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   5
         Top             =   165
         Width           =   1920
      End
   End
End
Attribute VB_Name = "Main_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ViewOption As Integer 'Holds the value of the button that the user clicked on
Dim ContClick As Integer 'This is a counter - it gets incremeted when the player clicks on the continue button

Private Sub cmd_cont_Click()
    'Clear the current text
    Text1 = vbNullString
    'Increment the continue clicker
    Inc ContClick%
    'Show the next bit of text for the current viewing selection
    Select Case ViewOption%
        Case 0
            The_Game
        Case 1
            The_Bets
    End Select
    Picture1.SetFocus '(This just gets the focus off of the command button)
End Sub

Private Sub Command1_Click(Index As Integer)
    ContClick% = 0
    ViewOption% = Index
    Text1 = vbNullString
    'Show the first bit of text for the viewing selection
    Select Case Index
        Case 0
            Call The_Game
            Picture1.SetFocus '(This just gets the focus off of the command button)
        Case 1
            Call The_Bets
            Picture1.SetFocus '(This just gets the focus off of the command button)
        Case 2
            Call The_End
    End Select
End Sub

Private Sub Form_Load()
    Call Greeting
End Sub

Sub The_End()
    'Verify that the player really wants to exit the help
    If MsgBox("Are you sure you want to exit Help?", vbYesNo, "Just checking...") = vbYes Then
        Unload Me
    Else
        Picture1.SetFocus '(This just gets the focus off of the command button)
    End If
End Sub

Sub Greeting()
    Dim Msg As String
    'Start the process by explaining the options
    Msg$ = "Welcome to Crimson Cubes - A casino Craps simulation game." & vbCrLf$ & vbCrLf$
    Msg$ = Msg$ & "Please select a button to read the information described." & vbCrLf$ & vbCrLf$
    Msg$ = Msg$ & "The Game: 'Craps' terminology and some general dice information." & vbCrLf$ & vbCrLf$
    Msg$ = Msg$ & "The Bets: A detailed description on the bets and what they pay." & vbCrLf$ & vbCrLf$
    Msg$ = Msg$ & "The End: Close the help screen."
    Text1 = Msg$
End Sub

Sub The_Game()
    Dim Msg As String
    'This is a basic overview of the game of Craps
    'As the player clicks the continue button, the contclick% variable get incremented which determines which piece of text gets displayed
    Select Case ContClick%
        Case 0
            Msg$ = "The game of Craps is played using a common pair of 6 sided dice." & vbCrLf$
            Msg$ = Msg$ & "This allows for numbers to be rolled from 2 to 12 and the bets are based on how many " & vbCrLf$
            Msg$ = Msg$ & "ways there are to roll specific numbers.  The following is a table of possible rolls:" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & " 2: 1-1" & vbCrLf$
            Msg$ = Msg$ & " 3: 1-2, 2-1" & vbCrLf$
            Msg$ = Msg$ & " 4: 1-3, 2-2, 3-1" & vbCrLf$
            Msg$ = Msg$ & " 5: 1-4, 2-3, 3-2, 4-1" & vbCrLf$
            Msg$ = Msg$ & " 6: 1-5, 2-4, 3-3, 4-2, 5-1" & vbCrLf$
            Msg$ = Msg$ & " 7: 1-6, 2-5, 3-4, 4-3, 5-2, 6-1" & vbCrLf$
            Msg$ = Msg$ & " 8: 2-6, 3-5, 4-4, 5-3, 6-2" & vbCrLf$
            Msg$ = Msg$ & " 9: 3-6, 4-5, 5-4, 6-3" & vbCrLf$
            Msg$ = Msg$ & "10: 4-6, 5-5, 6-4" & vbCrLf$
            Msg$ = Msg$ & "11: 5-6, 6-5" & vbCrLf$
            Msg$ = Msg$ & "12: 6-6" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "The concept to get here is how many different ways there are to roll any given number "
            Msg$ = Msg$ & "against how many ways there are to roll a seven.  This is what Craps is based on."
            'Turn on the continue button
            cmd_cont.Visible = True
        Case 1
            Msg$ = "Here is a list of some terms you should know:" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Shooter: The person currently rolling the dice is the shooter.  In this game, you are "
            Msg$ = Msg$ & "always the shooter. & vbcrlf$ & vbcrlf$"
            Msg$ = Msg$ & "Come out roll: This is basically the first roll of a series.  Some bets can only be placed "
            Msg$ = Msg$ & "on a come out roll while can only be placed if it is not.  The easiest way to tell if  you "
            Msg$ = Msg$ & "are on a come out roll brings us to the next term..." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Button:  The button is big round marker that is used to mark the 'point' (the number "
            Msg$ = Msg$ & "the shooter is trying to roll - more on this later).  If you are on a come out roll, the "
            Msg$ = Msg$ & "button is not visible.   When you are not on a come out roll, the button will be placed "
            Msg$ = Msg$ & "on either the 4,5,6,8,9 or 10"
        Case 2
            Msg$ = "Point:  This is what basic craps is all about. There is more "
            Msg$ = Msg$ & "information about this when you get to the bets but here is the general idea.  "
            Msg$ = Msg$ & "When you look at the layout you will see six large numbers - 4,5,6,8,9 and 10.  "
            Msg$ = Msg$ & "The term for these, oddly enough, is 'numbers'.  The goal of the shooter is to roll "
            Msg$ = Msg$ & "a 'number' on the come out roll.  If this happens, that number becomes the point and "
            Msg$ = Msg$ & "the button is placed on top of the 'number'.  Now the goal of the shooter "
            Msg$ = Msg$ & "is to roll the point again before rolling a seven.  Other numbers have their own "
            Msg$ = Msg$ & "impact and that will be covered in the bets section." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Making the point and 7 out:  When a shooter succeeds in rolling the point number "
            Msg$ = Msg$ & "before rolling a 7, the shooter has made the point, you are once again on a come "
            Msg$ = Msg$ & "out roll and the whole thing starts over.  If a 7 comes up, the shooter has sevened "
            Msg$ = Msg$ & "out, the table is cleared of all working bets (in general) and the whole thing starts over."
        Case 3
            Msg$ = "There is not much you need to know to play this game.  To make a bet, click on the images "
            Msg$ = Msg$ & "of the gold bars that equal the amount of the bet you want to make.  To make a "
            Msg$ = Msg$ & "$20 bet, for example, you could click on the +$5 bar 4 times, or +$25 once and then "
            Msg$ = Msg$ & "the -$5 once.  When you get the amount you want, you can drag the bar over and drop "
            Msg$ = Msg$ & "it on a bet or just simply click on the bet - your choice.  Once you have placed your bet(s), "
            Msg$ = Msg$ & "roll the dice by clicking on the 'Roll' button.  You will be informed of any wins or losses in "
            Msg$ = Msg$ & "the message box that runs across the top of the layout.  Whenever you try to do something "
            Msg$ = Msg$ & "that is not allowed, you will get a message there also.  The maximum bet is $200, minimum "
            Msg$ = Msg$ & "is $1." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "When you are ready to quit, click the 'Cash Out' button."
            'This is the piece of text for this section so turn on the continue button
            cmd_cont.Visible = False
    End Select
    'Put the piece of text in the text box to be read
    Text1 = Msg$
End Sub

Sub The_Bets()
    Dim Msg As String
    'This section describes the bets in detail
    Select Case ContClick%
        Case 0
            Msg$ = "There are a whole lot of bets that can be placed on a craps table.  This section will "
            Msg$ = Msg$ & "describe them all, starting with the simpler ones and progressing on to the more "
            Msg$ = Msg$ & "complex ones." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "There are some bets that are a one roll shot.  That is; the next roll will either be a "
            Msg$ = Msg$ & "winner or a loser.  Let's start with them." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Field: The field bet is a rather prominent section on table.  The numbers in the field "
            Msg$ = Msg$ & "are 2,3,4,9,10,11 & 12.  If the next roll is one of these numbers, you win.  If the next "
            Msg$ = Msg$ & "roll is a 5,6,7 or 8, you lose.  The Field pays even money except for the 2 which "
            Msg$ = Msg$ & "pays 2:1 and the 12 which pays 3:1." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Any 7: This is one of the center bets.  That is one of the bets that would be in the "
            Msg$ = Msg$ & "center of a table.  On the layout here it is the section of bets on the right side of the "
            Msg$ = Msg$ & "playing field.  To place a center bet bet at a table you basically throw your bet "
            Msg$ = Msg$ & "toward the center of the table and shout out the bet want to make.  The Stick will "
            Msg$ = Msg$ & "take your chips and place them on that bet for you.  Back to the Any 7 bet - if the "
            Msg$ = Msg$ & "next roll is a seven you win. If the next roll is not a seven you lose. A winning Any 7 "
            Msg$ = Msg$ & "bet pays 4:1."
            cmd_cont.Visible = True
        Case 1
            Msg$ = "Craps:  Craps is a term as well as a bet.  2, 3 and 12 are Craps.  If the next roll is Craps, "
            Msg$ = Msg$ & "you win.  Any other number and you lose.  Craps pays 6:1" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "The last of the one shot type bets is the Horn.  The Horn in comprised of the four "
            Msg$ = Msg$ & "following numbers: 2,3,11 and 12.  There are various ways of betting the Horn at a "
            Msg$ = Msg$ & "table.  You can bet on any one number and if that is the next number that is rolled, "
            Msg$ = Msg$ & "you win.  You can bet the Horn, which covers all four numbers at the same time.  This "
            Msg$ = Msg$ & "also cuts the payoff by a factor of four.  You can bet High Horn which covers 11 & 12.  "
            Msg$ = Msg$ & "The winnings get cut in half as they do if you bet Low Horn which covers the 2 & 3.  "
            Msg$ = Msg$ & "This game does allow for individual bets but does not allow for High or Low Horn bets.  "
            Msg$ = Msg$ & "The way the Horn itself is handled is thus:  Place your bet in the title area of the bet.  "
            Msg$ = Msg$ & "This will cause that amount to be bet on all four of the Horn numbers.  In other words, "
            Msg$ = Msg$ & "if you place a $1 chip in the area labled HORN you will end up with $1 on the 2, $1 on "
            Msg$ = Msg$ & "the 3..." & vbCrLf$
            Msg$ = Msg$ & "The Horn bets pay as follows:  3 and 11 pay 15:1 - 2 and 12 pay 30:1"
        Case 2
            Msg$ = "Since we are in the neighborhood, let's look at the last of the center bets." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Hardways: A bet is said to be rolled the hard way when it is rolled as doubles.  "
            Msg$ = Msg$ & "This goes back to how many ways there are to roll any given number again.  If  you "
            Msg$ = Msg$ & "look at four, there are three ways to roll it.  The two easy ways are 1-3 and 3-1 and "
            Msg$ = Msg$ & "then there is the only hard way 2-2.  A Hardway bet can only be won if the selected "
            Msg$ = Msg$ & "number is rolled hard.  If you bet on a Hard 4 and the dice come up 1-3 or 3-1, you "
            Msg$ = Msg$ & "lose.  You also lose if a seven comes up.  Betting on the Hardways is similar to betting "
            Msg$ = Msg$ & "on the Horn.  You can bet on any single Hard number or place your bet in the title "
            Msg$ = Msg$ & "area and all four Hard numbers will be covered by that amount." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "The Hardways pay as follows:  4 and 10 pay 9:1 - 6 and 8 pay 7:1."
        Case 3
            Msg$ = "There are two more simple bets - Big and Place" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Big 6 and Big 8: These are pretty straight foreward.  You bet on one or the other or "
            Msg$ = Msg$ & "even both if you like.  If the number, 6 or 8, comes up before a 7 you win.  If the 7 "
            Msg$ = Msg$ & "is rolled first you lose.  These bets pay 1:1." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Place Bets:  While Place bets have a designated spot on the table, they are usually "
            Msg$ = Msg$ & "not labeled.  You just know that the bet exists and where it goes.  At a table, you "
            Msg$ = Msg$ & "set your chips in front of the dealer and tell him or her to 'Place the 6' or whatever "
            Msg$ = Msg$ & "number you wish to make a Place bet on.  In this game, you drag your chip to the "
            Msg$ = Msg$ & "empty box just above the number (not the empty box on the very top of the play "
            Msg$ = Msg$ & "field - that is for Don't bets which will come soon enough).  The payoffs on Place "
            Msg$ = Msg$ & "bets is 5:7 on 4,5,9, & 10 and 6:7 on 6 & 8 so get the most for your money ("
            Msg$ = Msg$ & "because the casinos don't pay odds cents and they round down) you should bet in "
            Msg$ = Msg$ & "multiples of five for 4,5,9,& 10 and multiples of six on 6 & 8."
        Case 4
            Msg$ = "When you make a Place bet, this also apllies to Come and Don't bets which are coming up, the "
            Msg$ = Msg$ & "dealer takes control of your chips and places them on the layout according to the "
            Msg$ = Msg$ & "type of bet but also according to where you are standing around the table.  When "
            Msg$ = Msg$ & "you make one these bets, pay attention to where on the particular area for the type "
            Msg$ = Msg$ & "of bet and the related number your chips are placed.  In other words, when you "
            Msg$ = Msg$ & "Place the 6, the dealer will take your chips and put them in the Place box behind the "
            Msg$ = Msg$ & "6 - this is where Place bets go.  But within the box, your chips will be put in a specific "
            Msg$ = Msg$ & "spot that relates to your physical position at the table.  This is how the dealers can "
            Msg$ = Msg$ & "keep track of everything.  I like to call this spot your address.  Once you know your "
            Msg$ = Msg$ & "address, you can easily scan the table and see where all your chips are and what bets "
            Msg$ = Msg$ & "you have working."
        Case 5
            Msg$ = "These next few bets are the ones that you will play the most so pay attention!" & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Pass: The Pass bet can only be made on a come out roll and it kind of has two "
            Msg$ = Msg$ & "phases.  The first phase is what happens on the come out roll and the second "
            Msg$ = Msg$ & "phase what happens after the come out roll.  On a come out roll, if the shooter "
            Msg$ = Msg$ & "rolls Craps (2,3 or 12) you lose your bet.  If the shooter rolls a 7 or 11 you win "
            Msg$ = Msg$ & "you bet.  If the shooter rolls a number (4,5,6,8,9 or 10) then that number becomes "
            Msg$ = Msg$ & "the Point. sound familiar?  Once a Point has been established, phase two kicks in.  "
            Msg$ = Msg$ & "Now your bet will ride until the Point is made or the shooter sevens out.  The Pass "
            Msg$ = Msg$ & "bet pays 1:1.  Now for the important and most complicated part.  The Odds bet.  "
            Msg$ = Msg$ & "The Odds bet is not marked in any way, shape or form.  You just have to "
            Msg$ = Msg$ & "know it can be made - and now you do.  Here is how it works.  Somewhere up on "
            Msg$ = Msg$ & "the inside wall of the table is a little sign that tells you what the table limits are.  "
            Msg$ = Msg$ & "It will state the minimum and maximum bet allowed and it will also state the amount "
            Msg$ = Msg$ & "of Odds that are allowed.  Some only allow single Odds, most allow double and I've "
            Msg$ = Msg$ & "seen tables with a minimum Odds bet of 3X and a maximum of 5X.  My preference is "
            Msg$ = Msg$ & "a table that allows double, or 2X, Odds.  From now on, all references made to Odds "
            Msg$ = Msg$ & "will assume double."
        Case 6
            Msg$ = "Here is what you do to 'take odds' on a bet.  First of all, you have to have a Pass bet made "
            Msg$ = Msg$ & "to take odds on.  Once a Point is established, you can take odds on that Pass bet "
            Msg$ = Msg$ & "and the amount of that Odds bet can be up to twice the original Pass bet.  If you "
            Msg$ = Msg$ & "were at a table, you would simply place your chips behind your Pass bet.  In this "
            Msg$ = Msg$ & "game you drop your bet anywhere in the Pass area and the program will place it "
            Msg$ = Msg$ & "on top of your Pass bet.  The Odds bet is called Odds because it pay according to "
            Msg$ = Msg$ & "the actual odds of the number being rolled versus a seven.  Let's say that the Point "
            Msg$ = Msg$ & "is 10.  You had bet $5 on Pass.  Now you can take a $10 Odds bet on your $5 Pass "
            Msg$ = Msg$ & "bet.  If the shooter rolls a 10, you will your regular Pass bet at even money.  The "
            Msg$ = Msg$ & "Odds bet would pay at 2:1 and here's why.  Going back to how many ways to roll a "
            Msg$ = Msg$ & "number once again; there are six ways to roll 7 and three ways to roll a 10 so the "
            Msg$ = Msg$ & "Odds are 6:3 or 2:1.  Here is the full table: 4 and 10 pay 2:1, 5 and 9 pay 3:2, "
            Msg$ = Msg$ & "6 and 8 pay 6:5.  6 and 8 have a special status.  Because the pay is 6:5 you can "
            Msg$ = Msg$ & "round tup he Odds bet to the nearest multiple of 5 - somewhat similar to Place bets."
        Case 7
            Msg$ = "The Pass bet is the table control bet.  It determines when the dice get passed to the next "
            Msg$ = Msg$ & "shooter.  Since it can only be made on a come out roll, everybody who bet on it "
            Msg$ = Msg$ & "have their chips riding on the same number.  The Come bet is like a private Pass bet.  " & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Come:  You can make a Come bet anytime that it is not a come out roll.  You can also "
            Msg$ = Msg$ & "make as many Come bets as you like.  The winning and losing of a Come bet is just like "
            Msg$ = Msg$ & "a Pass bet but instead of working for the whole table it only works for you.  When it is "
            Msg$ = Msg$ & "not a come out roll, drop a bet in the Come area.  If the next roll is 7 or 11 you win.  "
            Msg$ = Msg$ & "Of course you would end up losing you Pass bet if a 7 is rolled.  2,3 and 12 means you "
            Msg$ = Msg$ & "lose your Come bet - no effect your Pass bet.  If a number is rolled, your bet is moved "
            Msg$ = Msg$ & "from the Come area to the actual number on the play field and that becomes a "
            Msg$ = Msg$ & "personal point number.. When your at a table, your chips are located on the number "
            Msg$ = Msg$ & "according to your position at the table.  Pay attention to where the dealer places "
            Msg$ = Msg$ & "your chips because that is your address."
        Case 8
            Msg$ = "Just like the Pass bet, your bet will now hang out and wait for the shooter to roll that number "
            Msg$ = Msg$ & "again or a 7.  And just like the Pass bet, you can take odds on a Come bet.  At a table, "
            Msg$ = Msg$ & "you would hand the chips to the dealer and say something like 'Odds on the 8' and "
            Msg$ = Msg$ & "your Odds bet would be set on top of your Come bet but offset a bit to distinguish it "
            Msg$ = Msg$ & "as an Odds bet.  To take odds on a Come bet in this game, drop the chip in the number "
            Msg$ = Msg$ & "that has the Come bet riding on it.  The program will place it on top of the come bet - "
            Msg$ = Msg$ & "offset just a bit to simulate how the casino would handle it."
        Case 9
            Msg$ = "Now we are going to take everything you know about Pass and Come bets and flip it 180 degrees." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "Don't:  The Don't bet works for both Don't Pass and Don't Come bets.  These work just "
            Msg$ = Msg$ & "like Pass and Come except that the winning and losing is reversed.  The Pass and Don't "
            Msg$ = Msg$ & "Pass are called line bets.  You must have made a line bet to be eligible to roll the dice.  "
            Msg$ = Msg$ & "These two line bets have a couple of terms associated with them.  The more popular is "
            Msg$ = Msg$ & "called betting with or against the dice.  The other reference is betting the right way and "
            Msg$ = Msg$ & "betting the wrong way.  Note that right and wrong are not moral judgements.  Don't Pass "
            Msg$ = Msg$ & "and Don't Come are a little more difficult to distinguish because they use the same area "
            Msg$ = Msg$ & "of the play field.  Unlkie the Pass and Come bets which are very obviously separarted.  "
            Msg$ = Msg$ & "Aside from being considered a line bet which gets you the dice if you want them (Come "
            Msg$ = Msg$ & "bets are not line bets because you can not make a Come bet on the come out roll) "
            Msg$ = Msg$ & "the Don't can really be thought of as Don't Come at all times.  From here on out, I will "
            Msg$ = Msg$ & "just be using 'Don't' to refer to it."
        Case 10
            Msg$ = "Don't bets can be made at anytime.  If the next roll is a 7 or 11, you lose.  If it is a 2 or a 3, you "
            Msg$ = Msg$ & "win - 12 has no effect.  If the next roll is a number, your bet will be moved from the Don't "
            Msg$ = Msg$ & "area to the box at the very top of the play field above that number.  Now what you want "
            Msg$ = Msg$ & "to happen is to a seven rolled before that number.  Since the  Don't works as a Don't Come, "
            Msg$ = Msg$ & "you can have more than one Don't bet at any given time.  If any one of those numbers "
            Msg$ = Msg$ & "is rolled, you would lose that number.  If a 7 is rolled, you would win all the Don't bets "
            Msg$ = Msg$ & "that you have working at that time.  Just like Pass and Come, you have an odds option "
            Msg$ = Msg$ & "with Don't bets.  This is called laying odds against a number and the numbers get a little "
            Msg$ = Msg$ & "hairy so brace yourself."
        Case 11
            Msg$ = "To lay odds against a number, the process is the same as for Come, give your chips to the "
            Msg$ = Msg$ & "dealer and say 'Odds on the 6' and your chips will be put on the Don't 6 bet, offset.  "
            Msg$ = Msg$ & "The big difference is how much you can bet.  Since there are more ways to roll a 7 "
            Msg$ = Msg$ & "than any other number, the chances of a 7 coming up before a point number a greater "
            Msg$ = Msg$ & "and therefore Don't bets are a pretty good deal.  But, the casino knows this and "
            Msg$ = Msg$ & "compensates by making you risk more money.  Let's use 4 and 10 for an example "
            Msg$ = Msg$ & "because the numbers are easier.  For Pass and Come the payoff on the Odds are 2:1.  "
            Msg$ = Msg$ & "On a Don't Odds bet, the payoff is 1:2.  In other words, you get paid $1 for every $2 "
            Msg$ = Msg$ & "that you bet.  Since we are dealing with double Odds, you are allowed to win twice "
            Msg$ = Msg$ & "your original Don't bet amount with your Odds bet.  Let's say you make a $5 Don't bet "
            Msg$ = Msg$ & "and the shooter rolls a 4.  You are allowed to win $10 on an Odds bet against your "
            Msg$ = Msg$ & "original $5 Don't bet. Since the payoff is 1:2 your maximum Odds bet would be $20.  "
            Msg$ = Msg$ & "Sticking with an original bet of $5, on a 5 or 9 the payoff is 2:3 which results in $15 and "
            Msg$ = Msg$ & "for 6 and 8 at 6:5 we have $12."
        Case 12
            Msg$ = "At this point, most people go that's nuts.  Why would I want to bet more than I would win?  "
            Msg$ = Msg$ & "One reason was already touched on - there are more ways to roll a seven than any "
            Msg$ = Msg$ & "other number.  The other thing to consider is the net result of winning with the dice "
            Msg$ = Msg$ & "versus winning against the dice is the same.  Here is an example: you make a $5 Pass "
            Msg$ = Msg$ & "bet and the shooter rolls a 4.  You take $10 Odds on it and you win.  You get paid "
            Msg$ = Msg$ & "even money for the Pass bet and 2:1 for the Odds bet plus your original bet is "
            Msg$ = Msg$ & "returned to you giving the following results: $5 Pass returned + $5 winnings + $10 Odds "
            Msg$ = Msg$ & "returned + $20 winnings = total money in the bank of $40.  Let's make a $5 Don't bet "
            Msg$ = Msg$ & "this time:  shooter rolls a 4 and you lay $20 Odds against it.  This time the shooter "
            Msg$ = Msg$ & "rolls a 7 and you win: $5 Don't bet returned + $5 winnings + $20 Odds returned + "
            Msg$ = Msg$ & "$10 winnings = total money in the bank of $40." & vbCrLf$ & vbCrLf$
            Msg$ = Msg$ & "To get a brief description of a  bet during play, hit the question mark and then click "
            Msg$ = Msg$ & "on the bet in the play field.  To get out of 'What's This' mode, hit Esc. or place a bet."
        cmd_cont.Visible = False
    End Select
    Text1 = Msg$
End Sub
