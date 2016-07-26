VERSION 5.00
Begin VB.Form frmCard 
   Caption         =   "Poker Hand Simulator"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPercentages 
      Caption         =   "Percentages"
      Height          =   375
      Left            =   6960
      TabIndex        =   38
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CheckBox chkDisplay 
      Caption         =   "Display"
      Height          =   375
      Left            =   6960
      TabIndex        =   36
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   6960
      TabIndex        =   35
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox txtHandCount 
      Height          =   375
      Left            =   4440
      TabIndex        =   34
      Text            =   "0"
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   10
      Left            =   4800
      TabIndex        =   32
      Text            =   "0"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   9
      Left            =   4800
      TabIndex        =   31
      Text            =   "0"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   8
      Left            =   4800
      TabIndex        =   30
      Text            =   "0"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   29
      Text            =   "0"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   28
      Text            =   "0"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   27
      Text            =   "0"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   26
      Text            =   "0"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   25
      Text            =   "0"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   24
      Text            =   "0"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtHandPercentage 
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   23
      Text            =   "0"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   10
      Left            =   2880
      TabIndex        =   21
      Text            =   "0"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   9
      Left            =   2880
      TabIndex        =   19
      Text            =   "0"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   8
      Left            =   2880
      TabIndex        =   17
      Text            =   "0"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   7
      Left            =   2880
      TabIndex        =   15
      Text            =   "0"
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   13
      Text            =   "0"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      Text            =   "0"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   9
      Text            =   "0"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   7
      Text            =   "0"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      Text            =   "0"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtHand 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Text            =   "0"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   855
      Left            =   6840
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblHand 
      Caption         =   "Hand Type"
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblHandCount 
      Caption         =   "Hand Count:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label lblRoyalFlush 
      Caption         =   "Royal Flush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   22
      Top             =   7000
      Width           =   1215
   End
   Begin VB.Label lblStraightFlush 
      Caption         =   "Straight Flush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      Top             =   6520
      Width           =   1215
   End
   Begin VB.Label lblFourOfAKind 
      Caption         =   "Four of A Kind"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   6040
      Width           =   1215
   End
   Begin VB.Label lblFullHouse 
      Caption         =   "Full House"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   5560
      Width           =   1215
   End
   Begin VB.Label lblFlush 
      Caption         =   "Flush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   5085
      Width           =   1215
   End
   Begin VB.Label lblStraight 
      Caption         =   "Straight"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Label lblThreeOfAKind 
      Caption         =   "Three of A Kind"
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   4120
      Width           =   1215
   End
   Begin VB.Label lblTwoPair 
      Caption         =   "Two Pair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Label lblOnePair 
      Caption         =   "One Pair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3165
      Width           =   1215
   End
   Begin VB.Label lblPercentage 
      Caption         =   "Percentage (%)"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Caption         =   "Count"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblHighCard 
      Caption         =   "High Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2685
      Width           =   1215
   End
   Begin VB.Image imgCard 
      Height          =   1335
      Index           =   4
      Left            =   6840
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgCard 
      Height          =   1335
      Index           =   3
      Left            =   5400
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgCard 
      Height          =   1335
      Index           =   2
      Left            =   3960
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgCard 
      Height          =   1335
      Index           =   1
      Left            =   2520
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image imgCard 
      Height          =   1335
      Index           =   0
      Left            =   1080
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Poker Hand Simulator
'Luke Scollard
'ICS 3U1
'06/09/14

'Software Definition:
'The purpose of this program is to calculate the likelihood of various hands appearing in poker. It does this by shuffling
'a virtual deck of cards, then taking the first five cards in the deck, and checking which poker hand they correspond to.
'The program also creates a display showing the cards in the hand, and outputs the count of the various hands, as well as
'the percentage of that hand's occurence.

'Design Decisions:
'A design decision is to allow the program to run until the long integer variable experiences overflow, which is a very high
'value, and thus allows the most amount of hands to be dealt and analyzed. This also makes the percentages given more accurate,
'as there are more hands dealt, and thus causing any anomalies, such as more Royal Flushes appearing than normal, to be lost as
'odds even themselves out to their proper probabilities, as the anomalies are lost due to the huge amount of consistant data.

'A design decision is to show a visual display of the cards being drawn. The reason for this visual display is both to show
'the user what cards are being analyzed, and to fulfill the required visual component of the project. The display also shows the user
'user what hand the program identifies the five cards as create.

'A design decision of this project is to allow the user to turn off the visual display of the cards. This allows for the user
'to choose whether they want the program to run faster without the display, or slower, but with the display. The user also has
'the ability to stop the percentages from being calculated, also for the purpose of speeding up the program. If the percentages
'are turned off, the percentages will be calculated on every millionth hand, so the user can still see them.

'A design decision is to allow the user to pause the program in order to see the visual display of the hand. This allows the
'user to compare the visual display of the hand to the hand the program identifies the cards as, giving them the chance to check
'that the programs algorithms are accurate.


'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'Deck                   General        clsCards        This variable allows the program to access the Cards32.dll functions
'Card()                 General        Integer         This array contains the deck of cards, with each of the fifty two variables
'                                                      containing a number that represents one card.
'Hands()                General        Long            This array records the number of times certain hands shows up. There are
'                                                      ten variables in the array, one for each hand.




Dim Deck As clsCards
Dim Card(0 To 51) As Integer
Dim Hands(1 To 10) As Long


Sub Build_Deck()

'Purpose: The purpose of this subroutine is to build the deck.

'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'x                      Build_Deck     Integer         This variable acts as a counter in the For Loop which assigns the card values
'                                                      to the Card Array.

    Dim x As Integer
    Set Deck = New clsCards
    
    For x = 0 To 51
        Card(x) = x
    Next
End Sub

Sub Shuffle_Deck()

'Purpose: The purpose of this subroutine is to shuffle the deck into a random order.

'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'x                      Shuffle_Deck   Integer         This variable acts as a counter in the For Loop which randomizes the order of
'                                                      the cards.
'y                      Shuffle_Deck   Integer         This variable acts as a counter in the For Loop which randomizes the order of
'                                                      the cards.
'Swap                   Shuffle_Deck   Integer         This variable is used to hold the value so that the two card values can be
'                                                      swapped.

    Dim x As Integer
    Dim y As Integer
    Dim Swap As Integer
    
    Randomize
    For y = 0 To 51
        For x = 0 To 51
            If Rnd(1) <= 0.5 Then
                Swap = Card(y)
                Card(y) = Card(x)
                Card(x) = Swap
            End If
        Next x
    Next y
End Sub

Private Sub cmdStart_Click()

'Purpose: The purpose of this subroutine is to create the visual display for the cards, and to calculate the percentages of each hand
'showing up.

'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'x                      cmdStart       Integer         This variable acts as a counter in the For Loop that sets the picture for the
'                                                      cards' visual display. It also makes the display disappear if the user turns the
'                                                      display off.
'Count                  cmdStart       Long            This variable holds the value of the number of hands dealt.

    Dim x As Integer
    Dim Count As Long
        
    Count = 0
    
    Do
        Call Shuffle_Deck
        
        If chkDisplay.Value = 0 Then
            For x = 0 To 4
                imgCard(x).Visible = False
            Next x
            lblHand.Visible = False
        End If
        
        If chkDisplay.Value = 1 Then
            For x = 0 To 4
                imgCard(x).Visible = True
                imgCard(x).Picture = Deck.SetCardImage(Card(x))
            Next x
            lblHand.Visible = True
        End If
        
        DoEvents
        
        Call Find_Hand
        
        Do While chkPause.Value = 1
            DoEvents
        Loop
        
        Count = Count + 1
        txtHandCount.Text = Count
        
        If chkPercentages.Value = 1 Or Count / 1000000 = Int(Count / 1000000) Then
            For x = 1 To 10
                txtHandPercentage(x).Text = Int(((txtHand(x).Text / Count) * 100) * 1000000) / 1000000
            Next x
        End If
    Loop
    
End Sub

Public Sub Form_Load()

'Purpose: This subroutine causes the form to load, then calls the deck to be built, before setting all the hand values to 0.

'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'x                      Form_Load()    Integer         This variable is used as a counter in the For Loop that sets all the hands'
'                                                      values to 0.

    Dim x As Integer
    
    Call Build_Deck
    
    For x = 1 To 10
        Hands(x) = 0
    Next x
End Sub

Private Sub Find_Hand()

'Purpose: This subroutine applies various algorithms to check to see which hand the cards create. It finds how many pairs there are in
'the hand, if it is a straight and if it is a flush, before updating the count for the hand the cards fit into.

'Variable Dictionary:

'Variable Name          Scope          Type            Purpose
'x                      Find_Hand()    Integer         This variable is used as a counter in several For Loops.
'y                      Find_Hand()    Integer         This variable is used as a counter in the For Loop that checks for pairs.
'N1                     Find_Hand()    Integer         This variable is used to switch the two numbers in the Bubble Sort that sorts
'                                                      the card values so the program can check if the hand is a straight.
'Switched               Find_Hand()    Boolean         This variable causes the Bubble Sort to end if none of the numbers are switched
'                                                      in the current run, meaning it is fully sorted.
'Card_Divide()          Find_Hand()    Single          This variable contains the number of the card divided by 13. It is used to find
'                                                      which suit the card belongs to.
'Card_Suit()            Find_Hand()    String          This variable contains the suit of the cards that is being checked. The variable
'                                                      is used to find whether the hand is a flush.
'Card_Value()           Find_Hand()    Integer         This variable contains the value of the card, without taking the suit into account.
'                                                      The value will be between 1 (Ace) and 13 (King).
'Counter                Find_Hand()    Integer         This variable counts the number of pairs there is in the hand. The value is then
'                                                      to calculate whether the hand is One Pair, Two Pair, Three of a Kind, Full House,
'                                                      or Four of a Kind.
'Hand                   Find_Hand()    Integer         This variable contains the value, from 1 to 10, of the hand the cards create. It is
'                                                      used to update the count, by allowing the program to update the exact count array for
'                                                      the hand.
'One_Pair               Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a One Pair.
'Two_Pair               Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Two Pair.
'Three_of_a_Kind        Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Three of a Kind.
'Straight               Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Straight.
'Flush                  Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Flush.
'Full_House             Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Full House.
'Four_of_a_Kind         Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Four of a Kind.
'Straight_Flush         Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Straight Flush.
'Royal_Straight         Find_Hand()    Boolean         This boolean is used to tell the program that cards have a straight which could become
'                                                      a Royal Flush, if the hands are a flush.
'Royal_Flush            Find_Hand()    Boolean         This boolean is used to tell the program that the hand the cards create is a Royal Flush.


    Dim x As Integer
    Dim y As Integer
    Dim N1 As Integer
    Dim Switched As Boolean
    
    
    Dim Card_Divide(0 To 4) As Single
    Dim Card_Suit(0 To 4) As String
    Dim Card_Value(0 To 4) As Integer
    Dim Counter As Integer
    Dim Hand As Integer
    
    Dim One_Pair As Boolean
    Dim Two_Pair As Boolean
    Dim Three_of_a_Kind As Boolean
    Dim Straight As Boolean
    Dim Flush As Boolean
    Dim Full_House As Boolean
    Dim Four_of_a_Kind As Boolean
    Dim Straight_Flush As Boolean
    Dim Royal_Straight As Boolean
    Dim Royal_Flush As Boolean
    
    Flush = False
    Switched = True
    
    
    For x = 0 To 4
        Card_Divide(x) = Card(x) / 13
        
        'Finds the suits of the individual cards
        
        If Card_Divide(x) >= 0 And Card_Divide(x) < 1 Then
            Card_Suit(x) = "S"
        ElseIf Card_Divide(x) >= 1 And Card_Divide(x) < 2 Then
            Card_Suit(x) = "C"
        ElseIf Card_Divide(x) >= 2 And Card_Divide(x) < 3 Then
            Card_Suit(x) = "D"
        ElseIf Card_Divide(x) >= 3 And Card_Divide(x) < 4 Then
            Card_Suit(x) = "H"
        End If
        
        'Finds the value of the individual cards
        
        Card_Value(x) = (Card(x) Mod 13) + 1
    Next x
    
    Counter = 0
    
    'Finds how many matching cards there are
    
    For x = 0 To 4
        For y = x To 4
            If Card_Value(x) = Card_Value(y) And x <> y Then
                Counter = Counter + 1
            End If
        Next y
    Next x
    
    'Finds which hand the cards create due to the number of matching cards
    
    If Counter = 1 Then
        One_Pair = True
    ElseIf Counter = 2 Then
        Two_Pair = True
    ElseIf Counter = 3 Then
        Three_of_a_Kind = True
    ElseIf Counter = 4 Then
        Full_House = True
    ElseIf Counter = 6 Then
        Four_of_a_Kind = True
    End If
    
    If Counter = 0 Then
        Straight = True
        Flush = True
        
        'Finds if all the suits are the same
        
        For x = 0 To 4
            If Card_Suit(0) <> Card_Suit(x) Then
                Flush = False
            End If
        Next x
        
        'Sorts the cards into lowest to highest order
        
        Do Until Switched = False
            Switched = False
            For x = 0 To 3
                If Card_Value(4 - x) < Card_Value(4 - x - 1) Then
                    N1 = Card_Value(4 - x)
                    Card_Value(4 - x) = Card_Value(4 - x - 1)
                    Card_Value(4 - x - 1) = N1
                    Switched = True
                End If
            Next x
        Loop
        
        'Checks to see if cards create a straight
        
        For x = 0 To 4
            If Card_Value(0) + x <> Card_Value(x) Then
                Straight = False
            End If
        Next x
        
        'Checks to see if cards create a royal flush
        
        If Card_Value(0) = 1 And Card_Value(1) = 10 And Card_Value(2) = 11 And Card_Value(3) = 12 And Card_Value(4) = 13 Then
            Royal_Straight = True
        End If
    End If
    
    If Straight = True And Flush = True Then
        Straight_Flush = True
    End If
    
    If Royal_Straight = True And Flush = True Then
        Royal_Flush = True
    ElseIf Royal_Straight = True Then
        Straight = True
    End If
    
    If Royal_Flush = True Then
        Hand = 10
        lblHand.Caption = "Royal Flush"
    ElseIf Straight_Flush = True Then
        Hand = 9
        lblHand.Caption = "Straight Flush"
    ElseIf Four_of_a_Kind = True Then
        Hand = 8
        lblHand.Caption = "Four of a Kind"
    ElseIf Full_House = True Then
        Hand = 7
        lblHand.Caption = "Full House"
    ElseIf Flush = True Then
        Hand = 6
        lblHand.Caption = "Flush"
    ElseIf Straight = True Then
        Hand = 5
        lblHand.Caption = "Straight"
    ElseIf Three_of_a_Kind = True Then
        Hand = 4
        lblHand.Caption = "Three of a Kind"
    ElseIf Two_Pair = True Then
        Hand = 3
        lblHand.Caption = "Two Pair"
    ElseIf One_Pair = True Then
        Hand = 2
        lblHand.Caption = "One Pair"
    Else
        Hand = 1
        lblHand.Caption = "High Card"
    End If
    
    Hands(Hand) = Hands(Hand) + 1
    
    txtHand(Hand).Text = Hands(Hand)
         
End Sub

