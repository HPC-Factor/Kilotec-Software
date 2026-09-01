Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function sndPlaySound Lib "coredll.dll" Alias "sndPlaySoundW" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

Public BANK_INIT As Integer
Public BANK_MAX As Long
Public BET_MIN As Integer
Public BET_MAX As Integer
Public BET_INC As Integer

Const MODE_NOINPUT = 0
Const MODE_GETBET = 1
Const MODE_HITORSTAND = 2
Const MODE_STARTFIRST = 3
Const MODE_STARTOVER = 4
Const MODE_GAMEOVER = 5

Const DEALER_FIRST = 0
Const PLAYER_FIRST = 1
Const DEALER_SECOND = 2
Const PLAYER_SECOND = 3
Const PLAYER_NEXT = 4
Const DEALER_NEXT = 5

Const TOTAL_DECK_CARDS = 52
Const TOTAL_HAND_CARDS = 7

Public deck(52)
Public player(52)
Public dealer(52)

Public bet As Integer
Public bank As Integer

Public nextDealerCard As Integer
Public nextPlayerCard As Integer
Public nextCardToDraw As Integer
Public hideDealerCard As Boolean

Public playMode As String
Public dealMode As String

Public i As Integer
Public x As Integer
Public cardRandom As Integer
Public cardShuffle As Integer
Public t As Integer
Public card As Integer

Public dealercount As Integer
Public playercount As Integer

Public GamesPlayed As Integer
Public CardsDealt As Integer
Public BlackJacks As Integer
Public Won As Integer
Public Lost As Integer
Public Draw As Integer

Public Sound As Boolean




Public Sub ShuffleDeck()


Randomize Timer

For cardShuffle = 0 To TOTAL_DECK_CARDS

    cardRandom = Int(Rnd() * TOTAL_DECK_CARDS)
    t = deck(cardShuffle)
    
    deck(cardShuffle) = deck(cardRandom)
    deck(cardRandom) = t
    
Next cardShuffle

nextCardToDraw = 0
playMode = MODE_GETBET


End Sub

Public Sub DealCard()

PlayWaveFile (App.Path & "\card.wav")
card = deck(nextCardToDraw)
nextCardToDraw = nextCardToDraw + 1
CardsDealt = CardsDealt + 1

If nextCardToDraw >= TOTAL_DECK_CARDS Then ShuffleDeck


    If dealMode = PLAYER_NEXT Then
        processPlayerHit card
    ElseIf dealMode = DEALER_NEXT Then
        processDealerHit card
    Else
        dealInitialCards card
    End If

End Sub

Public Sub processPlayerHit(ByVal card As Integer)
player(nextPlayerCard) = card
nextPlayerCard = nextPlayerCard + 1
CardsDealt = CardsDealt + 1
DisplayHands

Dim playercount As Integer
Dim dealercount As Integer

playercount = countHand(player, nextPlayerCard)
If playercount > 21 Then
    hideDealerCard = False
    DisplayHands
    dealercount = countHand(dealer, nextDealerCard)
    MsgBox "BUST!! You loose!", vbOKOnly
    bank = bank - bet
    Lost = Lost + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
End If
'displayhand(player.nextplayercard,0,1)


End Sub
Public Sub processDealerHit(ByVal card As Integer)
dealer(nextDealerCard) = card
nextDealerCard = nextDealerCard + 1
CardsDealt = CardsDealt + 1
DisplayHands
dealerHit
End Sub


Public Sub dealInitialCards(ByVal card As Integer)
If dealMode = DEALER_FIRST Then hideDealerCard = True

If Int(dealMode And 1) = 0 Then
    dealer(nextDealerCard) = card
    nextDealerCard = nextDealerCard + 1
Else
    player(nextPlayerCard) = card
    nextPlayerCard = nextPlayerCard + 1
End If

DisplayHands
dealMode = dealMode + 1
If dealMode = PLAYER_NEXT Then
    processTestCards
Else
    DealCard
End If


Dim score As Integer
Dim aces As Integer

For i = 1 To 1
    x = dealer(i) Mod 13
    If x = 0 Then aces = aces + 1
    If x = 0 Then
        score = score + 11
    ElseIf x > 9 Then
        score = score + 10
    Else
        score = score + x + 1
    End If
Next i
    
    
    If score > 21 Then 'And aces > 0 Then
        Do While aces > 0
            score = score - 10
            aces = aces - 1
        Loop
    End If
    
    dealercount = score

playercount = countHand(player, nextPlayerCard)

frmMain.Label3.Caption = dealercount
frmMain.Label2.Caption = playercount
End Sub

Public Sub DisplayHands()
    displayHand dealer, nextDealerCard, hideDealerCard, 0
    displayHand player, nextPlayerCard, 0, 1
End Sub

Public Function displayHand(ByVal hand As Integer, ByVal nextCard As Integer, ByVal hidefirstcard As Boolean, ByVal who As Integer)
Dim Zeichen As String
Dim Karte As Integer

For card = 0 To nextCard - 1
    If hidefirstcard = True And card = 0 Then
        frmMain.Dealer1.Visible = True
        frmMain.Dealer1.Picture = App.Path & "\back.bmp"
    Else
        Dim suit As Integer
        Dim face As Integer
        Dim Farbe As String
        Dim Color As String
        
        suit = Int(hand(card) / 13)
        face = Int(hand(card) Mod 13)
           
        Select Case face
        
            Case 0
            Karte = 1
          
            Case 9
            Karte = 10
            
            Case 10
            Karte = 11
            
            Case 11
            Karte = 12
            
            Case 12
            Karte = 13
            
            Case Else
            Karte = face + 1
           
        
        End Select
    End If
    
    
    If suit = 0 Then Color = "h"
    If suit = 1 Then Color = "d"
    If suit = 2 Then Color = "s"
    If suit = 3 Then Color = "c"
            

    If who = 0 And card = 0 Then frmMain.Label5.Caption = Color & Karte
    'If who = 0 And card = 0 Then frmMain.Dealer1.Visible = True: frmMain.Dealer1.Picture = App.Path & "\" & Color & Karte & ".bmp" 'frmMain.Label5.Caption = Color & Karte
    If who = 0 And card = 1 Then frmMain.Dealer2.Visible = True: frmMain.Dealer2.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 0 And card = 2 Then frmMain.Dealer3.Visible = True: frmMain.Dealer3.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 0 And card = 3 Then frmMain.Dealer4.Visible = True: frmMain.Dealer4.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 0 And card = 4 Then frmMain.Dealer5.Visible = True: frmMain.Dealer5.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 0 And card = 5 Then frmMain.Dealer6.Visible = True: frmMain.Dealer6.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 0 And card = 6 Then frmMain.Dealer7.Visible = True: frmMain.Dealer7.Picture = App.Path & "\" & Color & Karte & ".bmp"
    
    If who = 1 And card = 0 Then frmMain.Player1.Visible = True: frmMain.Player1.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 1 Then frmMain.Player2.Visible = True: frmMain.Player2.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 2 Then frmMain.Player3.Visible = True: frmMain.Player3.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 3 Then frmMain.Player4.Visible = True: frmMain.Player4.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 4 Then frmMain.Player5.Visible = True: frmMain.Player5.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 5 Then frmMain.Player6.Visible = True: frmMain.Player6.Picture = App.Path & "\" & Color & Karte & ".bmp"
    If who = 1 And card = 6 Then frmMain.Player7.Visible = True: frmMain.Player7.Picture = App.Path & "\" & Color & Karte & ".bmp"
 





Next card
End Function


Public Sub processTestCards()
If nextDealerCard = 2 And countHand(dealer, nextDealerCard) = 21 Then
    hideDealerCard = False
    DisplayHands
    dealerBlackjack
ElseIf nextPlayerCard = 2 And countHand(player, nextPlayerCard) = 21 Then
    hideDealerCard = False
    DisplayHands
    PlayerBlackjack
Else
    playMode = MODE_HITORSTAND
End If



End Sub

Public Function countHand(ByVal hand As Integer, ByVal nextCard As Integer)
Dim score As Integer
Dim aces As Integer

For i = 0 To nextCard - 1
    x = hand(i) Mod 13
    If x = 0 Then aces = aces + 1
    If x = 0 Then
        score = score + 11
    ElseIf x > 9 Then
        score = score + 10
    Else
        score = score + x + 1
    End If
Next i
    
    
    If score > 21 Then 'And aces > 0 Then
        Do While aces > 0
            score = score - 10
            aces = aces - 1
        Loop
    End If
    
    countHand = score
End Function

Public Sub dealerHit()
Dim Start As Single
Start = Timer

dealercount = countHand(dealer, nextDealerCard)
playercount = countHand(player, nextPlayerCard)

frmMain.Dealer1.Visible = True: frmMain.Dealer1.Picture = App.Path & "\" & frmMain.Label5.Caption & ".bmp"

frmMain.Label3.Caption = dealercount

If dealercount < 17 Then
    dealMode = DEALER_NEXT
    DealCard
    
    'Do While Timer < Start + 1
    'Loop
    
Else
    DisplayHands
    If dealercount > 21 Then
    MsgBox "BUST!! The bank lost!", vbOKOnly
    playMode = MODE_GAMEOVER
    dealercount = 0
    bank = bank + bet
    Won = Won + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
ElseIf dealercount > playercount Then
DisplayHands
frmMain.Dealer1.Visible = True: frmMain.Dealer1.Picture = App.Path & "\" & frmMain.Label5.Caption & ".bmp"
    MsgBox "The bank wins!", vbOKOnly
    playMode = MODE_GAMEOVER
    dealercount = 0
    bank = bank - bet
    Lost = Lost + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
ElseIf dealercount < playercount Then
DisplayHands
    MsgBox "The bank lost!", vbOKOnly
    playMode = MODE_GAMEOVER
    dealercount = 0
    bank = bank + bet
    Won = Won + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
Else
    MsgBox "It's a draw!", vbOKOnly
    DisplayHands
    Draw = Draw + 1
    playMode = MODE_GAMEOVER
    dealercount = 0
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
End If
End If


processTestCards

End Sub


Public Sub dealerBlackjack()
DisplayHands
If nextPlayerCard = 2 And countHand(player, nextPlayerCard = 21) Then
    MsgBox "BLACKJACK, draw!", vbOKOnly
    Draw = Draw + 1
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
Else
    MsgBox "BLACKJACK, you lost!", vbOKOnly
    frmMain.Dealer1.Visible = True: frmMain.Dealer1.Picture = App.Path & "\" & frmMain.Label5.Caption & ".bmp"
    bank = bank - bet
    Lost = Lost + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
End If
playMode = MODE_STARTOVER
End Sub

Public Sub PlayerBlackjack()
MsgBox "BLACKJACK!! YOU WIN!", vbOKOnly
    bank = bank + bet
    bank = bank + (bet * 1.5)
    BlackJacks = BlackJacks + 1
    Won = Won + 1
    frmMain.lblMoney.Caption = "$ " & Int(bank)
    frmMain.cmdDeal.Enabled = True
    frmMain.cmdMinus.Enabled = True
    frmMain.cmdPlus.Enabled = True
    frmMain.cmdHit.Enabled = False
    frmMain.cmdStay.Enabled = False
    Exit Sub
End Sub

Public Sub NewGame()
GamesPlayed = GamesPlayed + 1
'frmMain.cmdDeal.Enabled = True
'frmMain.cmdMinus.Enabled = True
'frmMain.cmdPlus.Enabled = True
If bank < BET_MIN Then
    MsgBox "GAME OVER - You don't have enough money anymore!", vbOKOnly, "GAME OVER"
    Exit Sub
    App.End
End If

'Dealer1.Picture = App.Path & "\Back.bmp"
frmMain.Dealer1.Visible = False
frmMain.Dealer2.Visible = False
frmMain.Dealer3.Visible = False
frmMain.Dealer4.Visible = False
frmMain.Dealer5.Visible = False
frmMain.Dealer6.Visible = False
frmMain.Dealer7.Visible = False

frmMain.Player1.Visible = False
'Player1.Picture = App.Path & "\Back.bmp"
frmMain.Player2.Visible = False
frmMain.Player3.Visible = False
frmMain.Player4.Visible = False
frmMain.Player5.Visible = False
frmMain.Player6.Visible = False
frmMain.Player7.Visible = False

frmMain.Label2.Caption = "0"
frmMain.Label3.Caption = "0"

For i = 0 To TOTAL_DECK_CARDS
    deck(i) = i
Next i
nextDealerCard = 0
nextPlayerCard = 0

hideDealerCard = True

ShuffleDeck
DisplayHands
End Sub





Sub PlayWaveFile(WavFile)
    Dim ret
    Dim flags
    flags = SND_ASYNC Or SND_NODEFAULT
    If Sound = True Then ret = sndPlaySound(WavFile, flags)
End Sub
Public Sub Center(frm As Form)
Dim x As Integer
Dim y As Integer

x = (Screen.Width - frm.Width) / 2
y = (Screen.Height - frm.Height) / 2

frm.Left = x
frm.Top = y
End Sub
