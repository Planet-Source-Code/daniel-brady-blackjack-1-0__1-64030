VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blackjack"
   ClientHeight    =   3960
   ClientLeft      =   915
   ClientTop       =   3030
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameOptions 
      Caption         =   "Options"
      Height          =   1695
      Left            =   3960
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
      Begin VB.CommandButton cmdStand 
         Caption         =   "Stand"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdHit 
         Caption         =   "Hit"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeal 
         Caption         =   "Deal"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frameProfit 
      Caption         =   "Profit"
      Height          =   615
      Left            =   3960
      TabIndex        =   8
      Top             =   840
      Width           =   1935
      Begin VB.TextBox txtProfit 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frameStatus 
      Caption         =   "Status"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtStatus 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame framePlayer 
      Caption         =   "Your Hand"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3700
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   6
         Left            =   2520
         Picture         =   "frmMain.frx":0000
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":1E41
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   4
         Left            =   1560
         Picture         =   "frmMain.frx":3C82
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   3
         Left            =   1080
         Picture         =   "frmMain.frx":5AC3
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   2
         Left            =   600
         Picture         =   "frmMain.frx":7904
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgPlayersCard 
         Height          =   1440
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":9745
         Top             =   280
         Width           =   1065
      End
   End
   Begin VB.Frame frameDealer 
      Caption         =   "Dealer's Hand"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3700
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   6
         Left            =   2520
         Picture         =   "frmMain.frx":B586
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   5
         Left            =   2040
         Picture         =   "frmMain.frx":D3C7
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   4
         Left            =   1560
         Picture         =   "frmMain.frx":F208
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   3
         Left            =   1080
         Picture         =   "frmMain.frx":11049
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   2
         Left            =   600
         Picture         =   "frmMain.frx":12E8A
         Top             =   280
         Width           =   1065
      End
      Begin VB.Image imgDealersCard 
         Height          =   1440
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":14CCB
         Top             =   280
         Width           =   1065
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   6000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New Game"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu HowToPlay 
         Caption         =   "How To Play"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim card(1 To 52) As String
Dim cardused(1 To 52) As String
Dim deck(1 To 52) As String
Dim profit As Integer
Dim playerscore(1 To 5) As Integer, dealerscore(1 To 5) As Integer
Dim playercards As Integer, dealercards As Integer
Dim currentcard As Integer
Dim playerscard(1 To 6) As String, dealerscard(1 To 6) As String
Dim playercount As Integer, dealercount As Integer
Dim random As Integer
Dim aces As Integer
Dim x As Integer

Private Sub Form_Load()
Randomize
cmdDeal.Enabled = False
For x = 1 To 13
    card(x) = "Heart" & x
Next x
For x = 1 To 13
    card(x + 13) = "Diamond" & x
Next x
For x = 1 To 13
    card(x + 26) = "Club" & x
Next x
For x = 1 To 13
    card(x + 39) = "Spade" & x
Next x
Call shuffle
Call deal
End Sub

Private Sub New_Click()
cmdDeal.Enabled = False
cmdHit.Enabled = True
cmdStand.Enabled = True
Call shuffle
Call deal
profit = 0
txtProfit.Text = profit
txtStatus.Text = ""
End Sub

Private Sub Exit_Click()
Unload Me
Unload frmHowToPlay
Unload frmAbout
End Sub

Private Sub HowToPlay_Click()
frmHowToPlay.Visible = True
End Sub

Private Sub About_Click()
frmAbout.Visible = True
End Sub

Private Sub cmdDeal_Click()
txtStatus.Text = ""
cmdDeal.Enabled = False
cmdHit.Enabled = True
cmdStand.Enabled = True
Call shuffle
Call deal
End Sub

Private Sub cmdHit_Click()
Call hit_player
Call calculate_player
End Sub

Private Sub cmdStand_Click()
cmdDeal.Enabled = True
cmdHit.Enabled = False
cmdStand.Enabled = False
Call calculate_player
Call dealersturn
End Sub

Sub shuffle()
For x = 1 To 52
    cardused(x) = ""
Next x
For x = 1 To 52
    random = Int(Rnd * 52) + 1
    deck(x) = card(random)
    If cardused(random) <> "" Then
        x = x - 1
    Else
        cardused(random) = "True"
    End If
Next x
End Sub

Sub deal()
For x = 3 To 6
    imgDealersCard(x).Visible = False
    imgPlayersCard(x).Visible = False
Next x
playercount = 0
dealercount = 0
currentcard = 1

playercount = playercount + 1
imgPlayersCard(playercount).Picture = LoadPicture(App.Path & "\images\" & deck(currentcard) & ".gif")
playerscard(playercount) = deck(currentcard)
currentcard = currentcard + 1

playercount = playercount + 1
imgPlayersCard(playercount).Picture = LoadPicture(App.Path & "\images\" & deck(currentcard) & ".gif")
playerscard(playercount) = deck(currentcard)
currentcard = currentcard + 1

dealercount = dealercount + 1
imgDealersCard(dealercount).Picture = LoadPicture(App.Path & "\images\" & deck(currentcard) & ".gif")
dealerscard(dealercount) = deck(currentcard)
currentcard = currentcard + 1

imgDealersCard(2).Picture = LoadPicture(App.Path & "\images\Back.gif")

Call calculate_player
End Sub

Sub hit_player()
playercount = playercount + 1
imgPlayersCard(playercount).Picture = LoadPicture(App.Path & "\images\" & deck(currentcard) & ".gif")
imgPlayersCard(playercount).Visible = True
playerscard(playercount) = deck(currentcard)
currentcard = currentcard + 1
End Sub

Sub calculate_player()
playerscore(1) = 0
playerscore(2) = 0
aces = 0
For x = 1 To playercount
    Select Case playerscard(x)
        Case Is = "Heart1", "Diamond1", "Club1", "Spade1"
            playerscore(1) = playerscore(1) + 1
            playerscore(2) = playerscore(2) + 11
            If aces = 2 Then
                playerscore(3) = playerscore(2) + 1
            End If
            If aces = 3 Then
                playerscore(4) = playerscore(3) + 1
            End If
            If aces = 4 Then
                playerscore(5) = playerscore(4) + 1
            End If
        Case Is = "Heart2", "Diamond2", "Club2", "Spade2"
            playerscore(1) = playerscore(1) + 2
            playerscore(2) = playerscore(2) + 2
        Case Is = "Heart3", "Diamond3", "Club3", "Spade3"
            playerscore(1) = playerscore(1) + 3
            playerscore(2) = playerscore(2) + 3
        Case Is = "Heart4", "Diamond4", "Club4", "Spade4"
            playerscore(1) = playerscore(1) + 4
            playerscore(2) = playerscore(2) + 4
        Case Is = "Heart5", "Diamond5", "Club5", "Spade5"
            playerscore(1) = playerscore(1) + 5
            playerscore(2) = playerscore(2) + 5
        Case Is = "Heart6", "Diamond6", "Club6", "Spade6"
            playerscore(1) = playerscore(1) + 6
            playerscore(2) = playerscore(2) + 6
        Case Is = "Heart7", "Diamond7", "Club7", "Spade7"
            playerscore(1) = playerscore(1) + 7
            playerscore(2) = playerscore(2) + 7
        Case Is = "Heart8", "Diamond8", "Club8", "Spade8"
            playerscore(1) = playerscore(1) + 8
            playerscore(2) = playerscore(2) + 8
        Case Is = "Heart9", "Diamond9", "Club9", "Spade9"
            playerscore(1) = playerscore(1) + 9
            playerscore(2) = playerscore(2) + 9
        Case Is = "Heart10", "Diamond10", "Club10", "Spade10", "Heart11", "Diamond11", "Club11", "Spade11", "Heart12", "Diamond12", "Club12", "Spade12", "Heart13", "Diamond13", "Club13", "Spade13"
            playerscore(1) = playerscore(1) + 10
            playerscore(2) = playerscore(2) + 10
    End Select
Next x
If playerscore(1) > 21 Then
    txtStatus.Text = "You bust!"
    profit = profit - 10
    txtProfit.Text = profit
    cmdDeal.Enabled = True
    cmdHit.Enabled = False
    cmdStand.Enabled = False
End If
End Sub

Sub dealersturn()
Call hit_dealer
Call calculate_dealer
If dealerscore(1) < 17 And dealerscore(2) <> 17 And dealerscore(2) <> 18 And dealerscore(2) <> 19 And dealerscore(2) <> 20 And dealerscore(2) <> 21 Then
    Call hit_dealer
    Call calculate_dealer
End If
If dealerscore(1) < 17 And dealerscore(2) <> 17 And dealerscore(2) <> 18 And dealerscore(2) <> 19 And dealerscore(2) <> 20 And dealerscore(2) <> 21 Then
    Call hit_dealer
    Call calculate_dealer
End If
If dealerscore(1) < 17 And dealerscore(2) <> 17 And dealerscore(2) <> 18 And dealerscore(2) <> 19 And dealerscore(2) <> 20 And dealerscore(2) <> 21 Then
    Call hit_dealer
    Call calculate_dealer
End If
If dealerscore(1) < 17 And dealerscore(2) <> 17 And dealerscore(2) <> 18 And dealerscore(2) <> 19 And dealerscore(2) <> 20 And dealerscore(2) <> 21 Then
    Call hit_dealer
    Call calculate_dealer
End If
If dealerscore(1) > 21 And (playerscore(2) <> 21 Or playercount <> 2) Then
    txtStatus.Text = "Dealer busts!"
    profit = profit + 10
    txtProfit.Text = profit
    cmdDeal.Enabled = True
    cmdHit.Enabled = False
    cmdStand.Enabled = False
    Exit Sub
End If
Call winner
End Sub

Sub hit_dealer()
dealercount = dealercount + 1
imgDealersCard(dealercount).Picture = LoadPicture(App.Path & "\images\" & deck(currentcard) & ".gif")
imgDealersCard(dealercount).Visible = True
dealerscard(dealercount) = deck(currentcard)
currentcard = currentcard + 1
End Sub

Sub calculate_dealer()
dealerscore(1) = 0
dealerscore(2) = 0
aces = 0
For x = 1 To dealercount
    Select Case dealerscard(x)
        Case Is = "Heart1", "Diamond1", "Club1", "Spade1"
            dealerscore(1) = dealerscore(1) + 1
            dealerscore(2) = dealerscore(2) + 11
            If aces = 2 Then
                dealerscore(3) = dealerscore(2) + 1
            End If
            If aces = 3 Then
                dealerscore(4) = dealerscore(3) + 1
            End If
            If aces = 4 Then
                dealerscore(5) = dealerscore(4) + 1
            End If
        Case Is = "Heart2", "Diamond2", "Club2", "Spade2"
            dealerscore(1) = dealerscore(1) + 2
            dealerscore(2) = dealerscore(2) + 2
        Case Is = "Heart3", "Diamond3", "Club3", "Spade3"
            dealerscore(1) = dealerscore(1) + 3
            dealerscore(2) = dealerscore(2) + 3
        Case Is = "Heart4", "Diamond4", "Club4", "Spade4"
            dealerscore(1) = dealerscore(1) + 4
            dealerscore(2) = dealerscore(2) + 4
        Case Is = "Heart5", "Diamond5", "Club5", "Spade5"
            dealerscore(1) = dealerscore(1) + 5
            dealerscore(2) = dealerscore(2) + 5
        Case Is = "Heart6", "Diamond6", "Club6", "Spade6"
            dealerscore(1) = dealerscore(1) + 6
            dealerscore(2) = dealerscore(2) + 6
        Case Is = "Heart7", "Diamond7", "Club7", "Spade7"
            dealerscore(1) = dealerscore(1) + 7
            dealerscore(2) = dealerscore(2) + 7
        Case Is = "Heart8", "Diamond8", "Club8", "Spade8"
            dealerscore(1) = dealerscore(1) + 8
            dealerscore(2) = dealerscore(2) + 8
        Case Is = "Heart9", "Diamond9", "Club9", "Spade9"
            dealerscore(1) = dealerscore(1) + 9
            dealerscore(2) = dealerscore(2) + 9
        Case Is = "Heart10", "Diamond10", "Club10", "Spade10", "Heart11", "Diamond11", "Club11", "Spade11", "Heart12", "Diamond12", "Club12", "Spade12", "Heart13", "Diamond13", "Club13", "Spade13"
            dealerscore(1) = dealerscore(1) + 10
            dealerscore(2) = dealerscore(2) + 10
    End Select
Next x
End Sub

Sub winner()
If playerscore(2) = 21 And playercount = 2 Then
    If dealerscore(2) = 21 And dealercount = 2 Then
        txtStatus.Text = "Push!"
        cmdDeal.Enabled = True
        cmdHit.Enabled = False
        cmdStand.Enabled = False
        Exit Sub
    Else
        txtStatus.Text = "Blackjack!"
        profit = profit + 15
        txtProfit.Text = profit
        cmdDeal.Enabled = True
        cmdHit.Enabled = False
        cmdStand.Enabled = False
        Exit Sub
    End If
End If
If dealerscore(2) = 21 And dealercount = 2 Then
    txtStatus.Text = "Dealer wins!"
    profit = profit - 10
    txtProfit.Text = profit
    cmdDeal.Enabled = True
    cmdHit.Enabled = False
    cmdStand.Enabled = False
    Exit Sub
End If
If playerscore(1) >= playerscore(2) Or playerscore(2) > 21 Then
    If dealerscore(1) >= dealerscore(2) Or dealerscore(2) > 21 Then
        If playerscore(1) = dealerscore(1) Then
            txtStatus.Text = "Push!"
            cmdDeal.Enabled = True
            cmdHit.Enabled = False
            cmdStand.Enabled = False
            Exit Sub
        End If
    Else
        If playerscore(1) = dealerscore(2) Then
            txtStatus.Text = "Push!"
            cmdDeal.Enabled = True
            cmdHit.Enabled = False
            cmdStand.Enabled = False
            Exit Sub
        End If
    End If
Else
    If dealerscore(1) >= dealerscore(2) Or dealerscore(2) > 21 Then
        If playerscore(2) = dealerscore(1) Then
            txtStatus.Text = "Push!"
            cmdDeal.Enabled = True
            cmdHit.Enabled = False
            cmdStand.Enabled = False
            Exit Sub
        End If
    Else
        If playerscore(2) = dealerscore(2) Then
            txtStatus.Text = "Push!"
            cmdDeal.Enabled = True
            cmdHit.Enabled = False
            cmdStand.Enabled = False
            Exit Sub
        End If
    End If
End If
If playercount = 6 And playerscore(1) <= 21 Then
    If dealercount = 6 And dealerscore(1) <= 21 Then
        txtStatus.Text = "Push!"
        cmdDeal.Enabled = True
        cmdHit.Enabled = False
        cmdStand.Enabled = False
        Exit Sub
    Else
        txtStatus.Text = "Six and under!"
        profit = profit + 10
        txtProfit.Text = profit
        cmdDeal.Enabled = True
        cmdHit.Enabled = False
        cmdStand.Enabled = False
        Exit Sub
    End If
End If
If (playerscore(1) > dealerscore(1) And playerscore(1) > dealerscore(2) And dealerscore(2) <= 21) Or (playerscore(1) > dealerscore(1) And dealerscore(2) > 21) Or (playerscore(2) > dealerscore(1) And playerscore(2) > dealerscore(2) And playerscore(2) <= 21) Or (playerscore(3) > dealerscore(1) And playerscore(3) > dealerscore(2) And playerscore(3) <= 21) Or (playerscore(4) > dealerscore(1) And playerscore(4) > dealerscore(2) And playerscore(4) <= 21) Or (playerscore(5) > dealerscore(1) And playerscore(5) > dealerscore(2) And playerscore(5) <= 21) Then
    txtStatus.Text = "You win!"
    profit = profit + 10
    txtProfit.Text = profit
    cmdDeal.Enabled = True
    cmdHit.Enabled = False
    cmdStand.Enabled = False
    Exit Sub
End If
If (dealerscore(1) > playerscore(1) And dealerscore(1) > playerscore(2) And playerscore(2) <= 21) Or (dealerscore(1) > playerscore(1) And playerscore(2) > 21) Or (dealerscore(2) > playerscore(1) And dealerscore(2) > playerscore(2) And dealerscore(2) <= 21) Then
    txtStatus.Text = "Dealer wins!"
    profit = profit - 10
    txtProfit.Text = profit
    cmdDeal.Enabled = True
    cmdHit.Enabled = False
    cmdStand.Enabled = False
    Exit Sub
End If
End Sub
