VERSION 5.00
Begin VB.Form frmHowToPlay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to Play"
   ClientHeight    =   3795
   ClientLeft      =   75
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   $"frmHowToPlay.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   $"frmHowToPlay.frx":00B6
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmHowToPlay.frx":01D8
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmHowToPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub
