VERSION 5.00
Begin VB.Form Welcome 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=Math Game=-  Welcome"
   ClientHeight    =   3150
   ClientLeft      =   6420
   ClientTop       =   4650
   ClientWidth     =   4320
   Icon            =   "MathWelcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Choose your language"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
      Begin VB.OptionButton GermanOption 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deutsch"
         Height          =   315
         Left            =   2520
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton FrenchOption 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Français"
         Height          =   315
         Left            =   1320
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton EnglishOption 
         BackColor       =   &H00E0E0E0&
         Caption         =   "English"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton ExitCmd 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton PlayCmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label WelcomeLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Welcome to the Math Game!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'When the user clicks on the english-option,
Private Sub EnglishOption_Click()
    'change the caption of the frame, the form and the buttons to what's written below
    Frame1.Caption = "Choose your language"
    WelcomeLabel.Caption = "Welcome to the Math Game!"
    PlayCmd.Caption = "&Play"
    ExitCmd.Caption = "&Exit"
    Welcome.Caption = "-=Math Game=-  Welcome"
    'and store 1 in the language variable.
    Language = 1
End Sub

'When the user clicks on the exit-button,
Private Sub ExitCmd_Click()
    'end the programm.
    End
End Sub

Private Sub Form_Load()
    'By default Language is 1 (english), then it can change by clicking on the other options
    Language = 1
End Sub

'When the user clicks on the french-option,
Private Sub FrenchOption_Click()
    'change the caption of the frame, the form and the buttons to what's written below
    Frame1.Caption = "Choisissez votre langue"
    WelcomeLabel.Caption = "Bienvenue au Jeux de Math !"
    PlayCmd.Caption = "&Jouer"
    ExitCmd.Caption = "&Quitter"
    Welcome.Caption = "-=Jeu de Math=-  Bienvenue"
    'and store 2 in the language variable.
    Language = 2
End Sub

'When you click on the german-option,
Private Sub GermanOption_Click()
    'change the caption of the frame, the form and the buttons to what's written below
    Frame1.Caption = "Wahlen ihre Sprache aus"
    WelcomeLabel.Caption = "Willkommen im Mathespiel!"
    PlayCmd.Caption = "&Spielen"
    ExitCmd.Caption = "&Abbrechen"
    Welcome.Caption = "-=Mathspiele=-  Willkommen"
    'and store 3 in the language variable.
    Language = 3
End Sub

'When the user clicks on the play-button,
Private Sub PlayCmd_Click()
    'unload the welcome-form
    Unload Me
    'and show the settings-form.
    Settings.Visible = True
End Sub
