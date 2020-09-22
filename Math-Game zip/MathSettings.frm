VERSION 5.00
Begin VB.Form Settings 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=Math Game=-  Settings"
   ClientHeight    =   3750
   ClientLeft      =   5700
   ClientTop       =   3930
   ClientWidth     =   5760
   Icon            =   "MathSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5760
   Begin VB.CommandButton ExitCmd 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton ContinueCmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox NameTxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Game mode"
      Height          =   2295
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   2295
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Divisions only"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Multiplications only"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   8
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Substractions only"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Additions only"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Everything randomly"
         Height          =   315
         Left            =   120
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.HScrollBar LevelScroll 
      Height          =   375
      Left            =   360
      Max             =   3
      Min             =   1
      TabIndex        =   1
      Top             =   2640
      Value           =   1
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1485
      Left            =   240
      Picture         =   "MathSettings.frx":0442
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter your name :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label LevelLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'When the settings-form is loaded,
Private Sub Form_Load()
    'depending on the language
    Select Case Language
        'change the caption of the different labels, options and buttons depending on the language
        Case 2 '(french)
            LevelLabel.Caption = "Facile"
            Option1.Caption = "Tout aléatoirement"
            Option2.Caption = "Additions seulement"
            Option3.Caption = "Soustractions seulement"
            Option4.Caption = "Multiplications seulement"
            Option5.Caption = "Divisions seulement"
            Frame1.Caption = "Type de jeux"
            Label1.Caption = "Entrez votre nom :"
            Settings.Caption = "-=Jeu de Math=-  Paramètres"
            ContinueCmd.Caption = "&Continuer"
            ExitCmd.Caption = "&Quitter"
        Case 3 '(german)
            LevelLabel.Caption = "Leicht"
            Option1.Caption = "Alles gemischt"
            Option2.Caption = "Nur Additionen"
            Option3.Caption = "Nur Substraktionen"
            Option4.Caption = "Nur Multiplikationen"
            Option5.Caption = "Nur Teilungen"
            Label1.Caption = "Geben Sie Ihren Name ein:"
            Settings.Caption = "-=Mathspiele=-  Einstellung"
            ContinueCmd.Caption = "&Weiter"
            ExitCmd.Caption = "&Abbrechen"
    End Select
    'and set the color of the level-label.
    LevelLabel.ForeColor = RGB(1, 135, 205)
End Sub

'When the user clicks on the continue-button,
Private Sub ContinueCmd_Click()
    'check if his name is entered, if not,
    If NameTxt.Text = "" Then
        'send a message to the user depending on his language,
        Select Case Language
            Case 1 '
                MsgBox "You forgot to enter your name.", vbOKOnly + vbCritical, "-=Math Game=-  Mistake"
            Case 2 '
                MsgBox "Vous avez oublié d'inscrire votre nom.", vbOKOnly + vbCritical, "-=Jeu de Math=-  Erreur"
            Case 3 '
                MsgBox "Sie haben vergessen, Ihren Namen einzugeben.", vbOKOnly + vbCritical, "-=Mathspiele=-  Fehler"
        End Select
        'and exit the sub;
        Exit Sub
    'but if it is properly entered,
    Else
        'call the following subs,
        MajGame
        MajLevel
        MajLuser
        'hide the settings-form
        Settings.Visible = False
        'and show the questions-form.
        Questions.Visible = True
    End If
End Sub

'When this sub is called,
Private Sub MajLuser()
    'set the luser variable to what the user wrote in the name-text.
    Luser = NameTxt.Text
End Sub

'When this sub is called,
Private Sub MajGame()
    'depending on which option is clicked, set the Game-variable.
    If Option1.Value = True Then _
        Game = 1
    If Option2.Value = True Then _
        Game = 2
    If Option3.Value = True Then _
        Game = 3
    If Option4.Value = True Then _
        Game = 4
    If Option5.Value = True Then _
        Game = 5
End Sub

'When this sub is called,
Private Sub MajLevel()
    'set the level variable to the value of the scroll bar.
    Level = LevelScroll.Value
End Sub

'When the user clicks on the exit-button,
Private Sub ExitCmd_Click()
    'end the programm.
    End
End Sub

'When the user moves the scroll bar,
Private Sub LevelScroll_Change()
    'depending on the value of the scroll bar,
    If LevelScroll.Value = 1 Then
        'change the color of the text
        LevelLabel.ForeColor = RGB(1, 135, 205)
        'and depending on the language
        Select Case Language
            'write the appropriate level in the level-label, etc.
            Case 1
                LevelLabel.Caption = "Easy"
            Case 2
                LevelLabel.Caption = "Facile"
            Case 3
                LevelLabel.Caption = "Leicht"
        End Select
    ElseIf LevelScroll.Value = 2 Then
        LevelLabel.Caption = "Normal"
        LevelLabel.ForeColor = RGB(1, 120, 120)
    ElseIf LevelScroll.Value = 3 Then
        LevelLabel.ForeColor = RGB(1, 155, 155)
        Select Case Language
            Case 1
                LevelLabel.Caption = "Hard"
            Case 2
                LevelLabel.Caption = "Difficile"
            Case 3
                LevelLabel.Caption = "Schwierig"
        End Select
    End If
End Sub
