VERSION 5.00
Begin VB.Form Questions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-=Math Game=-"
   ClientHeight    =   2400
   ClientLeft      =   6630
   ClientTop       =   5445
   ClientWidth     =   3990
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MathQuestions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3990
   Begin VB.CommandButton SettingsCmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton ExitCmd 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton OkCmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox AnswerTxt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Find..."
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label XBLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label SLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label XALabel 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Questions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Xa As Integer 'the first number the user will use.
Dim Xb As Integer 'the second number the user will use.
Dim Result As Integer 'the result of the operation between Xa and Xb, it will be compared with what the user writes.
Dim ScorePoints As Integer 'the points of the user.
Dim ScoreTotal As Integer 'the number of operations the user did.
Dim PickGame As Integer 'a variable used when it's a random game.

Private Sub AnswerTxt_KeyPress(KeyAscii As Integer)
    'The user can't write anything else than numbers and back spaces (keyascii = 8) in the answer-text.
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

'When the form is loaded,
Private Sub Form_Load()
    'change the caption of the form and buttons depending on the language
    Select Case Language
        Case 2
            Questions.Caption = "-=Jeu de Math=-"
            SettingsCmd.Caption = "&Paramètres"
            ExitCmd.Caption = "&Quitter"
            Label1.Caption = "Trouvez..."
        Case 3
            Questions.Caption = "-=Mathespiel=-"
            SettingsCmd.Caption = "&Parameter"
            ExitCmd.Caption = "&Abbrechen"
            Label1.Caption = "Finden..."
    End Select
    'set the score to 0/0
    ScorePoints = 0
    ScoreTotal = 0
    'randomize the random numbers
    Randomize
    'and call this sub.
    ChooseMaj
End Sub

'When this sub is called,
Private Sub ChooseMaj()
    'depending on the game, call one of these subs.
    Select Case Game
        Case 1
            MajRandom
        Case 2
            MajAdd
        Case 3
            MajSub
        Case 4
            MajMul
        Case 5
            MajDiv
    End Select
End Sub


'When the user clicks on this button,
Private Sub OkCmd_Click()
    'check if he wrote an answer or not,
    If AnswerTxt.Text = "" Then
        'if not, tell him by sending a message depending on his language,
        Select Case Language
            Case 1
                MsgBox "You forgot to write the answer.", vbOKOnly + vbCritical, "-=Math Game=-  Mistake"
            Case 2
                MsgBox "Vous avez oublié d'écrire votre réponse.", vbOKOnly + vbCritical, "-=Jeu de Math=-  Erreur"
            Case 3
                MsgBox "Sie haben vergessen, das Ergebnis einzugeben.", vbOKOnly + vbCritical, "-=Mathespiel=-  Fehler"
        End Select
        'and exit the sub;
        Exit Sub
    'if the user wrote an answer and it's correct,
    ElseIf AnswerTxt.Text = Result Then
        'call the subs below,
        GoodAnswer
        ChooseMaj
        'clear the answer text
        AnswerTxt.Text = ""
        'and put the focus there;
        AnswerTxt.SetFocus
    'if the user wrote an answer and it's wrong,
    ElseIf AnswerTxt.Text <> Result Then
        'send him a message telling so depending on his language,
        Select Case Language
            Case 1
                MsgBox "No, that isn't the right answer.", vbExclamation, "-=Math Game=-  Mistake"
            Case 2
                MsgBox "Non, ce n'est pas la bonne réponse.", vbExclamation, "-=Jeu de Math=-  Erreur"
            Case 3
                MsgBox "Nein, dass ist nicht die richtige Antwort.", vbExclamation, "-=Mathespiel=-  Fehler"
        End Select
        'call this sub,
        BadAnswer
        'clear
        AnswerTxt.Text = ""
        'and set the focus to the answer text.
        AnswerTxt.SetFocus
    End If
End Sub

'When this sub is called,
Private Sub MajRandom()
    'give a random number (between 1 and 4) to PickGame
    PickGame = Rnd * 3 + 1
    'and use that random number to call one of the subs below.
    Select Case PickGame
        Case 1
            MajAdd
        Case 2
            MajSub
        Case 3
            MajMul
        Case 4
            MajDiv
    End Select
End Sub

'When this sub is called,
Private Sub MajAdd()
    'depending on the level,
    Select Case Level
        Case 1
            'give Xa, Bb and result a new value.
            Xa = Rnd * 9 + 1
            Xb = Rnd * 9 + 1
            Result = Xa + Xb
        Case 2
            Xa = Rnd * 99 + 1
            Xb = Rnd * 99 + 1
            Result = Xa + Xb
        Case 3
            Xa = Rnd * 999 + 1
            Xb = Rnd * 99 + 1
            Result = Xa + Xb
    End Select
    'write the value for the user
    XALabel.Caption = Xa
    XBLabel.Caption = Xb
    'and write the sign for the user.
    SLabel.Caption = "+"
End Sub

'When this sub is called,
Private Sub MajSub()
    'same for MajAdd,
    Select Case Level
        Case 1
            Do
                Xa = Rnd * 19 + 1
                Xb = Rnd * 9 + 1
            Loop Until Xa > Xb 'but make sure that Xa is greater than Xb.
            Result = Xa - Xb
        Case 2
            Do
                Xa = Rnd * 99 + 1
                Xb = Rnd * 9 + 1
            Loop Until Xa > Xb
            Result = Xa - Xb
        Case 3
            Do
                Xa = Rnd * 99 + 1
                Xb = Rnd * 99 + 1
            Loop Until Xa > Xb
            Result = Xa - Xb
    End Select
    XALabel.Caption = Xa
    XBLabel.Caption = Xb
    SLabel.Caption = "-"
End Sub

'When this sub is called,
Private Sub MajMul()
    'Same for MajAdd
    Select Case Level
        Case 1
            Do
                Xa = Rnd * 8 + 2
                Xb = Rnd * 9 + 1
                Result = Xa * Xb
            Loop Until Result < 51 'but make sure that the answer isn't greater than 51 for the easy level.
        Case 2
            Do
                Xa = Rnd * 13 + 2
                Xb = Rnd * 10 + 1
                Result = Xa * Xb
            Loop Until Result < 101 'and the answer isn't greater than 101 for the normal level.
        Case 3
            Xa = Rnd * 18 + 2
            Xb = Rnd * 19 + 1
            Result = Xa * Xb
    End Select
    XALabel.Caption = Xa
    XBLabel.Caption = Xb
    SLabel.Caption = "x"
End Sub

'When this sub is called,
Private Sub MajDiv()
    'Same for MajAdd,
    Select Case Level
        Case 1
            Do
                Randomize
                Xa = Rnd * 19 + 1
                Xb = Rnd * 9 + 2
                Result = Xa / Xb
            Loop Until Xa Mod Xb = 0 'but make sure that the result isn't a decimal number.
        Case 2
            Do
                Xa = Rnd * 39 + 1
                Xb = Rnd * 9 + 2
                Result = Xa / Xb
            Loop Until Xa Mod Xb = 0 And Result <> 1
        Case 3
            Do
                Xa = Rnd * 99 + 1
                Xb = Rnd * 99 + 2
                Result = Xa / Xb
            Loop Until Xa Mod Xb = 0 And Result <> 1
    End Select
    XALabel.Caption = Xa
    XBLabel.Caption = Xb
    SLabel.Caption = "÷"
End Sub

'When this sub is called,
Private Sub BadAnswer()
    'Add 1 to the original score.
    ScoreTotal = ScoreTotal + 1
End Sub

'When this sub is called,
Private Sub GoodAnswer()
    'Add 1 to the user's points and to the original score.
    ScorePoints = ScorePoints + 1
    ScoreTotal = ScoreTotal + 1
End Sub

'When this sub is called,
Private Sub ExitCmd_Click()
    'call this sub
    GiveScore
    'and end the programm.
    End
End Sub

'When this sub is called,
Private Sub GiveScore()
    'give the score to the user depending on his language.
    Select Case Language
        Case 1
            MsgBox Luser & ", you got " & ScorePoints & " out of " & ScoreTotal & ".", vbOKOnly + vbInformation, "-=Math Game=-  Score"
        Case 2
            MsgBox Luser & ", vous avez eu " & ScorePoints & " sur " & ScoreTotal & ".", vbOKOnly + vbInformation, "-=Jeu de Math=-  Score"
        Case 3
            MsgBox Luser & ", Sie haben " & ScorePoints & " von " & ScoreTotal & " richtig.", vbOKOnly + vbInformation, "-=Mathespiel=-  Ergebnis"
    End Select
End Sub

'When the user clicks on this button,
Private Sub SettingsCmd_Click()
    'call this sub,
    GiveScore
    'unload the form,
    Unload Me
    'set the score to 0/0
    ScoreTotal = 0
    ScorePoints = 0
    'and show the settings-form.
    Settings.Visible = True
End Sub
