VERSION 5.00
Begin VB.Form Tris 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3450
      ScaleMode       =   0  'User
      ScaleWidth      =   3454.936
      TabIndex        =   0
      Top             =   0
      Width           =   3450
   End
   Begin VB.Menu partitaMenuBtn 
      Caption         =   "Partita"
      Begin VB.Menu NuovaPartitaMenuBtn 
         Caption         =   "Nuova Partita"
      End
      Begin VB.Menu separator1 
         Caption         =   "-"
      End
      Begin VB.Menu EsciMenuBtn 
         Caption         =   "Esci"
      End
   End
End
Attribute VB_Name = "Tris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Turno As Integer
Dim Matrice(3, 3) As Integer
Dim Vincitore As String

Private Sub EsciMenuBtn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Reset
End Sub

Private Sub NuovaPartitaMenuBtn_Click()
    Reset
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim x1 As Integer: x1 = x / Picture1.Width * 230  'pos pixel x
    Dim y1 As Integer: y1 = y / Picture1.Height * 230 'pos pixel y
    Dim x2 As Integer: x2 = IIf(x1 < 73, 0, IIf(x1 >= 79 And x1 < 151, 1, IIf(x1 >= 157, 2, -1))) 'pos cella x
    Dim y2 As Integer: y2 = IIf(y1 < 73, 0, IIf(y1 >= 79 And y1 < 151, 1, IIf(y1 >= 157, 2, -1))) 'pos cella y
    If Not x2 = -1 And Not y2 = -1 Then 'controllo click non su riga nera
        If Matrice(x2, y2) = 0 Then 'controllo casella vuota
            Mark x2, y2 'segna
            Vincitore = CheckWin 'controlla vincitore
            If Not Vincitore = "" Then 'controlla se c'è un esito
                MsgBox (Vincitore&"\nNuova partita?")
            End If
        End If
    End If
End Sub

Private Function CheckWin() As String
    Dim result As Integer: result = 0
    Dim tie As Boolean: tie = True
    'controllo tris orizzontale e verticale
    For i = 0 To 2
        If Matrice(0, i) = Matrice(1, i) And Matrice(1, i) = Matrice(2, i) And Not Matrice(0, i) = 0 Then
            result = Matrice(0, i)
        ElseIf Matrice(i, 0) = Matrice(i, 1) And Matrice(i, 1) = Matrice(i, 2) And Not Matrice(i, 0) = 0 Then
            result = Matrice(i, 0)
        End If
    Next
    'controllo tris diagonale
    If (Matrice(0, 0) = Matrice(1, 1) And Matrice(1, 1) = Matrice(2, 2) And Not Matrice(1, 1) = 0) Or (Matrice(0, 2) = Matrice(1, 1) And Matrice(2, 0) = Matrice(0, 2) And Not Matrice(1, 1) = 0) Then
        result = Matrice(1, 1)
    End If
    'controllo pareggio
    For y_loc = 0 To 2
        For x_loc = 0 To 2
            If Matrice(x_loc, y_loc) = 0 Then
                tie = False
                Exit For
            End If
        Next
    Next
    CheckWin = IIf(result = 1, "Il vincitore è il giocatore 1 (O)!", IIf(result = 2, "Il vincitore è il giocatore 2 (X)!", IIf(tie, "Pareggio!", "")))
End Function

Private Sub Mark(x As Integer, y As Integer)
    Matrice(x, y) = Turno 'segna in memoria
    Set Cmd1 = Controls.Add("vb.PictureBox", "PIC" & CStr(x) & CStr(y)) 'crea nuova picturebox
    Cmd1.Width = Picture1.Width / 230 * 72
    Cmd1.Height = Picture1.Width / 230 * 72
    Cmd1.Top = IIf(y = 1, 79, IIf(y = 2, 157, 1)) * Picture1.Width / 230 'coordinata y
    Cmd1.Appearance = 0
    Cmd1.BorderStyle = 0
    Cmd1.Left = IIf(x = 1, 79, IIf(x = 2, 157, 1)) * Picture1.Width / 230 'coordinata x
    Cmd1.Picture = LoadResPicture(100 + Turno, vbResBitmap)
    Cmd1.Visible = True
    Cmd1.ZOrder (fmTop)
    Turno = IIf(Turno = 1, 2, 1)
End Sub

Private Sub Reset()
    Vincitore = ""
    Turno = 1
     For y_loc = 0 To 2
        For x_loc = 0 To 2
            Matrice(x_loc, y_loc) = 0
            For Each ctl In Controls
                If ctl.Name = "PIC" & CStr(x_loc) & CStr(y_loc) Then
                    Controls.Remove ctl
                End If
            Next
        Next
    Next
End Sub
