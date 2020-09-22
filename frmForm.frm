VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmYourProgGoesHere 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Your program would replace this one."
   ClientHeight    =   1260
   ClientLeft      =   3315
   ClientTop       =   6480
   ClientWidth     =   10365
   Icon            =   "frmForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPause 
      Height          =   1215
      Left            =   10080
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   5640
      Top             =   4320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   4320
   End
   Begin RichTextLib.RichTextBox txtRitch 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   2143
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmForm.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label lblMsg 
      Caption         =   "Status : Click Check to check spelling."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Menu mnupopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu cmdAddWord 
         Caption         =   "Add Word"
      End
      Begin VB.Menu mnusuggestions 
         Caption         =   "suggestions"
      End
   End
End
Attribute VB_Name = "frmYourProgGoesHere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetTopWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const FLAGS = SWP_NOSIZE Or SWP_NOMOVE
Private Const HWND_NOTOPMOST = -2
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Moo) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Type Moo
    X As Long
    Y As Long
    End Type
    
Dim strOldText
Dim strNewText
Dim num
Dim hoverglobal

Dim howmanydefs
Dim word(1024)
Dim replaceWith(1024)


Private Sub Find_Words()
On Error Resume Next
Dim exists As Boolean
Dim n As Long
Dim awrds As Integer

'frmSpelling.lstWrong.Clear
'frmSpelling.Visible = False

DoEvents

ds1 = Timer

'txtBody.Text = Replace(txtBody.Text, vbNewLine, "")
'frmYourProgGoesHere.txtRitch.Text = txtBody.Text
length = 0
endlength = 0

' Checks the text for words and puts them in the word listbox
Dim ds() As String
ds = Split(frmYourProgGoesHere.txtRitch.Text)

'checks the words to make sure they are spelled correctly
For n = 0 To UBound(ds)

    ProgressBar1.Value = (n * 100 / UBound(ds))
    DoEvents
    
    Dim tmpstr As String
    tmpstr = ds(n)
    tmpstr = Trim(tmpstr)
    tmpstr = Replace(tmpstr, "http://", "")
    tmpstr = Replace(tmpstr, "www.", "")
    tmpstr = Replace(tmpstr, ".com", "")
    tmpstr = Replace(tmpstr, ".net", "")
    tmpstr = Replace(tmpstr, ".org", "")
    tmpstr = Replace(tmpstr, ",", "")
    tmpstr = Replace(tmpstr, "?", "")
    tmpstr = Replace(tmpstr, """", "")
    tmpstr = Replace(tmpstr, ".", "")
    
    If IsNumeric(tmpstr) Then
        tmpstr = ""
    End If
    
    Select Case UCase(Left(tmpstr, 1))
         Case "A": awrds = 0
         Case "B": awrds = 1
         Case "C": awrds = 2
         Case "D": awrds = 3
         Case "E": awrds = 4
         Case "F": awrds = 5
         Case "G": awrds = 6
         Case "H": awrds = 7
         Case "I": awrds = 8
         Case "J": awrds = 9
         Case "K": awrds = 10
         Case "L": awrds = 11
         Case "M": awrds = 12
         Case "N": awrds = 13
         Case "O": awrds = 14
         Case "P": awrds = 15
         Case "Q": awrds = 16
         Case "R": awrds = 17
         Case "S": awrds = 18
         Case "T": awrds = 19
         Case "U": awrds = 20
         Case "V": awrds = 21
         Case "W": awrds = 22
         Case "X": awrds = 23
         Case "Y": awrds = 24
         Case "Z": awrds = 25
         Case Else: awrds = 26
      End Select

      exists = alphabetWords(awrds).Exist(tmpstr)
      
     If tmpstr = "" Then  'anythign that is blank shult not be counted as a typo
        exists = True
     End If
     
     If LCase(ds(n)) <> "shouldn't" And LCase(ds(n)) <> "won't" And LCase(ds(n)) <> "can't" And exists = False And ds(n) <> "I" And ds(n) <> "A" And ds(n) <> "a" Then        'dont treat I as a typo
         If ds(n) <> " " And ds(n) <> """" And ds(n) <> vbNullString And ds(n) <> vbNewLine Then
            With frmYourProgGoesHere.txtRitch
                endlength = Len(ds(n)) + length
                .SelStart = length
                .SelLength = endlength - length
                .SelBold = True
                .SelColor = vbRed
                'length = endlength + 1
            End With
            Else
                'length = length + 1
         End If
    Else
            With frmYourProgGoesHere.txtRitch
                endlength = Len(ds(n)) + length
                .SelStart = length
                .SelLength = endlength - length
                .SelBold = False
                .SelColor = 0
                'length = endlength + 1
            End With
        
        exists = False
    End If
      If (ds(n) = " ") Then
        length = length + Len(ds(n)) + 1
      Else
        length = length + Len(ds(n)) + 1
      End If
Next n
 
ProgressBar1.Value = 100

'Me.lblMsg.Caption = "Status : " & (UBound(ds) + 1) & " Word(s) completed in " & Str(Round(Timer, 4) - Round(ds1, 4)) & " seconds."
End Sub

Private Sub cmdAddWord_Click()
    Timer2.Enabled = False
    Timer1.Enabled = False
    If (txtRitch.Text <> "") Then
        myString = Mid(txtRitch.Text, txtRitch.SelStart + 1, txtRitch.SelLength)
        myString = Replace(myString, " ", "")
        addWord (myString)
    End If
    Call cmdFind_Click
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub cmdFind_Click()
    Call Find_Words
End Sub

Private Sub cmdPause_Click()
    
    If Timer1.Enabled Then
        cmdPause.BackColor = &HC0&
    Else
        cmdPause.BackColor = &HFF00&
    End If
    
    Timer2.Enabled = Not Timer2.Enabled
    Timer1.Enabled = Not Timer1.Enabled
End Sub
Private Sub Form_Load()
  setWindowNotOnTop = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Me.Caption = "Spell checker :) "
  cmdPause.BackColor = vbGreen
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
'Dim dsyes As Integer
'dsyes = MsgBox("Do You want to clear the memory before exiting ?", vbYesNo)

'If dsyes = vbYes Then

Me.lblMsg.Caption = "Status : Please wailt while the program exits !"
DoEvents

closer

'End If

  End
End Sub

Function Hover()
    Dim nKm1 As Moo
    Dim nKmX As Long
    Dim nKmY As Long
    Call GetCursorPos(nKm1)
    nKmX = nKm1.X
    nKmY = nKm1.Y
    hoverglobal = WindowFromPointXY(nKmX, nKmY)
    Hover = hoverglobal
End Function

Function Get_Text(child)
On Error Resume Next
    Dim GetTrim
    Dim TrimSpace$
    Dim getstring
    
    wind = GetForegroundWindow()
    If child <> Me.hwnd And Me.hwnd <> wind And wind <> frmSpelling.hwnd Then
    
        GetTrim = SendMessageByNum(child, 14, 0&, 0&)
        TrimSpace$ = Space$(GetTrim)
        getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
        
        'Dim sSave As String
        'If (TrimSpace$ = "") Then
            'GetIEText (child)
        'End If
        
        If Len(TrimSpace$) > 1024 Then 'disable because too much text
            Get_Text = ""
            Timer1.Enabled = False
        Else
            Get_Text = TrimSpace$
        End If
        
    Else
        Get_Text = strOldText
    End If
    
End Function


Private Sub mnusuggestions_Click()

    myString = ""
    If (txtRitch.Text <> "") Then
        myString = Mid(txtRitch.Text, txtRitch.SelStart + 1, txtRitch.SelLength)
        myString = Replace(myString, " ", "")
    End If
    frmSpelling.Show
    setWindowNotOnTop = SetWindowPos(frmSpelling.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    
    frmSpelling.lstWrong.Clear
    frmSpelling.lstWrong.AddItem (myString)
    frmSpelling.lstWrong.Selected(0) = True
    Call frmSpelling.cmdSuggestion_Click
    
    
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    strNewText = Get_Text(hoverglobal)
    strNewText = StripHTML(strNewText)
    strNewText = Replace(strNewText, vbNewLine, "")
    If (strNewText <> strOldText) Then
        frmYourProgGoesHere.txtRitch.Text = strNewText
        Call cmdFind_Click
    End If
    strOldText = strNewText
End Sub


Private Sub Timer2_Timer()
    If GetAsyncKeyState(&H1) Then 'if the left mouse button is down
        Call Hover
        If Timer1.Enabled = False Then
            Timer1.Enabled = True
        End If
    End If
End Sub

Function StripHTML(sHTML)
Dim sTemp As String, lSpot1 As Long, lSpot2 As Long, lSpot3 As Long
    sTemp = sHTML & "<>"
    
    justincase = 0
    Do
        lSpot1& = InStr(lSpot3& + 1, sTemp$, "<")
        lSpot2& = InStr(lSpot1& + 1, sTemp$, ">")
        
        If lSpot1& = lSpot3& Or lSpot1& < 1 Then Exit Do
        If lSpot2& < lSpot1& Then lSpot2& = lSpot1& + 1
        sTemp$ = Left$(sTemp$, lSpot1& - 1) + Right$(sTemp$, Len(sTemp$) - lSpot2&)
        lSpot3& = lSpot1& - 1
        justincase = justincase + 1
        
        If (justincase > 2000) Then 'in case for some unseen reason it tryes to lock up indefinatly
            Exit Do
        End If
    Loop
    StripHTML = sTemp$
End Function

Private Sub txtRitch_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnupopup
    End If
    
    If (Button = vbLeftButton) Then
        If (txtRitch.SelLength <= 0) Then
            Start = InStrRev(txtRitch.Text, " ", txtRitch.SelStart + 1)
            If Start < 0 Then
            Start = 0
            End If
            endpoint = InStr(txtRitch.SelStart + 1, txtRitch.Text, " ")
            If endpoint <= 0 Then
                endpoint = Len(txtRitch.Text)
            End If
            txtRitch.SelStart = Start
            If endpoint - Start > 0 Then
                txtRitch.SelLength = endpoint - Start
            End If
        End If
    End If
End Sub


