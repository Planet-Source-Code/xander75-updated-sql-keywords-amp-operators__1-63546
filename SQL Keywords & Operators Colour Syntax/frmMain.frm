VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Keywords & Colours"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtQuery 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8070
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0ECA
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00C65D21&
      Height          =   4605
      Left            =   105
      Top             =   105
      Width           =   9165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMain
' DateTime  : 07/12/2005 09:51
' Author    : Alexander Mungall
' Purpose   : SQL Query Interface
'---------------------------------------------------------------------------------------

Private Sub Form_Load()
    ' Set the Keywords and Operators into their Arrays
    Call SetKeywords
    Call SetOperators
End Sub

Private Sub txtQuery_KeyUp(KeyCode As Integer, Shift As Integer)
    ' If Ctrl+V i.e. Paste then default all the colours respectively
    If KeyCode = vbKeyV And Shift = vbCtrlMask Then
        LockWindowUpdate Me.hwnd
        txtQuery.SelStart = 0
        txtQuery.SelLength = Len(txtQuery.Text)
        txtQuery.SelColor = &H0&
        txtQuery.SelLength = 0
        Call SetDefaultColours(txtQuery)
        LockWindowUpdate 0&
        Exit Sub
    End If
    
    ' If the user presses the Return Key or Arrow Keys ignore this
    Select Case KeyCode
        Case 13, 37, 38, 39, 40
            Exit Sub
        Case 32
            ' Parse the users input and if a word is found then set its corresponding colour
            iCursorPos = txtQuery.SelStart
            FirstVisibleLine = SendMessage(txtQuery.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
            Call FindSpacePosition
            For i = 0 To 174
                ' If MultipleWords is false then 1 word found
                ' else parse the multiple words
                If MultipleWords = False Then
                    If Trim(UCase(sFoundWord)) = KeywordArray(i).sKeyword Then
                        ColorWord txtQuery, KeywordArray(i).sKeyword, KeywordArray(i).sColour, 8, False, False, False
                    End If
                Else
                    For x = 0 To UBound(sSplit())
                        sFoundWord = sSplit(x)
                        If Trim(UCase(sFoundWord)) = KeywordArray(i).sKeyword Then
                            ColorWord txtQuery, KeywordArray(i).sKeyword, KeywordArray(i).sColour, 8, False, False, False
                        End If
                    Next
                End If
            Next
            
            ' Parse the users input and if an operator is found then set its corresponding colour
            For i = 0 To 8
                ColorDefaultWord txtQuery, OperatorArray(i).sOperator, OperatorArray(i).sColour, 8, False, False, False
            Next
            LockWindowUpdate 0&
            charIndex = SendMessage(txtQuery.hwnd, EM_LINEINDEX, ByVal FirstVisibleLine, ByVal CLng(0))
            txtQuery.SetFocus
            If charIndex <> -1 Then
                txtQuery.SelStart = charIndex
            End If
            txtQuery.SelStart = iCursorPos
            Exit Sub
    End Select
    
    ' Parse the users input and if a word is found then set its corresponding colour
    iCursorPos = txtQuery.SelStart
    FirstVisibleLine = SendMessage(txtQuery.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    Call FindWordPosition
    For i = 0 To 174
        ' If MultipleWords is false then 1 word found
        ' else parse the multiple words
        If MultipleWords = False Then
            If Trim(UCase(sFoundWord)) = KeywordArray(i).sKeyword Then
                ColorWord txtQuery, KeywordArray(i).sKeyword, KeywordArray(i).sColour, 8, False, False, False
            End If
        Else
            For x = 0 To UBound(sSplit())
                sFoundWord = sSplit(x)
                If Trim(UCase(sFoundWord)) = KeywordArray(i).sKeyword Then
                    ColorWord txtQuery, KeywordArray(i).sKeyword, KeywordArray(i).sColour, 8, False, False, False
                End If
            Next
        End If
    Next
    
    ' Parse the users input and if an operator is found then set its corresponding colour
    For i = 0 To 8
        ColorDefaultWord txtQuery, OperatorArray(i).sOperator, OperatorArray(i).sColour, 8, False, False, False
    Next
    LockWindowUpdate 0&
    charIndex = SendMessage(txtQuery.hwnd, EM_LINEINDEX, ByVal FirstVisibleLine, ByVal CLng(0))
    txtQuery.SetFocus
    If charIndex <> -1 Then
        txtQuery.SelStart = charIndex
    End If
    txtQuery.SelStart = iCursorPos
End Sub

Sub FindSpacePosition()
    ' Flag used to check if a user has concatenated words
    MultipleWords = False
    
    ' Get the first letter of the word
    sQueryStr = Left(txtQuery.Text, txtQuery.SelStart)
    sTempStr = Replace(sQueryStr, vbCrLf, "  ")
    sTempStr = StrReverse(Trim(sTempStr))
    iPos = InStr(sTempStr, " ")
    If iPos = 0 Then
        iFirstLetterPos = 1
    Else
        sTempStr = StrReverse(sTempStr)
        iFirstLetterPos = (Len(sTempStr)) - (iPos) + 2
        sFoundWord = Mid(sTempStr, iFirstLetterPos, 1)
    End If
    
    ' Get the last letter of the word
    sQueryStr = txtQuery.Text
    sQueryStr = Replace(sQueryStr, vbCrLf, "  ")
    sTempStr = Mid(sQueryStr, iFirstLetterPos, Len(txtQuery.Text))
    iPos = InStr(sTempStr, " ")
    If iPos = 0 Then
        iLastLetterPos = Len(sQueryStr)
        sFoundWord = Mid(sQueryStr, iFirstLetterPos, ((iLastLetterPos + 1) - iFirstLetterPos))
    Else
        iLastLetterPos = iPos - 1
        sFoundWord = Mid(txtQuery.Text, iFirstLetterPos, iLastLetterPos)
    End If
    
    ' If the input is concatenated i.e. SUBSTR(CASE then parse for multiple words
    ' else only 1 word found
    If InStr(sFoundWord, "(") Or InStr(sFoundWord, ")") Or InStr(sFoundWord, "+") Or InStr(sFoundWord, "-") Or _
       InStr(sFoundWord, "*") Or InStr(sFoundWord, "/") Or InStr(sFoundWord, "<") Or InStr(sFoundWord, ">") Or _
       InStr(sFoundWord, "~") Or InStr(sFoundWord, "'") Then
        If InStr(sFoundWord, "(") Then sFoundWord = Replace(sFoundWord, "(", " ")
        If InStr(sFoundWord, ")") Then sFoundWord = Replace(sFoundWord, ")", " ")
        If InStr(sFoundWord, "+") Then sFoundWord = Replace(sFoundWord, "+", " ")
        If InStr(sFoundWord, "-") Then sFoundWord = Replace(sFoundWord, "-", " ")
        If InStr(sFoundWord, "*") Then sFoundWord = Replace(sFoundWord, "*", " ")
        If InStr(sFoundWord, "/") Then sFoundWord = Replace(sFoundWord, "/", " ")
        If InStr(sFoundWord, "<") Then sFoundWord = Replace(sFoundWord, "<", " ")
        If InStr(sFoundWord, ">") Then sFoundWord = Replace(sFoundWord, ">", " ")
        If InStr(sFoundWord, "~") Then sFoundWord = Replace(sFoundWord, "~", " ")
        If InStr(sFoundWord, "'") Then sFoundWord = Replace(sFoundWord, "'", " ")
        sSplit() = Split(sFoundWord, " ")
        MultipleWords = True
        LockWindowUpdate Me.hwnd
        txtQuery.SelStart = iFirstLetterPos - 1
        txtQuery.SelLength = iLastLetterPos
        txtQuery.SelColor = &H0&
        txtQuery.SelLength = 0
    Else
        LockWindowUpdate Me.hwnd
        txtQuery.SelStart = iFirstLetterPos - 1
        txtQuery.SelLength = iLastLetterPos
        txtQuery.SelColor = &H0&
        txtQuery.SelLength = 0
    End If
End Sub

Sub FindWordPosition()
    ' Flag used to check if a user has concatenated words
    MultipleWords = False
    
    ' Get the first letter of the word
    sQueryStr = Left(txtQuery.Text, txtQuery.SelStart)
    sTempStr = Replace(sQueryStr, vbCrLf, "  ")
    sTempStr = StrReverse(sTempStr)
    iPos = InStr(sTempStr, " ")
    If iPos = 0 Then
        iFirstLetterPos = 1
    Else
        sTempStr = StrReverse(sTempStr)
        iFirstLetterPos = (Len(sTempStr)) - (iPos) + 2
        sFoundWord = Mid(sTempStr, iFirstLetterPos, 1)
    End If
    
    ' Get the last letter of the word
    sQueryStr = txtQuery.Text
    sQueryStr = Replace(sQueryStr, vbCrLf, "  ")
    sTempStr = Mid(sQueryStr, iFirstLetterPos, Len(txtQuery.Text))
    iPos = InStr(sTempStr, " ")
    If iPos = 0 Then
        iLastLetterPos = Len(sQueryStr)
        sFoundWord = Mid(sQueryStr, iFirstLetterPos, ((iLastLetterPos + 1) - iFirstLetterPos))
    Else
        iLastLetterPos = iPos - 1
        sFoundWord = Mid(txtQuery.Text, iFirstLetterPos, iLastLetterPos)
    End If
    
    ' If the input is concatenated i.e. SUBSTR(CASE then parse for multiple words
    ' else only 1 word found
    If InStr(sFoundWord, "(") Or InStr(sFoundWord, ")") Or InStr(sFoundWord, "+") Or InStr(sFoundWord, "-") Or _
       InStr(sFoundWord, "*") Or InStr(sFoundWord, "/") Or InStr(sFoundWord, "<") Or InStr(sFoundWord, ">") Or _
       InStr(sFoundWord, "~") Or InStr(sFoundWord, "'") Then
        If InStr(sFoundWord, "(") Then sFoundWord = Replace(sFoundWord, "(", " ")
        If InStr(sFoundWord, ")") Then sFoundWord = Replace(sFoundWord, ")", " ")
        If InStr(sFoundWord, "+") Then sFoundWord = Replace(sFoundWord, "+", " ")
        If InStr(sFoundWord, "-") Then sFoundWord = Replace(sFoundWord, "-", " ")
        If InStr(sFoundWord, "*") Then sFoundWord = Replace(sFoundWord, "*", " ")
        If InStr(sFoundWord, "/") Then sFoundWord = Replace(sFoundWord, "/", " ")
        If InStr(sFoundWord, "<") Then sFoundWord = Replace(sFoundWord, "<", " ")
        If InStr(sFoundWord, ">") Then sFoundWord = Replace(sFoundWord, ">", " ")
        If InStr(sFoundWord, "~") Then sFoundWord = Replace(sFoundWord, "~", " ")
        If InStr(sFoundWord, "'") Then sFoundWord = Replace(sFoundWord, "'", " ")
        sSplit() = Split(sFoundWord, " ")
        MultipleWords = True
        LockWindowUpdate Me.hwnd
        txtQuery.SelStart = iFirstLetterPos - 1
        txtQuery.SelLength = iLastLetterPos
        txtQuery.SelColor = &H0&
        txtQuery.SelLength = 0
    Else
        LockWindowUpdate Me.hwnd
        txtQuery.SelStart = iFirstLetterPos - 1
        txtQuery.SelLength = iLastLetterPos
        txtQuery.SelColor = &H0&
        txtQuery.SelLength = 0
    End If
End Sub
