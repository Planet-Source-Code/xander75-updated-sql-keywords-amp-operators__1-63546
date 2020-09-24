Attribute VB_Name = "modSQLKeywords"
'---------------------------------------------------------------------------------------
' Module    : modSQLKeywords
' DateTime  : 07/12/2005 09:33
' Author    : Alexander Mungall
' Purpose   : Colours Keywords and Operators
'---------------------------------------------------------------------------------------
        
    ' Arrays
    Public KeywordArray(175) As KeywordList
    Type KeywordList
        sKeyword As String
        sColour As ColorConstants
    End Type
    
    Public OperatorArray(9) As OperatorList
    Type OperatorList
        sOperator As String
        sColour As ColorConstants
    End Type
    
    ' Booleans
    Public MultipleWords As Boolean
    
    ' Integers
    Public iFirstLetterPos As Integer
    Public iLastLetterPos As Integer
    
    ' Strings
    Public sFoundWord As String
    Public sQueryStr As String
    Public sSplit() As String
    Public sTempStr As String
    
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Const EM_LINEINDEX = &HBB
    Public Const EM_GETFIRSTVISIBLELINE = &HCE

' Used for resetting the entire textbox
Public Function ColorDefaultWord(RTF As RichTextBox, sWords As String, sColor As ColorConstants, _
SFontsize As Integer, Sbold As Boolean, Sitalic As Boolean, Sbullet As Boolean)
    search = 1
    Do Until search = 0
        search = InStr(search, RTF.Text, sWords, vbTextCompare)
        If search > 0 Then
            With RTF
                .SelStart = search - 1
                .SelLength = Len(sWords)
                If .SelStart = 0 Then
                    GetFirstLetter = ""
                    GetLastLetter = Mid(RTF.Text, (.SelStart + .SelLength) + 1, 1)
                    If GetLastLetter <> "" Then AscLastLetter = Asc(GetLastLetter)
                    If (GetFirstLetter = "" Or GetFirstLetter = " " Or AscFirstLetter = 10 Or AscFirstLetter = 13) And (GetLastLetter = "" Or GetLastLetter = " " Or AscLastLetter = 10 Or AscLastLetter = 13) Then
                        .SelColor = sColor
                    Else
                        If .SelColor = vbBlack Then .SelColor = vbBlack
                    End If
                    .SelFontSize = SFontsize
                    .SelBold = Sbold
                    .SelItalic = Sitalic
                    .SelBullet = Sbullet
                Else
                    .SelStart = search
                    .SelLength = Len(sWords)
                    GetFirstLetter = Mid(RTF.Text, .SelStart - 1, 1)
                    AscFirstLetter = Asc(GetFirstLetter)
                    GetLastLetter = Mid(RTF.Text, (.SelStart + .SelLength), 1)
                    If GetLastLetter <> "" Then AscLastLetter = Asc(GetLastLetter)

                    If sWords = "*" Or sWords = "+" Or sWords = "-" Or sWords = "/" Or sWords = "(" Or sWords = ")" Or sWords = "~" Or _
                       sWords = "<" Or sWords = ">" Or sWords = "+" Then
                        .SelStart = search - 1
                        .SelLength = Val(Len(sWords))
                        .SelColor = sColor
                    Else
                        If (AscFirstLetter = "10" Or AscFirstLetter = "13" Or AscFirstLetter = "32" Or _
                            AscFirstLetter = "40" Or AscFirstLetter = "41" Or AscFirstLetter = "42" Or AscFirstLetter = "43" Or _
                            AscFirstLetter = "45" Or AscFirstLetter = "47" Or AscFirstLetter = "60" Or AscFirstLetter = "61" Or _
                            AscFirstLetter = "62" Or AscFirstLetter = "126" And (AscFirstLetter <> "95") And (AscFirstLetter <> "95")) Then
                            If (AscLastLetter = "10" Or AscLastLetter = "13" Or AscLastLetter = "32" Or _
                                AscLastLetter = "40" Or AscLastLetter = "41" Or AscLastLetter = "42" Or AscLastLetter = "43" Or _
                                AscLastLetter = "45" Or AscLastLetter = "47" Or AscLastLetter = "60" Or AscLastLetter = "61" Or _
                                AscLastLetter = "62" Or AscLastLetter = "126" Or GetLastLetter = "" And (AscLastLetter <> "95")) Then
                                .SelStart = search - 1
                                .SelLength = Len(sWords)
                                .SelColor = sColor
                            Else
                                If .SelColor = vbBlack Then .SelColor = vbBlack
                            End If
                        Else
                            If .SelColor = vbBlack Then .SelColor = vbBlack
                        End If
                    End If
                    .SelFontSize = SFontsize
                    .SelBold = Sbold
                    .SelItalic = Sitalic
                    .SelBullet = Sbullet
                End If
            End With
            search = search + Len(sWords)
        End If
    Loop
    With RTF
        .SelStart = Len(RTF.Text)
        .SelColor = vbBlack
        .SelFontSize = 8
        .SelBold = False
        .SelItalic = False
        .SelBullet = False
    End With
End Function

' Used to set the colours of individual words
Public Function ColorWord(RTF As RichTextBox, sWords As String, sColor As ColorConstants, _
SFontsize As Integer, Sbold As Boolean, Sitalic As Boolean, Sbullet As Boolean)
    search = 1
    Do Until search = 0
        search = InStr(search, RTF.Text, sWords, vbTextCompare)
        If search > 0 Then
            If search >= iFirstLetterPos Then
                With RTF
                    .SelStart = search - 1
                    .SelLength = Len(sWords)
                    If .SelStart = 0 Then
                        GetFirstLetter = ""
                    Else
                        GetFirstLetter = Mid(RTF.Text, .SelStart + 1, 1)
                        AscFirstLetter = Asc(GetFirstLetter)
                    End If
                    GetLastLetter = Mid(RTF.Text, (.SelStart + .SelLength) - 1, 1)
                    If GetLastLetter <> "" Then AscLastLetter = Asc(GetLastLetter)
                    CheckLetter = Mid(RTF.Text, (.SelStart + .SelLength) + 1, 1)
                    If CheckLetter <> "" Then AscCheckLetter = Asc(CheckLetter)
                    If (GetFirstLetter <> "" Or GetFirstLetter <> " " Or AscFirstLetter <> 10 Or AscFirstLetter <> 13 Or AscFirstLetter <> 32) And _
                       (GetLastLetter <> "" Or GetLastLetter <> " " Or AscLastLetter <> 10 Or AscLastLetter <> 13 Or AscLastLetter <> 32) And _
                       (CheckLetter <> "_") Then
                        If AscCheckLetter <> "" Then
                            Select Case AscCheckLetter
                                Case 1 To 64
                                    .SelColor = sColor
                                Case 65 To 90
                                    .SelColor = vbBlack
                                Case 91 To 96
                                    .SelColor = sColor
                                Case 97 To 122
                                    .SelColor = vbBlack
                                Case 123 To 127
                                    .SelColor = sColor
                            End Select
                        Else
                            .SelColor = sColor
                        End If
                    Else
                        .SelColor = vbBlack
                    End If
                    .SelFontSize = SFontsize
                    .SelBold = Sbold
                    .SelItalic = Sitalic
                    .SelBullet = Sbullet
                End With
            End If
            search = search + Len(sWords)
        End If
    Loop
    With RTF
        .SelStart = Len(RTF.Text)
        .SelColor = vbBlack
        .SelFontSize = 8
        .SelBold = False
        .SelItalic = False
        .SelBullet = False
    End With
End Function

Public Sub SetDefaultColours(RTF As RichTextBox)
    iCursorPos = RTF.SelStart
    LockWindowUpdate frmMain.hwnd
    RTF.SelStart = 0
    RTF.SelLength = Len(RTF.Text)
    RTF.SelColor = &H0&
    For i = 0 To 174
        ColorDefaultWord RTF, KeywordArray(i).sKeyword, KeywordArray(i).sColour, 8, False, False, False
    Next
    For i = 0 To 8
        ColorDefaultWord RTF, OperatorArray(i).sOperator, OperatorArray(i).sColour, 8, False, False, False
    Next
    LockWindowUpdate 0&
    RTF.SelStart = iCursorPos
End Sub

Public Sub SetKeywords()
    ' Set the Keywords and Colours into the array
    KeywordArray(0).sKeyword = "ADD"
    KeywordArray(0).sColour = &HFF0000
    KeywordArray(1).sKeyword = "ALL"
    KeywordArray(1).sColour = &H808080
    KeywordArray(2).sKeyword = "ALTER"
    KeywordArray(2).sColour = &HFF0000
    KeywordArray(3).sKeyword = "AND"
    KeywordArray(3).sColour = &H808080
    KeywordArray(4).sKeyword = "ANY"
    KeywordArray(4).sColour = &H808080
    KeywordArray(5).sKeyword = "AS"
    KeywordArray(5).sColour = &HFF0000
    KeywordArray(6).sKeyword = "ASC"
    KeywordArray(6).sColour = &HFF0000
    KeywordArray(7).sKeyword = "AUTHORIZATION"
    KeywordArray(7).sColour = &HFF0000
    KeywordArray(8).sKeyword = "BACKUP"
    KeywordArray(8).sColour = &HFF0000
    KeywordArray(9).sKeyword = "BEGIN"
    KeywordArray(9).sColour = &HFF0000
    KeywordArray(10).sKeyword = "BETWEEN"
    KeywordArray(10).sColour = &H808080
    KeywordArray(11).sKeyword = "BREAK"
    KeywordArray(11).sColour = &HFF0000
    KeywordArray(12).sKeyword = "BROWSE"
    KeywordArray(12).sColour = &HFF0000
    KeywordArray(13).sKeyword = "BULK"
    KeywordArray(13).sColour = &HFF0000
    KeywordArray(14).sKeyword = "BY"
    KeywordArray(14).sColour = &HFF0000
    KeywordArray(15).sKeyword = "CASCADE"
    KeywordArray(15).sColour = &HFF0000
    KeywordArray(16).sKeyword = "CASE"
    KeywordArray(16).sColour = &HFF00FF
    KeywordArray(17).sKeyword = "CHECK"
    KeywordArray(17).sColour = &HFF0000
    KeywordArray(18).sKeyword = "CHECKPOINT"
    KeywordArray(18).sColour = &HFF0000
    KeywordArray(19).sKeyword = "CLOSE"
    KeywordArray(19).sColour = &HFF0000
    KeywordArray(20).sKeyword = "CLUSTERED"
    KeywordArray(20).sColour = &HFF0000
    KeywordArray(21).sKeyword = "COALESCE"
    KeywordArray(21).sColour = &HFF00FF
    KeywordArray(22).sKeyword = "COLLATE"
    KeywordArray(22).sColour = &HFF0000
    KeywordArray(23).sKeyword = "COLUMN"
    KeywordArray(23).sColour = &HFF0000
    KeywordArray(24).sKeyword = "COMMIT"
    KeywordArray(24).sColour = &HFF0000
    KeywordArray(25).sKeyword = "COMPUTE"
    KeywordArray(25).sColour = &HFF0000
    KeywordArray(26).sKeyword = "Constraint"
    KeywordArray(26).sColour = &HFF0000
    KeywordArray(27).sKeyword = "CONTAINS"
    KeywordArray(27).sColour = &HFF0000
    KeywordArray(28).sKeyword = "CONTAINSTABLE"
    KeywordArray(28).sColour = &HFF0000
    KeywordArray(29).sKeyword = "CONTINUE"
    KeywordArray(29).sColour = &HFF0000
    KeywordArray(30).sKeyword = "CONVERT"
    KeywordArray(30).sColour = &HFF00FF
    KeywordArray(31).sKeyword = "Create"
    KeywordArray(31).sColour = &HFF0000
    KeywordArray(32).sKeyword = "CROSS"
    KeywordArray(32).sColour = &H808080
    KeywordArray(33).sKeyword = "CURRENT"
    KeywordArray(33).sColour = &HFF0000
    KeywordArray(34).sKeyword = "CURRENT_DATE"
    KeywordArray(34).sColour = &HFF0000
    KeywordArray(35).sKeyword = "CURRENT_TIME"
    KeywordArray(35).sColour = &HFF0000
    KeywordArray(36).sKeyword = "CURRENT_TIMESTAMP"
    KeywordArray(36).sColour = &HFF00FF
    KeywordArray(37).sKeyword = "CURRENT_USER"
    KeywordArray(37).sColour = &HFF00FF
    KeywordArray(38).sKeyword = "CURSOR"
    KeywordArray(38).sColour = &HFF0000
    KeywordArray(39).sKeyword = "DATABASE"
    KeywordArray(39).sColour = &HFF0000
    KeywordArray(40).sKeyword = "DBCC"
    KeywordArray(40).sColour = &HFF0000
    KeywordArray(41).sKeyword = "DEALLOCATE"
    KeywordArray(41).sColour = &HFF0000
    KeywordArray(42).sKeyword = "DECLARE"
    KeywordArray(42).sColour = &HFF0000
    KeywordArray(43).sKeyword = "DEFAULT"
    KeywordArray(43).sColour = &HFF0000
    KeywordArray(44).sKeyword = "DELETE"
    KeywordArray(44).sColour = &HFF0000
    KeywordArray(45).sKeyword = "DENY"
    KeywordArray(45).sColour = &HFF0000
    KeywordArray(46).sKeyword = "DESC"
    KeywordArray(46).sColour = &HFF0000
    KeywordArray(47).sKeyword = "DISK"
    KeywordArray(47).sColour = &HFF0000
    KeywordArray(48).sKeyword = "DISTINCT"
    KeywordArray(48).sColour = &HFF0000
    KeywordArray(49).sKeyword = "DISTRIBUTED"
    KeywordArray(49).sColour = &HFF0000
    KeywordArray(50).sKeyword = "DOUBLE"
    KeywordArray(50).sColour = &HFF0000
    KeywordArray(51).sKeyword = "DROP"
    KeywordArray(51).sColour = &HFF0000
    KeywordArray(52).sKeyword = "DUMMY"
    KeywordArray(52).sColour = &HFF0000
    KeywordArray(53).sKeyword = "DUMP"
    KeywordArray(53).sColour = &HFF0000
    KeywordArray(54).sKeyword = "ELSE"
    KeywordArray(54).sColour = &HFF0000
    KeywordArray(55).sKeyword = "END"
    KeywordArray(55).sColour = &HFF0000
    KeywordArray(56).sKeyword = "ERRLVL"
    KeywordArray(56).sColour = &HFF0000
    KeywordArray(57).sKeyword = "ESCAPE"
    KeywordArray(57).sColour = &HFF0000
    KeywordArray(58).sKeyword = "EXCEPT"
    KeywordArray(58).sColour = &HFF0000
    KeywordArray(59).sKeyword = "EXEC"
    KeywordArray(59).sColour = &HFF0000
    KeywordArray(60).sKeyword = "EXECUTE"
    KeywordArray(60).sColour = &HFF0000
    KeywordArray(61).sKeyword = "EXISTS"
    KeywordArray(61).sColour = &H808080
    KeywordArray(62).sKeyword = "EXIT"
    KeywordArray(62).sColour = &HFF0000
    KeywordArray(63).sKeyword = "FETCH"
    KeywordArray(63).sColour = &HFF0000
    KeywordArray(64).sKeyword = "FILE"
    KeywordArray(64).sColour = &HFF0000
    KeywordArray(65).sKeyword = "FILLFACTOR"
    KeywordArray(65).sColour = &HFF0000
    KeywordArray(66).sKeyword = "FOR"
    KeywordArray(66).sColour = &HFF0000
    KeywordArray(67).sKeyword = "FOREIGN"
    KeywordArray(67).sColour = &HFF0000
    KeywordArray(68).sKeyword = "FREETEXT"
    KeywordArray(68).sColour = &HFF0000
    KeywordArray(69).sKeyword = "FREETEXTTABLE"
    KeywordArray(69).sColour = &HFF0000
    KeywordArray(70).sKeyword = "FROM"
    KeywordArray(70).sColour = &HFF0000
    KeywordArray(71).sKeyword = "FULL"
    KeywordArray(71).sColour = &HFF0000
    KeywordArray(72).sKeyword = "FUNCTION"
    KeywordArray(72).sColour = &HFF0000
    KeywordArray(73).sKeyword = "GOTO"
    KeywordArray(73).sColour = &HFF0000
    KeywordArray(74).sKeyword = "GRANT"
    KeywordArray(74).sColour = &HFF0000
    KeywordArray(75).sKeyword = "GROUP"
    KeywordArray(75).sColour = &HFF0000
    KeywordArray(76).sKeyword = "HAVING"
    KeywordArray(76).sColour = &HFF0000
    KeywordArray(77).sKeyword = "HOLDLOCK"
    KeywordArray(77).sColour = &HFF0000
    KeywordArray(78).sKeyword = "IDENTITY"
    KeywordArray(78).sColour = &HFF0000
    KeywordArray(79).sKeyword = "IDENTITY_INSERT"
    KeywordArray(79).sColour = &HFF0000
    KeywordArray(80).sKeyword = "IDENTITYCOL"
    KeywordArray(80).sColour = &HFF0000
    KeywordArray(81).sKeyword = "IF"
    KeywordArray(81).sColour = &HFF0000
    KeywordArray(82).sKeyword = "IN"
    KeywordArray(82).sColour = &H808080
    KeywordArray(83).sKeyword = "INDEX"
    KeywordArray(83).sColour = &HFF0000
    KeywordArray(84).sKeyword = "INNER"
    KeywordArray(84).sColour = &HFF0000
    KeywordArray(85).sKeyword = "INSERT"
    KeywordArray(85).sColour = &HFF0000
    KeywordArray(86).sKeyword = "INTERSECT"
    KeywordArray(86).sColour = &H808080
    KeywordArray(87).sKeyword = "INTO"
    KeywordArray(87).sColour = &HFF0000
    KeywordArray(88).sKeyword = "IS"
    KeywordArray(88).sColour = &HFF0000
    KeywordArray(89).sKeyword = "JOIN"
    KeywordArray(89).sColour = &H808080
    KeywordArray(90).sKeyword = "KEY"
    KeywordArray(90).sColour = &HFF0000
    KeywordArray(91).sKeyword = "KILL"
    KeywordArray(91).sColour = &HFF0000
    KeywordArray(92).sKeyword = "LEFT"
    KeywordArray(92).sColour = &HFF00FF
    KeywordArray(93).sKeyword = "LIKE"
    KeywordArray(93).sColour = &H808080
    KeywordArray(94).sKeyword = "LINENO"
    KeywordArray(94).sColour = &HFF0000
    KeywordArray(95).sKeyword = "LOAD"
    KeywordArray(95).sColour = &HFF0000
    KeywordArray(96).sKeyword = "NATIONAL"
    KeywordArray(96).sColour = &HFF0000
    KeywordArray(97).sKeyword = "NOCHECK"
    KeywordArray(97).sColour = &HFF0000
    KeywordArray(98).sKeyword = "NONCLUSTERED"
    KeywordArray(98).sColour = &HFF0000
    KeywordArray(99).sKeyword = "NOT"
    KeywordArray(99).sColour = &H808080
    KeywordArray(100).sKeyword = "NULL"
    KeywordArray(100).sColour = &H808080
    KeywordArray(101).sKeyword = "NULLIF"
    KeywordArray(101).sColour = &HFF00FF
    KeywordArray(102).sKeyword = "OF"
    KeywordArray(102).sColour = &HFF0000
    KeywordArray(103).sKeyword = "OFF"
    KeywordArray(103).sColour = &HFF0000
    KeywordArray(104).sKeyword = "OFFSETS"
    KeywordArray(104).sColour = &HFF0000
    KeywordArray(105).sKeyword = "ON"
    KeywordArray(105).sColour = &HFF0000
    KeywordArray(106).sKeyword = "OPEN"
    KeywordArray(106).sColour = &HFF0000
    KeywordArray(107).sKeyword = "OPENDATASOURCE"
    KeywordArray(107).sColour = &HFF0000
    KeywordArray(108).sKeyword = "OPENQUERY"
    KeywordArray(108).sColour = &HFF0000
    KeywordArray(109).sKeyword = "OPENROWSET"
    KeywordArray(109).sColour = &HFF0000
    KeywordArray(110).sKeyword = "OPENXML"
    KeywordArray(110).sColour = &HFF0000
    KeywordArray(111).sKeyword = "OPTION"
    KeywordArray(111).sColour = &HFF0000
    KeywordArray(112).sKeyword = "OR"
    KeywordArray(112).sColour = &H808080
    KeywordArray(113).sKeyword = "ORDER"
    KeywordArray(113).sColour = &HFF0000
    KeywordArray(114).sKeyword = "OUTER"
    KeywordArray(114).sColour = &H808080
    KeywordArray(115).sKeyword = "OVER"
    KeywordArray(115).sColour = &HFF0000
    KeywordArray(116).sKeyword = "PERCENT"
    KeywordArray(116).sColour = &HFF0000
    KeywordArray(117).sKeyword = "PLAN"
    KeywordArray(117).sColour = &HFF0000
    KeywordArray(118).sKeyword = "PRECISION"
    KeywordArray(118).sColour = &HFF0000
    KeywordArray(119).sKeyword = "PRIMARY"
    KeywordArray(119).sColour = &HFF0000
    KeywordArray(120).sKeyword = "PRINT"
    KeywordArray(120).sColour = &HFF0000
    KeywordArray(121).sKeyword = "PROC"
    KeywordArray(121).sColour = &HFF0000
    KeywordArray(122).sKeyword = "PROCEDURE"
    KeywordArray(122).sColour = &HFF0000
    KeywordArray(123).sKeyword = "PUBLIC"
    KeywordArray(123).sColour = &HFF0000
    KeywordArray(124).sKeyword = "RAISERROR"
    KeywordArray(124).sColour = &HFF0000
    KeywordArray(125).sKeyword = "READ"
    KeywordArray(125).sColour = &HFF0000
    KeywordArray(126).sKeyword = "READTEXT"
    KeywordArray(126).sColour = &HFF0000
    KeywordArray(127).sKeyword = "RECONFIGURE"
    KeywordArray(127).sColour = &HFF0000
    KeywordArray(128).sKeyword = "REFERENCES"
    KeywordArray(128).sColour = &HFF0000
    KeywordArray(129).sKeyword = "REPLICATION"
    KeywordArray(129).sColour = &HFF0000
    KeywordArray(130).sKeyword = "RESTORE"
    KeywordArray(130).sColour = &HFF0000
    KeywordArray(131).sKeyword = "RESTRICT"
    KeywordArray(131).sColour = &HFF0000
    KeywordArray(132).sKeyword = "RETURN"
    KeywordArray(132).sColour = &HFF0000
    KeywordArray(133).sKeyword = "REVOKE"
    KeywordArray(133).sColour = &HFF0000
    KeywordArray(134).sKeyword = "RIGHT"
    KeywordArray(134).sColour = &HFF00FF
    KeywordArray(135).sKeyword = "ROLLBACK"
    KeywordArray(135).sColour = &HFF0000
    KeywordArray(136).sKeyword = "ROWCOUNT"
    KeywordArray(136).sColour = &HFF0000
    KeywordArray(137).sKeyword = "ROWGUIDCOL"
    KeywordArray(137).sColour = &HFF0000
    KeywordArray(138).sKeyword = "RULE"
    KeywordArray(138).sColour = &HFF0000
    KeywordArray(139).sKeyword = "SAVE"
    KeywordArray(139).sColour = &HFF0000
    KeywordArray(140).sKeyword = "SCHEMA"
    KeywordArray(140).sColour = &HFF0000
    KeywordArray(141).sKeyword = "SELECT"
    KeywordArray(141).sColour = &HFF0000
    KeywordArray(142).sKeyword = "SESSION_USER"
    KeywordArray(142).sColour = &HFF00FF
    KeywordArray(143).sKeyword = "SET"
    KeywordArray(143).sColour = &HFF0000
    KeywordArray(144).sKeyword = "SETUSER"
    KeywordArray(144).sColour = &HFF0000
    KeywordArray(145).sKeyword = "SHUTDOWN"
    KeywordArray(145).sColour = &HFF0000
    KeywordArray(146).sKeyword = "SOME"
    KeywordArray(146).sColour = &H808080
    KeywordArray(147).sKeyword = "STATISTICS"
    KeywordArray(147).sColour = &HFF0000
    KeywordArray(148).sKeyword = "SUBSTR"
    KeywordArray(148).sColour = &HFF00FF
    KeywordArray(149).sKeyword = "SYSTEM_USER"
    KeywordArray(149).sColour = &HFF00FF
    KeywordArray(150).sKeyword = "TABLE"
    KeywordArray(150).sColour = &HFF0000
    KeywordArray(151).sKeyword = "TEXTSIZE"
    KeywordArray(151).sColour = &HFF0000
    KeywordArray(152).sKeyword = "THEN"
    KeywordArray(152).sColour = &HFF0000
    KeywordArray(153).sKeyword = "TO"
    KeywordArray(153).sColour = &HFF0000
    KeywordArray(154).sKeyword = "TOP"
    KeywordArray(154).sColour = &HFF0000
    KeywordArray(155).sKeyword = "TRAN"
    KeywordArray(155).sColour = &HFF0000
    KeywordArray(156).sKeyword = "TRANSACTION"
    KeywordArray(156).sColour = &HFF0000
    KeywordArray(157).sKeyword = "TRIGGER"
    KeywordArray(157).sColour = &HFF0000
    KeywordArray(158).sKeyword = "TRUNCATE"
    KeywordArray(158).sColour = &HFF0000
    KeywordArray(159).sKeyword = "TSEQUAL"
    KeywordArray(159).sColour = &HFF0000
    KeywordArray(160).sKeyword = "UNION"
    KeywordArray(160).sColour = &HFF0000
    KeywordArray(161).sKeyword = "UNIQUE"
    KeywordArray(161).sColour = &HFF0000
    KeywordArray(162).sKeyword = "UPDATE"
    KeywordArray(162).sColour = &HFF0000
    KeywordArray(163).sKeyword = "UPDATETEXT"
    KeywordArray(163).sColour = &HFF0000
    KeywordArray(164).sKeyword = "USE"
    KeywordArray(164).sColour = &HFF0000
    KeywordArray(165).sKeyword = "USER"
    KeywordArray(165).sColour = &HFF00FF
    KeywordArray(166).sKeyword = "VALUES"
    KeywordArray(166).sColour = &HFF0000
    KeywordArray(167).sKeyword = "VARYING"
    KeywordArray(167).sColour = &HFF0000
    KeywordArray(168).sKeyword = "VIEW"
    KeywordArray(168).sColour = &HFF0000
    KeywordArray(169).sKeyword = "WAITFOR"
    KeywordArray(169).sColour = &HFF0000
    KeywordArray(170).sKeyword = "WHEN"
    KeywordArray(170).sColour = &HFF0000
    KeywordArray(171).sKeyword = "WHERE"
    KeywordArray(171).sColour = &HFF0000
    KeywordArray(172).sKeyword = "WHILE"
    KeywordArray(172).sColour = &HFF0000
    KeywordArray(173).sKeyword = "WITH"
    KeywordArray(173).sColour = &HFF0000
    KeywordArray(174).sKeyword = "WRITETEXT"
    KeywordArray(174).sColour = &HFF0000

End Sub

Public Sub SetOperators()
    ' Set the Operators and Colours into the array
    OperatorArray(0).sOperator = "*"
    OperatorArray(0).sColour = &H808080
    OperatorArray(1).sOperator = "-"
    OperatorArray(1).sColour = &H808080
    OperatorArray(2).sOperator = "+"
    OperatorArray(2).sColour = &H808080
    OperatorArray(3).sOperator = "/"
    OperatorArray(3).sColour = &H808000
    OperatorArray(4).sOperator = "~"
    OperatorArray(4).sColour = &H808080
    OperatorArray(5).sOperator = "("
    OperatorArray(5).sColour = &H808000
    OperatorArray(6).sOperator = ")"
    OperatorArray(6).sColour = &H808000
    OperatorArray(7).sOperator = "<"
    OperatorArray(7).sColour = &H808000
    OperatorArray(8).sOperator = ">"
    OperatorArray(8).sColour = &H808000
End Sub
