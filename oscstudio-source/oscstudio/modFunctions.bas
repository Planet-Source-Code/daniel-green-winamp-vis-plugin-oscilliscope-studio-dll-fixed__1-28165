Attribute VB_Name = "modFunctions"
Public Enum VBConversion
    vbUpperCase = 1& ' Converts the string to uppercase characters.
    vbLowerCase = 2& ' Converts the string to lowercase characters.
    vbProperCase = 3& ' Converts the first letter of every word in string to uppercase.
    vbWide = 4& ' Converts narrow (single-byte) characters in string to wide (double-byte) characters. Applies to Far Eastlocales.
    vbNarrow = 8& ' Converts wide (double-byte) characters in string to narrow (single-byte) characters. Applies to Far East locales.
    vbKatakana = 16& ' Converts Hiragana characters in string to Katakana characters. Applies to Japan only.
    vbHiragana = 32& ' Converts Katakana characters in string to Hiragana characters. Applies to Japan only.
    vbUnicode = 64& ' Converts the string toUnicode using the default code page of the system.
    vbFromUnicode = 128& ' Converts the string from Unicode to the default code page of the system.
    vb7bit = 256&
    vb8Bit = 257&
    vbQuotedPrintable = 258&
    vbFromQuotedPrintable = 259&
    vbBase64 = 260&
    vbFromBase64 = 261&
    vbUUEncode = 262&
    vbUUDecode = 263&
    vbReversed = 264&
    vbByteArray = 265&
    vbWordArray = 266&
    vbEBCDIC = 267&
    vbFromEBCDIC = 268&
End Enum

Public Function UnLeft(String1, Length As Long)
    UnLeft = Mid(String1, Length + 1)
End Function

Public Function UnRight(String1, Length As Long)
    UnRight = Left(String1, Len(String1) - Length)
End Function

Public Function UnMid(String1, Start As Long, Optional Length)
    If IsMissing(Length) Then
        UnMid = Left(String1, Start - 1)
    Else
        UnMid = Left(String1, Start - 1) & Mid(String1, Start + Length)
    End If
End Function

' opposite of split function - combines array into
' string using delimiter character (default = <space>)
'
Public Function Glue(Array1(), Optional Delimiter)
    Dim delim As String, i As Integer, OutStr As String
    If IsMissing(Delimiter) Then delim = " " Else delim = Delimiter
    OutStr = vbNullString
    For i = LBound(Array1) To UBound(Array1)
        OutStr = OutStr & CStr(Array1(i)) & IIf(i < UBound(Array1), delim, vbNullString)
    Next
    Glue = OutStr
End Function

' returns true if string contains any characters with bit 7 set
' (ascii value >= 128)
'
Public Function IsBinary(String1) As Boolean
    Dim Bytes() As Byte, isBin As Boolean, i As Long
    Bytes() = StrConv(CStr(String1), vbFromUnicode)
    isBin = False
    i = LBound(Bytes)
    While i <= UBound(Bytes) And Not isBin
        If Bytes(i) >= 128 Then isBin = True
        i = i + 1
    Wend
    IsBinary = isBin
End Function

' trims all whitespace characters from left and right of string
'
Public Function TrimWS(ByVal String1 As String) As String
    TrimWS = RTrimWS(LTrimWS(String1))
End Function

' trims all whitespace characters from left of string
'
Public Function LTrimWS(ByVal String1 As String) As String
    While String1 > "" And Asc(String1 & " ") <= 32
        String1 = Mid(String1, 2)
    Wend
    LTrimWS = String1
End Function

' trims all whitespace characters from right of string
'
Public Function RTrimWS(ByVal String1 As String) As String
    While String1 > "" And Asc(Right(" " & String1, 1)) <= 32
        String1 = Left(String1, Len(String1) - 1)
    Wend
    RTrimWS = String1
End Function

' trims quotes from string
'
Public Function TrimQuotes(ByVal String1 As String) As String
    String1 = TrimWS(String1)
    While Left(String1, 1) = Chr(34)
        String1 = Mid(String1, 2)
    Wend
    While Right(String1, 1) = Chr(34)
        String1 = Left(String1, Len(String1) - 1)
    Wend
    TrimQuotes = String1
End Function

Function GetFilename(Filename)
Dim a, text
a = InStr(Filename, "\")
text = Mid(Filename, a + 1, Len(Filename))
Do Until InStr(text, "\") = 0
text = Mid(text, a, Len(text))
Loop
GetFilename = text
End Function

Function GetDir(Filename As String)
Dim a As Long, text
a = InStr(Filename, "\")
If a = 0 Then Exit Function
text = Mid(Filename, a + 1, Len(Filename))
Do Until InStr(text, "\") = 0
text = Mid(text, a, Len(text))
Loop
a = Len(text)
text = UnRight(Filename, a)
GetDir = text
End Function

Public Function ReadText(path As String) As String
On Error Resume Next
    Dim temptxt As String, x
    temptxt = ""
    Open path For Input As #1
    Do While Not EOF(1)
        Line Input #1, x
        temptxt = temptxt + x + vbCrLf
    Loop
    Close #1
    ReadText = temptxt
End Function

'===============================================
'Words.bas - string handling functions for words
'Author: Evan Sims         [esims@arcola-il.com]
'Based on a module by Kevin O'Brien
'Version - 1.2 (Sept. 1996 - Dec 1999)
'
'These functions deal with "words".
'Words = blank-delimited strings
'Blank = any combination of one or more spaces,
'        tabs, line feeds, or carriage returns.
'
'Examples:
'      word("find 3 in here", 3)     = "in"      3rd word
'     words("find 3 in here")        = 4         number of words
'     split("here's /s more", "/s")  = "more"    Returns words after split identifier (/s)
'   delWord("find 3 in here", 1, 2)  = "in here" delete 2 words, start at 1
'   midWord("find 3 in here", 1, 2)  = "find 3"  return 2 words, start at 1
'   wordPos("find 3 in here", "in")  = 3         word-number of "in"
' wordCount("find 3 in here", "in")  = 1         occurrences of word "in"
' wordIndex("find 3 in here", "in")  = 8         position of "in"
' wordIndex("find 3 in here", 3)     = 8         position of 3rd word
' wordIndex("find 3 in here", "3")   = 6         position of "3"
'wordLength("find 3 in here", 3)     = 2         length of 3rd word
'
'Difference between Instr() and wordIndex():
'     InStr("find 3 in here", "in")   = 2
' wordIndex("find 3 in here", "in")   = 8
'
'     InStr("find 3 in here", "her")  = 11
' wordIndex("find 3 in here", "her")  = 0
'===============================================

Public Function Word(ByVal sSource As String, _
                                 n As Long) As String
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
Const SP    As String = " "
Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim x       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

sSource = CSpace(sSource)

'find the nth word
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function

Public Function Words(ByVal sSource As String) As Long
'=================================================
' Words returns the number of words in a string
' Usage:
'    Words("red blue green")   3
'=================================================
Const SP    As String = " "
Dim lSource As Long    'length of sSource
Dim pointer As Long    'start parameter of Instr()
Dim pos     As Long    'position of target in InStr()
Dim x       As Long    'word count

sSource = CSpace(sSource)
lSource = Len(sSource)
If lSource = 0 Then Exit Function

'count words
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'no more words
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
If Mid$(sSource, lSource, 1) = SP Then x = x - 1 'adjust if trailing space
Words = x
End Function

Public Function WordCount(ByVal sSource As String, _
                                sTarget As String) As Long
'=====================================================
' WordCount returns the number of times that
' word, sTarget, is found in sSource.
' Usage:
'    WordCount("a rose is a rose", "rose")     2
'=================================================
Const SP    As String = " "
Dim pointer As Long    'start parameter of Instr()
Dim lSource As Long    'length of sSource
Dim lTarget As Long    'length of sTarget
Dim pos     As Long    'position of target in InStr()
Dim x       As Long    'word count

lTarget = Len(sTarget)
lSource = Len(sSource)
sSource = CSpace(sSource)


'find target word
pointer = 1
Do While Mid$(sSource, pointer, 1) = SP       'skip consecutive spaces
   pointer = pointer + 1
Loop
If pointer > lSource Then Exit Function       'sSource contains no words

Do                                            'find position of sTarget
   pos = InStr(pointer, sSource, sTarget)
   If pos = 0 Then Exit Do                    'string not found
   If Mid$(sSource, pos + lTarget, 1) = SP _
   Or pos + lTarget > lSource Then            'must be a word
      If pos = 1 Then
         x = x + 1                            'word found
      ElseIf Mid$(sSource, pos - 1, 1) = SP Then
         x = x + 1                            'word found
      End If
   End If
   pointer = pos + lTarget
Loop
WordCount = x

End Function

Public Function WordPos(ByVal sSource As String, _
                              sTarget As String) As Long
'=====================================================
' WordPos returns the word number of the
' word, sTarget, in sSource.
' Usage:
'    WordPos("red blue green", "blue")    2
'=================================================
Const SP       As String = " "
Dim pointer    As Long    'start parameter of Instr()
Dim lSource    As Long    'length of sSource
Dim lTarget    As Long    'length of sTarget
Dim lPosTarget As Long    'position of target-word
Dim pos        As Long    'position of target in InStr()
Dim x          As Long    'word count

lTarget = Len(sTarget)
lSource = Len(sSource)
sSource = CSpace(sSource)


'find target word
pointer = 1
Do While Mid$(sSource, pointer, 1) = SP       'skip consecutive spaces
   pointer = pointer + 1
Loop
If pointer > lSource Then Exit Function       'sSource contains no words

Do                                            'find position of sTarget
   pos = InStr(pointer, sSource, sTarget)
   If pos = 0 Then Exit Function              'string not found
   If Mid$(sSource, pos + lTarget, 1) = SP _
   Or pos + lTarget > lSource Then            'must be a word
      If pos = 1 Then Exit Do                 'word found
      If Mid$(sSource, pos - 1, 1) = SP Then Exit Do
   End If
   pointer = pos + lTarget
Loop


'count words until position of sTarget
lPosTarget = pos                             'save position of sTarget
pointer = 1
x = 1
Do
   Do While Mid$(sSource, pointer, 1) = SP   'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If pointer >= lPosTarget Then Exit Do     'all words have been counted
   pos = InStr(pointer, sSource, SP)         'find next space
   If pos = 0 Then Exit Do                   'no more words
   x = x + 1                                 'increment word count
   pointer = pos + 1                         'start of next word
Loop
WordPos = x
End Function

Public Function WordIndex(ByVal sSource As String, _
                                vTarget As Variant) As Long
'===========================================================
' WordIndex returns the byte position of vTarget in sSource.
' vTarget can be a word-number or a string.
' Usage:
'   WordIndex("two plus 2 is four", 2)       5
'   WordIndex("two plus 2 is four", "2")    10
'   WordIndex("two plus 2 is four", "two")   1
'===========================================================
Const SP    As String = " "
Dim sTarget As String  'vTarget converted to String
Dim lTarget As Long    'vTarget converted to Long, or length of sTarget
Dim lSource As Long    'length of sSource
Dim pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim x       As Long    'word counter

lSource = Len(sSource)
sSource = CSpace(sSource)

If VarType(vTarget) = vbString Then GoTo strIndex
If Not IsNumeric(vTarget) Then Exit Function
lTarget = CLng(vTarget)                       'convert to Long

'find byte position of lTarget (word number)
x = 1
pointer = 1


Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   
   If x = lTarget Then                         'word-number of Target
      If pointer > lSource Then Exit Do        'beyond end of sSource
      WordIndex = pointer                      'position of word
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
   pointer = pos + 1
Loop

Exit Function
strIndex:
sTarget = CStr(vTarget)
lTarget = Len(sTarget)
If lTarget = 0 Then Exit Function              'nothing to count

'find byte position of sTarget (string)
pointer = 1
Do
   pos = InStr(pointer, sSource, sTarget)
   If pos = 0 Then Exit Do
   If Mid$(sSource, pos + lTarget, 1) = SP _
   Or pos + lTarget > lSource Then
      If pos = 1 Then Exit Do
      If Mid$(sSource, pos - 1, 1) = SP Then Exit Do
   End If
   pointer = pos + lTarget
Loop

WordIndex = pos

End Function

Public Function WordLength(ByVal sSource As String, _
                                       n As Long) As Long
'=========================================================
' Wordlength returns the length of the nth word in sSource
' Usage:
'    WordLength("red blue green", 2)    4
'=========================================================
Const SP    As String = " "
Dim lSource As Long   'length of sSource
Dim pointer As Long   'start parameter Instr()
Dim pos     As Long   'position of target with InStr()
Dim x       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

sSource = CSpace(sSource)
lSource = Len(sSource)

'find the nth word
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = lSource + 1      '   or if its the last word
      WordLength = lEnd - pointer
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop

End Function

Public Function DelWord(ByVal sSource As String, _
                                    n As Long, _
                      Optional vWords As Variant) As String
'===========================================================
' DelWord deletes from sSource, starting with the
' nth word for a length of vWords words.
' If vWords is omitted, all words from the nth word on are
' deleted.
' Usage:
'   DelWord("now is not the time", 3)     "now is"
'   DelWord("now is not the time", 3, 1)  "now is the time"
'===========================================================
Const SP    As String = " "
Dim lWords  As Long    'length of sTarget
Dim lSource As Long    'length of sSource
Dim pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim x       As Long    'word counter
Dim lStart  As Long    'position of word n
Dim lEnd    As Long    'position of space after last word

lSource = Len(sSource)
DelWord = sSource
sSource = CSpace(sSource)
If IsMissing(vWords) Then
   lWords = -1
ElseIf IsNumeric(vWords) Then
   lWords = CLng(vWords)
Else
   Exit Function
End If

If n = 0 Or lWords = 0 Then Exit Function      'nothing to delete

'find position of n
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lStart = pointer
      If lWords < 0 Then Exit Do
   End If
   
   If lWords > 0 Then                          'lWords was provided
      If x = n + lWords - 1 Then               'find pos of last word
         lEnd = InStr(pointer, sSource, SP)    'pos of space at end of word
         Exit Do                               'word found, done
      End If
   End If
   
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
If lStart = 0 Then Exit Function
If lEnd = 0 Then
   DelWord = Trim$(Left$(sSource, lStart - 1))
Else
   DelWord = Trim$(Left$(sSource, lStart - 1) & Mid$(sSource, lEnd + 1))
End If
End Function

Public Function MidWord(ByVal sSource As String, _
                                    n As Long, _
                      Optional vWords As Variant) As String
'===========================================================
' MidWord returns a substring sSource, starting with the
' nth word for a length of vWords words.
' If vWords is omitted, all words from the nth word on are
' returned.
' Usage:
'   MidWord("now is not the time", 3)     "not the time"
'   MidWord("now is not the time", 3, 2)  "not the"
'===========================================================
Const SP    As String = " "
Dim lWords  As Long    'vWords converted to long
Dim lSource As Long    'length of sSource
Dim pointer As Long    'start parameter of InStr()
Dim pos     As Long    'position of target in InStr()
Dim x       As Long    'word counter
Dim lStart  As Long    'position of word n
Dim lEnd    As Long    'position of space after last word

lSource = Len(sSource)
sSource = CSpace(sSource)
If IsMissing(vWords) Then
   lWords = -1
ElseIf IsNumeric(vWords) Then
   lWords = CLng(vWords)
Else
   Exit Function
End If

If n = 0 Or lWords = 0 Then Exit Function              'nothing to delete

'find position of n
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lStart = pointer
      If lWords < 0 Then Exit Do               'include rest of sSource
   End If
   
   If lWords > 0 Then                          'lWords was provided
      If x = n + lWords - 1 Then               'find pos of last word
         lEnd = InStr(pointer, sSource, SP)    'pos of space at end of word
         Exit Do                               'word found, done
      End If
   End If
   
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
If lStart = 0 Then Exit Function
If lEnd = 0 Then
   MidWord = Trim$(Mid$(sSource, lStart))
Else
   MidWord = Trim$(Mid$(sSource, lStart, lEnd - lStart))
End If
End Function

Public Function CSpace(sSource As String) As String
'==================================================
'CSpace converts blank characters
'(ascii: 9,10,13,160) to space (32)
'
'  cSpace("a" & vbTab   & "b")  "a b"
'  cSpace("a" & vbCrlf  & "b")  "a  b"
'==================================================
Dim pointer   As Long
Dim pos       As Long
Dim x         As Long
Dim iSpace(3) As Integer

' define blank characters
iSpace(0) = 9    'Horizontal Tab
iSpace(1) = 10   'Line Feed
iSpace(2) = 13   'Carriage Return
iSpace(3) = 160  'Hard Space

CSpace = sSource
For x = 0 To UBound(iSpace) ' replace all blank characters with space
   pointer = 1
   Do
      pos = InStr(pointer, CSpace, Chr$(iSpace(x)))
      If pos = 0 Then Exit Do
      Mid$(CSpace, pos, 1) = " "
      pointer = pos + 1
   Loop
Next x

End Function

Public Function SplitString(iSource As String, iTarget As String, Optional BeforeTarget As Boolean = False) As String
'==================================================
'Returns the characters before or after the split
'identifier. By default will return text after id,
'set BeforeTarget as true to return the text before
'it.
'==================================================
If BeforeTarget = True Then
   SplitString = DelWord(iSource, WordPos(iSource, iTarget))
Else
   SplitString = DelWord(iSource, 1, WordPos(iSource, iTarget))
End If

End Function

Public Function ReadFile(path As String) As String
'diz function returns the data from the given file
On Error Resume Next
    Dim temptxt As String
    temptxt = ""
    Dim free
    free = FreeFile
    Open path For Input As #free
    Do While Not EOF(free)
        Line Input #free, x
        temptxt = temptxt + x + vbCrLf
        DoEvents
    Loop
    Close #free
    ReadFile = temptxt
End Function

Function GetFilename1(Filename)
Dim a, text
a = InStr(Filename, "\")
text = Mid(Filename, a + 1, Len(Filename))
Do Until InStr(text, "\") = 0
text = Mid(text, a, Len(text))
Loop
GetFilename1 = text
End Function

Function GetDir1(Filename As String)
Dim a As Long, text
a = InStr(Filename, "\")
If a = 0 Then Exit Function
text = Mid(Filename, a + 1, Len(Filename))
Do Until InStr(text, "\") = 0
a = InStr(text, "\")
text = Mid(text, (a + 1), Len(text))
Loop
GetDir1 = text
End Function


