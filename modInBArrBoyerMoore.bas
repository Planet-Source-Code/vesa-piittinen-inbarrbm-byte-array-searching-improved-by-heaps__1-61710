Attribute VB_Name = "modInBArrBoyerMoore"
' InBArrBM - In Byte Array, Boyer-Moore optimized version
' Version 1.0 made by Vesa Piittinen, 2005-07-15
'
' About
' -----
' This is a function that can be used to search byte arrays.
' It is very fast, it can beat InStr in both BinaryCompare and
' TextCompare. Noticeable is that the function is made 100% VB6,
' no external API or TLB file used. You MUST enable all Advanced
' Optimizations before you compile the program or the function
' will not work in full speed!
'
' Notes about speed: the function can't beat InStr with short
' keywords, because one or two character long keywords are
' searched using brute force (there is no point to use Boyer-Moore
' algorithm with short keywords!)
'
' History
' -------
' 1.0 [2005-07-15]
' - initial release
' - optimized non-API, non-TLB, 100% Unicode compliant version

Option Explicit
' InBArr Boyer-Moore with both Unicode and ANSI support
Public Function InBArrBM(ByRef ByteArray() As Byte, ByRef KeyWord As String, Optional StartPos As Long = 0, Optional Compare As Long = vbBinaryCompare, Optional IsUnicode As Boolean = True) As Long
    ' optimization with Static; so there might be no need to redim these each time
    Static KeyBuffer() As Byte, KeyBufferU() As Byte, KeyDelta(65535) As Long
    Static OldKeyWord As String, OldCompare As Long, OldIsUnicode As Boolean
    ' other helper variables in use
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, blnComp As Boolean
    Dim KeyLenB As Long, KeyLen As Long, KeyLen1 As Long, KeyLen2 As Long, KeyLen3 As Long, KeyUpper As Long
    Dim FirstKeyByte As Byte, LastKeyByte As Byte, TempByte As Byte, TempByte3 As Byte
    Dim FirstKeyByte2 As Byte, LastKeyByte2 As Byte, TempByte2 As Byte, TempByte4 As Byte
    Dim FirstKeyByteU As Byte, LastKeyByteU As Byte
    Dim FirstKeyByte2U As Byte, LastKeyByte2U As Byte
    ' if we have no input, leave
    If (Not ByteArray) = True Then InBArrBM = -1: Exit Function
    ' get length of the key...
    KeyLenB = LenB(KeyWord)
    ' if keyword size is nothing, then we can just return the start position; nothing fits anywhere
    If KeyLenB = 0 Then InBArrBM = StartPos: Exit Function
    ' ...and optimize for the most needed values
    KeyLen = KeyLenB - 1
    KeyLen1 = KeyLenB - 2
    KeyLen2 = KeyLenB - 3
    KeyLen3 = KeyLenB - 4
    ' here we process keyword
    Do
        ' first check if there is a need to update the keyword
        ' if the user keeps using the same keyword, there is no point
        ' reseting keyword data over and over again
        If OldIsUnicode = IsUnicode Then
            ' check if compare mode has stayed the same
            If OldCompare = Compare Then
                ' check the keyword length has stayed the same
                If KeyLenB = LenB(OldKeyWord) Then
                    ' check the keyword byte-by-byte... suprised to see InStr here?
                    If InStr(1, KeyWord, OldKeyWord, vbBinaryCompare) = 1 Then Exit Do
                End If
            End If
        End If
        ' correct these settings
        OldKeyWord = KeyWord
        OldCompare = Compare
        OldIsUnicode = IsUnicode
        ' this takes the most time and is the main reason to keyword optimization
        ' it is a Boyer-Moore delta1 optimization table
        If KeyLen > 3 Then
            If IsUnicode Then
                ' in Unicode mode we have 65536 different characters...
                For A = 0 To 65535
                    KeyDelta(A) = KeyLenB
                Next A
            Else
                ' and in ANSI mode we have the lovely 256 characters
                B = (KeyLenB) \ 2
                For A = 0 To 255
                    KeyDelta(A) = B
                Next A
            End If
            ' now process the keyword into a form we can use
            If Compare = vbBinaryCompare Then
                ' convert to binarycompare byte array
                KeyBuffer = KeyWord
                ' finish up the Boyer-Moore delta1 optimization table
                If IsUnicode Then
                    ' Unicode
                    For A = 0 To KeyLen3 Step 2
                        B = KeyLen1 - A
                        KeyDelta(CLng(KeyBuffer(A)) Or (CLng(KeyBuffer(A + 1)) * &H100)) = B
                    Next A
                Else
                    ' ANSI
                    For A = 0 To KeyLen3 Step 2
                        B = (KeyLen1 - A) \ 2
                        KeyDelta(CLng(KeyBuffer(A))) = B
                    Next A
                End If
            Else
                ' convert to textcompare byte arrays
                KeyBuffer = LCase$(KeyWord)
                KeyBufferU = UCase$(KeyWord)
                ' finish up the Boyer-Moore delta1 optimization table
                If IsUnicode Then
                    ' Unicode
                    For A = 0 To KeyLen3 Step 2
                        B = KeyLen1 - A
                        C = A + 1
                        ' we take both lower case and upper case into account
                        If (KeyBuffer(A) = KeyBufferU(A)) And (KeyBuffer(C) = KeyBufferU(C)) Then
                            KeyDelta(CLng(KeyBuffer(A)) Or (CLng(KeyBuffer(C)) * &H100)) = KeyLenB
                        Else
                            KeyDelta(CLng(KeyBuffer(A)) Or (CLng(KeyBuffer(C)) * &H100)) = KeyLenB
                            KeyDelta(CLng(KeyBufferU(A)) Or (CLng(KeyBufferU(C)) * &H100)) = KeyLenB
                        End If
                    Next A
                Else
                    ' ANSI
                    For A = 0 To KeyLen3 Step 2
                        B = (KeyLen1 - A) \ 2
                        ' we tale both lower case and upper case into account
                        If (KeyBuffer(A) = KeyBufferU(A)) Then
                            KeyDelta(CLng(KeyBuffer(A))) = B
                        Else
                            KeyDelta(CLng(KeyBuffer(A))) = B
                            KeyDelta(CLng(KeyBufferU(A))) = B
                        End If
                    Next A
                End If
            End If
            Exit Do
        Else ' short key, we do brute force and not Boyer-Moore
            If Compare = vbBinaryCompare Then
                KeyBuffer = KeyWord
            Else
                KeyBuffer = LCase$(KeyWord)
                KeyBufferU = UCase$(KeyWord)
            End If
            Exit Do
        End If
    Loop
    ' correct start position silently
    If StartPos < 0 Then StartPos = 0
    ' check if in Unicode or in ANSI mode
    If IsUnicode Then
        ' can't do this before this point: check if keyword is longer than the searchtext
        If KeyLen2 > UBound(ByteArray) Then InBArrBM = -1: Exit Function
        ' make sure we have a valid startpos (start from an even position)
        If (StartPos Mod 2) = 1 Then StartPos = StartPos - (StartPos Mod 2) + 2
        ' make sure startpos is within a valid range
        If StartPos > UBound(ByteArray) - KeyLen1 Then InBArrBM = -1: Exit Function
        If Compare = vbBinaryCompare Then
            If KeyLen > 3 Then ' Boyer-Moore
                ' optimization: get the last character
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByte2 = KeyBuffer(KeyLen)
                ' where to begin?
                A = StartPos + KeyLen1
                ' last byte?
                B = UBound(ByteArray)
                ' loop until no match
                Do
                    ' optimization
                    TempByte = ByteArray(A)
                    TempByte2 = ByteArray(A + 1)
                    ' comparing a byte against a byte seems to be faster
                    ' than comparing a byte array item against a byte
                    If TempByte = LastKeyByte Then
                        If TempByte2 = LastKeyByte2 Then
                            ' check all the rest of the characters
                            ' and compare them against the searchtext
                            D = A - 2
                            For C = KeyLen3 To 0 Step -2
                                TempByte = ByteArray(D)
                                TempByte2 = ByteArray(D + 1)
                                TempByte3 = KeyBuffer(C)
                                TempByte4 = KeyBuffer(C + 1)
                                If Not (TempByte = TempByte3) And (TempByte2 = TempByte4) Then Exit For
                                D = D - 2
                            Next C
                            ' check if the loop ran to the bitter end
                            If C < 0 Then
                                ' a match is found!
                                InBArrBM = D + 2
                                Exit Function
                            Else
                                ' jump X bytes depending on what character we are
                                ' checking; this is the main idea of Boyer-Moore
                                A = D + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                            End If
                        Else
                            ' jump X bytes depending on what character we are
                            ' checking; this is the main idea of Boyer-Moore
                            A = A + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                        End If
                    Else
                        ' jump X bytes depending on what character we are
                        ' checking; this is the main idea of Boyer-Moore
                        A = A + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                    End If
                Loop Until A > B
            ElseIf KeyLen > 1 Then ' brute force
                ' optimization: get the first and the last character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByte2 = KeyBuffer(1)
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByte2 = KeyBuffer(KeyLen)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) - KeyLen1 Step 2
                    ' optimization: check the lower byte of first keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Then
                        ' optimization: check the lower byte of last keyword char
                        TempByte = ByteArray(A + KeyLen1)
                        If TempByte = LastKeyByte Then
                            ' optimization: check the upper byte of first char
                            TempByte2 = ByteArray(A + 1)
                            If TempByte2 = FirstKeyByte2 Then
                                ' optimization: check the upper byte of last char
                                TempByte2 = ByteArray(A + KeyLen)
                                If TempByte2 = LastKeyByte2 Then
                                    ' we have a match!
                                    InBArrBM = A
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next A
            Else ' brute force
                ' optimization: get the only character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByte2 = KeyBuffer(1)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) Step 2
                    ' optimization: check the lower byte of only keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Then
                        ' optimization: check the upper byte of only char
                        TempByte2 = ByteArray(A + 1)
                        If TempByte2 = FirstKeyByte2 Then
                            ' we have a match!
                            InBArrBM = A
                            Exit Function
                        End If
                    End If
                Next A
            End If
        Else 'vbTextCompare
            If KeyLen > 3 Then ' Boyer-Moore
                ' optimization: get the last character
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByte2 = KeyBuffer(KeyLen)
                LastKeyByteU = KeyBufferU(KeyLen1)
                LastKeyByte2U = KeyBufferU(KeyLen)
                ' where to begin?
                A = StartPos + KeyLen1
                ' last byte?
                B = UBound(ByteArray)
                Do
                    ' optimization
                    TempByte = ByteArray(A)
                    TempByte2 = ByteArray(A + 1)
                    ' comparing a byte against a byte seems to be faster
                    ' than comparing a byte array item against a byte
                    If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                        If TempByte2 = LastKeyByte2 Or TempByte2 = LastKeyByte2U Then
                            ' check all the rest of the characters
                            ' and compare them against the searchtext
                            D = A - 2
                            For C = KeyLen3 To 0 Step -2
                                TempByte = ByteArray(D)
                                TempByte2 = ByteArray(D + 1)
                                TempByte3 = KeyBuffer(C)
                                TempByte4 = KeyBuffer(C + 1)
                                blnComp = ((TempByte = TempByte3) And (TempByte2 = TempByte4))
                                If Not blnComp Then
                                    TempByte3 = KeyBufferU(C)
                                    TempByte4 = KeyBufferU(C + 1)
                                    blnComp = ((TempByte = TempByte3) And (TempByte2 = TempByte4))
                                    If Not blnComp Then Exit For
                                End If
                                D = D - 2
                            Next C
                            ' check if the loop ran to the bitter end
                            If C < 0 Then
                                ' a match is found!
                                InBArrBM = D + 2
                                Exit Function
                            Else
                                ' jump X bytes depending on what character we are
                                ' checking; this is the main idea of Boyer-Moore
                                A = D + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                            End If
                        Else
                            ' jump X bytes depending on what character we are
                            ' checking; this is the main idea of Boyer-Moore
                            A = A + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                        End If
                    Else
                        ' jump X bytes depending on what character we are
                        ' checking; this is the main idea of Boyer-Moore
                        A = A + KeyDelta(CLng(TempByte) Or (CLng(TempByte2) * &H100))
                    End If
                Loop Until A > B
            ElseIf KeyLen > 1 Then ' brute force
                ' optimization: get the first and the last character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByte2 = KeyBuffer(1)
                FirstKeyByteU = KeyBufferU(0)
                FirstKeyByte2U = KeyBufferU(1)
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByte2 = KeyBuffer(KeyLen)
                LastKeyByteU = KeyBufferU(KeyLen1)
                LastKeyByte2U = KeyBufferU(KeyLen)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) - KeyLen1 Step 2
                    ' optimization: check the lower byte of first keyword char
                    TempByte = ByteArray(A)
                    If (TempByte = FirstKeyByte Or TempByte = FirstKeyByteU) Then
                        ' optimization: check the lower byte of last keyword char
                        TempByte = ByteArray(A + KeyLen1)
                        If (TempByte = LastKeyByte Or TempByte = LastKeyByteU) Then
                            ' optimization: check the upper byte of first char
                            TempByte2 = ByteArray(A + 1)
                            If (TempByte2 = FirstKeyByte2 Or TempByte2 = FirstKeyByte2U) Then
                                ' optimization: check the upper byte of last char
                                TempByte2 = ByteArray(A + KeyLen)
                                If (TempByte2 = LastKeyByte2 Or TempByte2 = LastKeyByte2U) Then
                                    ' we have a match!
                                    InBArrBM = A
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next A
            Else ' brute force
                ' optimization: get the only character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByte2 = KeyBuffer(1)
                FirstKeyByteU = KeyBufferU(0)
                FirstKeyByte2U = KeyBufferU(1)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) Step 2
                    ' optimization: check the lower byte of only keyword char
                    TempByte = ByteArray(A)
                    If (TempByte = FirstKeyByte Or TempByte = FirstKeyByteU) Then
                        ' optimization: check the upper byte of only char
                        TempByte2 = ByteArray(A + 1)
                        If (TempByte2 = FirstKeyByte2 Or TempByte2 = FirstKeyByte2U) Then
                            ' we have a match!
                            InBArrBM = A
                            Exit Function
                        End If
                    End If
                Next A
            End If
        End If
    Else 'ANSI
        ' we need this in ANSI mode
        KeyUpper = KeyLen1 \ 2
        ' check keyword isn't longer than the searchtext
        If KeyUpper > UBound(ByteArray) Then InBArrBM = -1: Exit Function
        ' make sure we have a valid startpos
        If StartPos > UBound(ByteArray) - KeyUpper Then InBArrBM = -1: Exit Function
        If Compare = vbBinaryCompare Then
            If KeyLen > 3 Then ' Boyer-Moore
                ' optimization: get the last character
                LastKeyByte = KeyBuffer(KeyLen1)
                ' where to begin?
                A = StartPos + KeyUpper
                ' last byte?
                B = UBound(ByteArray)
                ' loop until no match
                Do
                    ' optimization
                    TempByte = ByteArray(A)
                    ' comparing a byte against a byte seems to be faster
                    ' than comparing a byte array item against a byte
                    If TempByte = LastKeyByte Then
                        ' check all the rest of the characters
                        ' and compare them against the searchtext
                        D = A - 1
                        For C = KeyLen3 To 0 Step -2
                            TempByte = ByteArray(D)
                            TempByte2 = KeyBuffer(C)
                            If Not (TempByte = TempByte2) Then Exit For
                            D = D - 1
                        Next C
                        ' check if the loop ran to the bitter end
                        If C < 0 Then
                            ' a match is found!
                            InBArrBM = D + 1
                            Exit Function
                        Else
                            ' jump X bytes depending on what character we are
                            ' checking; this is the main idea of Boyer-Moore
                            A = D + KeyDelta(TempByte)
                        End If
                    Else
                        ' jump X bytes depending on what character we are
                        ' checking; this is the main idea of Boyer-Moore
                        A = A + KeyDelta(TempByte)
                    End If
                Loop Until A > B
            ElseIf KeyLen > 1 Then ' brute force
                ' optimization: get the first and the last character
                FirstKeyByte = KeyBuffer(0)
                LastKeyByte = KeyBuffer(KeyLen1)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) - KeyUpper
                    ' optimization: check the first keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Then
                        ' optimization: check the last keyword char
                        TempByte2 = ByteArray(A + KeyUpper)
                        If TempByte2 = LastKeyByte Then
                            ' we have a match!
                            InBArrBM = A
                            Exit Function
                        End If
                    End If
                Next A
            Else ' brute force
                ' optimization: get the only character
                FirstKeyByte = KeyBuffer(0)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray)
                    ' optimization: check the only keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Then
                        ' we have a match!
                        InBArrBM = A
                        Exit Function
                    End If
                Next A
            End If
        Else 'vbTextCompare
            If KeyLen > 3 Then ' Boyer-Moore
                ' optimization: get the last character
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByteU = KeyBufferU(KeyLen1)
                ' where to begin?
                A = StartPos + KeyUpper
                ' last byte?
                B = UBound(ByteArray)
                ' loop until no match
                Do
                    ' optimization
                    TempByte = ByteArray(A)
                    ' comparing a byte against a byte seems to be faster
                    ' than comparing a byte array item against a byte
                    If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                        ' check all the rest of the characters
                        ' and compare them against the searchtext
                        D = A - 1
                        For C = KeyLen3 To 0 Step -2
                            TempByte = ByteArray(D)
                            TempByte2 = KeyBuffer(C)
                            blnComp = (TempByte = TempByte2)
                            If Not blnComp Then
                                TempByte2 = KeyBufferU(C)
                                blnComp = (TempByte = TempByte2)
                                If Not blnComp Then Exit For
                            End If
                            D = D - 1
                        Next C
                        ' check if the loop ran to the bitter end
                        If C < 0 Then
                            ' a match is found!
                            InBArrBM = D + 1
                            Exit Function
                        Else
                            ' jump X bytes depending on what character we are
                            ' checking; this is the main idea of Boyer-Moore
                            A = D + KeyDelta(TempByte)
                        End If
                    Else
                        ' jump X bytes depending on what character we are
                        ' checking; this is the main idea of Boyer-Moore
                        A = A + KeyDelta(TempByte)
                    End If
                Loop Until A > B
            ElseIf KeyLen > 1 Then ' brute force
                ' optimization: get the first and the last character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByteU = KeyBufferU(0)
                LastKeyByte = KeyBuffer(KeyLen1)
                LastKeyByteU = KeyBufferU(KeyLen1)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray) - KeyUpper
                    ' optimization: check the first keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                        ' optimization: check the last keyword char
                        TempByte2 = ByteArray(A + KeyUpper)
                        If TempByte2 = LastKeyByte Or TempByte2 = LastKeyByteU Then
                            ' we have a match!
                            InBArrBM = A
                            Exit Function
                        End If
                    End If
                Next A
            Else ' brute force
                ' optimization: get the only character
                FirstKeyByte = KeyBuffer(0)
                FirstKeyByteU = KeyBufferU(0)
                ' loop through all characters from startpos to end
                For A = StartPos To UBound(ByteArray)
                    ' optimization: check the only keyword char
                    TempByte = ByteArray(A)
                    If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                        ' we have a match!
                        InBArrBM = A
                        Exit Function
                    End If
                Next A
            End If
        End If
    End If
    InBArrBM = -1
End Function
