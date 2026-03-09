Attribute VB_Name = "PdfTXT"
Option Explicit

Public Function PDF_ExtractText(ByVal sFilePath As String) As String
    On Error GoTo Fail
    Dim bFile() As Byte
    Dim sRaw    As String
    Dim result  As String

    bFile = PDF_ReadFileBytes(sFilePath)
    If UBound(bFile) < 4 Then GoTo Fail
    If bFile(0) <> 37 Or bFile(1) <> 80 Or _
       bFile(2) <> 68 Or bFile(3) <> 70 Then GoTo Fail  ' %PDF magic bytes

    sRaw = PDF_BytesToLatin1(bFile)

    If InStr(1, sRaw, "/Encrypt", vbBinaryCompare) > 0 Then GoTo Fail

    result = PDF_ProcessAllStreams(bFile, sRaw)
    PDF_ExtractText = result
    Exit Function
Fail:
    PDF_ExtractText = ""
End Function

Private Function PDF_ReadFileBytes(ByVal sPath As String) As Byte()
    Dim iFile   As Integer
    Dim bData() As Byte
    iFile = FreeFile
    Open sPath For Binary Access Read As #iFile
    ReDim bData(0 To LOF(iFile) - 1)
    Get #iFile, , bData
    Close #iFile
    PDF_ReadFileBytes = bData
End Function

Private Function PDF_BytesToLatin1(bData() As Byte) As String
    Dim i    As Long
    Dim nLen As Long
    Dim buf  As String
    nLen = UBound(bData) + 1
    buf = String$(nLen, 0)
    For i = 0 To nLen - 1
        Mid$(buf, i + 1, 1) = Chr(bData(i))
    Next i
    PDF_BytesToLatin1 = buf
End Function

Private Function PDF_ProcessAllStreams(bFile() As Byte, sRaw As String) As String
    ' old amounts are a theoretical risk.
    Dim lPos         As Long
    Dim lStart       As Long
    Dim lEnd         As Long
    Dim lLen         As Long
    Dim sHeader      As String
    Dim bStream()    As Byte
    Dim sText        As String
    Dim result       As String
    Dim lSearch      As Long
    Dim lHeaderStart As Long
    Dim lDictOpen    As Long
    Dim lScan        As Long
    Dim iByte        As Long
    Dim pre3         As String

    Dim sCMapData As String
    Dim lCMPos    As Long
    Dim lCMS      As Long
    Dim lCME      As Long
    Dim lCMBI     As Long
    Dim bCMRaw()  As Byte
    Dim bCMDec()  As Byte
    Dim lCMScan   As Long
    Dim sCMText   As String
    Dim sAllRuns  As String
    Dim sCMParsed As String
    ' Pass 0: extract CMaps from ObjStm streams (PDF 1.5+).
    sCMapData = PDF_ExtractObjStmCMaps(bFile, sRaw)
    lCMScan = 1
    ' Pass 1: collect ToUnicode CMaps from regular streams.
    Do
        lCMPos = InStr(lCMScan, sRaw, "stream", vbBinaryCompare)
        If lCMPos = 0 Then Exit Do
        If lCMPos >= 4 Then
            If Mid$(sRaw, lCMPos - 3, 3) = "end" Then lCMScan = lCMPos + 6: GoTo SkipCM
        End If
        lCMS = lCMPos + 6
        Do While lCMS <= Len(sRaw)
            If bFile(lCMS - 1) = 32 Or bFile(lCMS - 1) = 9 Then
                lCMS = lCMS + 1
            Else
                Exit Do
            End If
        Loop
        If lCMS <= Len(sRaw) Then
            If bFile(lCMS - 1) = 13 Then lCMS = lCMS + 1
        End If
        If lCMS <= Len(sRaw) Then
            If bFile(lCMS - 1) = 10 Then lCMS = lCMS + 1
        End If
        lCME = InStr(lCMS, sRaw, "endstream", vbBinaryCompare)
        If lCME > 0 And lCME - lCMS < 32768 Then   ' CMaps are small; large blobs are font binaries, not CMaps
            ReDim bCMRaw(0 To lCME - lCMS - 1)
            For lCMBI = 0 To lCME - lCMS - 1
                bCMRaw(lCMBI) = bFile(lCMS - 1 + lCMBI)
            Next lCMBI
            Dim nPredCM As Long, nColsCM As Long, nECCM As Long
            Dim sCMHdr  As String
            Dim lCMHS   As Long
            lCMHS  = IIf(lCMPos - 512 < 1, 1, lCMPos - 512)
            sCMHdr = Mid$(sRaw, lCMHS, lCMPos - lCMHS)
            If InStr(1, sCMHdr, "/ASCII85Decode", vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/A85",           vbBinaryCompare) > 0 Then
                bCMDec = PDF_DecodeASCII85(bCMRaw)
            Else
                bCMDec = bCMRaw
            End If
            If InStr(1, sCMHdr, "/ASCIIHexDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/AHx",            vbBinaryCompare) > 0 Then
                bCMDec = PDF_DecodeASCIIHex(bCMDec)
            End If
            If InStr(1, sCMHdr, "/LZWDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/LZW",       vbBinaryCompare) > 0 Then
                PDF_ParseDecodeParms sCMHdr, nPredCM, nColsCM, nECCM
                bCMDec = PDF_DecompressLZW(bCMDec, nECCM <> 0)
            End If
            If InStr(1, sCMHdr, "/RunLengthDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/RL",              vbBinaryCompare) > 0 Then
                bCMDec = PDF_DecodeRunLength(bCMDec)
            End If
            If InStr(1, sCMHdr, "/FlateDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/Fl ",         vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/Fl>",         vbBinaryCompare) > 0 Or _
               InStr(1, sCMHdr, "/Fl/",         vbBinaryCompare) > 0 Then
                bCMDec = PDF_DecompressDeflate(bCMDec)
                PDF_ParseDecodeParms sCMHdr, nPredCM, nColsCM, nECCM
                If nPredCM >= 2 Then bCMDec = PDF_ApplyPredictor(bCMDec, nPredCM, nColsCM)
            End If
            If UBound(bCMDec) > 10 Then
                sCMText = PDF_BytesToLatin1(bCMDec)
                If InStr(1, sCMText, "beginbfchar",  vbBinaryCompare) > 0 Or _
                   InStr(1, sCMText, "beginbfrange", vbBinaryCompare) > 0 Then
                    sCMParsed = PDF_ParseCMap(sCMText)
                    If Len(sCMParsed) > 0 Then
                        If Len(sCMapData) > 0 Then sCMapData = sCMapData & Chr(1)
                        sCMapData = sCMapData & sCMParsed
                    End If
                End If
            End If
        End If
        lCMScan = IIf(lCME > 0, lCME + 9, lCMPos + 6)
SkipCM:
    Loop

    sAllRuns = ""
    lSearch = 1

    Do
        lPos = InStr(lSearch, sRaw, "stream", vbBinaryCompare)
        If lPos = 0 Then Exit Do

        If lPos >= 4 Then pre3 = Mid$(sRaw, lPos - 3, 3) Else pre3 = ""
        If pre3 = "end" Then
            lSearch = lPos + 6
        Else
            lStart = lPos + 6
            Do While lStart <= Len(sRaw)
                If bFile(lStart - 1) = 32 Or bFile(lStart - 1) = 9 Then
                    lStart = lStart + 1
                Else
                    Exit Do
                End If
            Loop
            If lStart <= Len(sRaw) Then
                If bFile(lStart - 1) = 13 Then lStart = lStart + 1
            End If
            If lStart <= Len(sRaw) Then
                If bFile(lStart - 1) = 10 Then lStart = lStart + 1
            End If

            lEnd = InStr(lStart, sRaw, "endstream", vbBinaryCompare)
            If lEnd = 0 Then Exit Do

            lLen = lEnd - lStart

            If lLen > 0 And lLen < 50000000 Then

                lScan = lPos - 1
                lDictOpen = 0
                Do While lScan > 0
                    If Mid$(sRaw, lScan, 2) = "<<" Then
                        lDictOpen = lScan
                        Exit Do
                    End If
                    If Mid$(sRaw, lScan, 6) = "endobj" Then Exit Do
                    lScan = lScan - 1
                Loop
                If lDictOpen = 0 Then
                    lHeaderStart = lPos - 512
                    If lHeaderStart < 1 Then lHeaderStart = 1
                Else
                    lHeaderStart = lDictOpen
                End If
                sHeader = Mid$(sRaw, lHeaderStart, lPos - lHeaderStart)

                If PDF_IsContentStream(sHeader) Then
                    ReDim bStream(0 To lLen - 1)
                    For iByte = 0 To lLen - 1
                        bStream(iByte) = bFile(lStart - 1 + iByte)
                    Next iByte

                    Dim nPred1 As Long, nCols1 As Long, nEC1 As Long
                    Dim sHdrR1 As String
                    sHdrR1 = PDF_ResolveFilterRef(sHeader, sRaw)
                    If InStr(1, sHdrR1, "/ASCII85Decode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/A85",           vbBinaryCompare) > 0 Then
                        bStream = PDF_DecodeASCII85(bStream)
                    End If
                    If InStr(1, sHdrR1, "/ASCIIHexDecode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/AHx",            vbBinaryCompare) > 0 Then
                        bStream = PDF_DecodeASCIIHex(bStream)
                    End If
                    If InStr(1, sHdrR1, "/LZWDecode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/LZW",       vbBinaryCompare) > 0 Then
                        PDF_ParseDecodeParms sHeader, nPred1, nCols1, nEC1
                        bStream = PDF_DecompressLZW(bStream, nEC1 <> 0)
                        If nPred1 >= 2 Then bStream = PDF_ApplyPredictor(bStream, nPred1, nCols1)
                    End If
                    If InStr(1, sHdrR1, "/RunLengthDecode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/RL",              vbBinaryCompare) > 0 Then
                        bStream = PDF_DecodeRunLength(bStream)
                    End If
                    If InStr(1, sHdrR1, "/FlateDecode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/Fl ",         vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/Fl>",         vbBinaryCompare) > 0 Or _
                       InStr(1, sHdrR1, "/Fl/",         vbBinaryCompare) > 0 Then
                        bStream = PDF_DecompressDeflate(bStream)
                        PDF_ParseDecodeParms sHeader, nPred1, nCols1, nEC1
                        If nPred1 >= 2 Then bStream = PDF_ApplyPredictor(bStream, nPred1, nCols1)
                    End If

                    If UBound(bStream) > 0 Then
                        sText = PDF_ExtractTextOps(PDF_BytesToLatin1(bStream), sCMapData)
                        If Len(sText) > 0 Then
                            If Len(sAllRuns) > 0 Then sAllRuns = sAllRuns & Chr(2)
                            sAllRuns = sAllRuns & sText
                        End If
                    End If
                End If
            End If

            lSearch = lEnd + 9
        End If
    Loop

    PDF_ProcessAllStreams = PDF_CleanText(PDF_SortAndJoin(sAllRuns))
End Function

Private Function PDF_ExtractObjStmCMaps(bFile() As Byte, sRaw As String) As String
    ' ObjStm (PDF 1.5+): scan for ToUnicode CMaps inside compressed object bundles.
    Dim lSearch  As Long
    Dim lPos     As Long
    Dim lScan    As Long
    Dim lDO      As Long
    Dim lHS      As Long
    Dim lStart   As Long
    Dim lEnd     As Long
    Dim lLen     As Long
    Dim pre3     As String
    Dim sHeader  As String
    Dim lFirst   As Long
    Dim lFPos    As Long
    Dim lFEnd    As Long
    Dim bRaw()   As Byte
    Dim bDec()   As Byte
    Dim sDecomp  As String
    Dim sPart2   As String
    Dim sParsed  As String
    Dim result   As String
    Dim iByte    As Long

    lSearch = 1
    Do
        lPos = InStr(lSearch, sRaw, "stream", vbBinaryCompare)
        If lPos = 0 Then Exit Do

        If lPos >= 4 Then pre3 = Mid$(sRaw, lPos - 3, 3) Else pre3 = ""
        If pre3 = "end" Then
            lSearch = lPos + 6
            GoTo NextObjStm
        End If

        lScan = lPos - 1: lDO = 0
        Do While lScan > 0
            If Mid$(sRaw, lScan, 2) = "<<" Then lDO = lScan: Exit Do
            If Mid$(sRaw, lScan, 6) = "endobj" Then Exit Do
            lScan = lScan - 1
        Loop
        If lDO = 0 Then lHS = IIf(lPos - 512 < 1, 1, lPos - 512) Else lHS = lDO
        sHeader = Mid$(sRaw, lHS, lPos - lHS)

        If InStr(1, sHeader, "/Type /ObjStm",  vbBinaryCompare) = 0 And _
           InStr(1, sHeader, "/Type/ObjStm",   vbBinaryCompare) = 0 Then
            lSearch = lPos + 6
            GoTo NextObjStm
        End If

        ' Read /First  - byte offset in decompressed stream where Part 2 begins
        lFirst = 0
        lFPos = InStr(1, sHeader, "/First", vbBinaryCompare)
        If lFPos > 0 Then
            lFPos = lFPos + 6
            Do While lFPos <= Len(sHeader)
                If Mid$(sHeader, lFPos, 1) = " " Or Mid$(sHeader, lFPos, 1) = Chr(10) Or _
                   Mid$(sHeader, lFPos, 1) = Chr(13) Or Mid$(sHeader, lFPos, 1) = Chr(9) Then
                    lFPos = lFPos + 1
                Else
                    Exit Do
                End If
            Loop
            lFEnd = lFPos
            Do While lFEnd <= Len(sHeader)
                Dim ch As String: ch = Mid$(sHeader, lFEnd, 1)
                If ch >= "0" And ch <= "9" Then lFEnd = lFEnd + 1 Else Exit Do
            Loop
            If lFEnd > lFPos Then lFirst = CLng(Val(Mid$(sHeader, lFPos, lFEnd - lFPos)))
        End If

        lStart = lPos + 6
        Do While lStart <= Len(sRaw)
            If bFile(lStart - 1) = 32 Or bFile(lStart - 1) = 9 Then
                lStart = lStart + 1
            Else
                Exit Do
            End If
        Loop
        If lStart <= Len(sRaw) Then
            If bFile(lStart - 1) = 13 Then lStart = lStart + 1
        End If
        If lStart <= Len(sRaw) Then
            If bFile(lStart - 1) = 10 Then lStart = lStart + 1
        End If

        lEnd = InStr(lStart, sRaw, "endstream", vbBinaryCompare)
        If lEnd = 0 Then Exit Do
        lLen = lEnd - lStart

        If lLen > 0 And lLen < 50000000 Then
            ReDim bRaw(0 To lLen - 1)
            For iByte = 0 To lLen - 1
                bRaw(iByte) = bFile(lStart - 1 + iByte)
            Next iByte

            Dim nPredOS As Long, nColsOS As Long, nECOS As Long
            Dim sHdrR2 As String
            sHdrR2 = PDF_ResolveFilterRef(sHeader, sRaw)
            If InStr(1, sHdrR2, "/ASCII85Decode", vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/A85",           vbBinaryCompare) > 0 Then
                bDec = PDF_DecodeASCII85(bRaw)
            Else
                bDec = bRaw
            End If
            If InStr(1, sHdrR2, "/ASCIIHexDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/AHx",            vbBinaryCompare) > 0 Then
                bDec = PDF_DecodeASCIIHex(bDec)
            End If
            If InStr(1, sHdrR2, "/LZWDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/LZW",       vbBinaryCompare) > 0 Then
                PDF_ParseDecodeParms sHeader, nPredOS, nColsOS, nECOS
                bDec = PDF_DecompressLZW(bDec, nECOS <> 0)
            End If
            If InStr(1, sHdrR2, "/RunLengthDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/RL",              vbBinaryCompare) > 0 Then
                bDec = PDF_DecodeRunLength(bDec)
            End If
            If InStr(1, sHdrR2, "/FlateDecode", vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/Fl ",         vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/Fl>",         vbBinaryCompare) > 0 Or _
               InStr(1, sHdrR2, "/Fl/",         vbBinaryCompare) > 0 Then
                bDec = PDF_DecompressDeflate(bDec)
                PDF_ParseDecodeParms sHeader, nPredOS, nColsOS, nECOS
                If nPredOS >= 2 Then bDec = PDF_ApplyPredictor(bDec, nPredOS, nColsOS)
            End If

            If UBound(bDec) > lFirst Then
                sDecomp = PDF_BytesToLatin1(bDec)
                sPart2 = Mid$(sDecomp, lFirst + 1)
                If InStr(1, sPart2, "beginbfchar",  vbBinaryCompare) > 0 Or _
                   InStr(1, sPart2, "beginbfrange", vbBinaryCompare) > 0 Then
                    sParsed = PDF_ParseCMap(sPart2)
                    If Len(sParsed) > 0 Then
                        If Len(result) > 0 Then result = result & Chr(1)
                        result = result & sParsed
                    End If
                End If
            End If

            lSearch = lEnd + 9
        Else
            lSearch = lPos + 6
        End If
NextObjStm:
    Loop

    PDF_ExtractObjStmCMaps = result
End Function

Private Sub PDF_ParseDecodeParms(ByVal sHeader As String, _
                                  ByRef nPredictor  As Long, _
                                  ByRef nColumns    As Long, _
                                  Optional ByRef nEarlyChange As Long = 1)
    ' Parses /Predictor, /Columns (/FlateDecode) and /EarlyChange (/LZWDecode) from header.
    Dim lP   As Long
    Dim lC   As Long
    Dim lEnd As Long

    nPredictor  = 1
    nColumns    = 1
    nEarlyChange = 1

    lP = InStr(1, sHeader, "/Predictor", vbBinaryCompare)
    If lP > 0 Then
        lP = lP + 10
        Do While lP <= Len(sHeader)
            Dim cp As String: cp = Mid$(sHeader, lP, 1)
            If cp = " " Or cp = Chr(9) Or cp = Chr(10) Or cp = Chr(13) Then
                lP = lP + 1
            Else
                Exit Do
            End If
        Loop
        lEnd = lP
        Do While lEnd <= Len(sHeader)
            Dim dp As String: dp = Mid$(sHeader, lEnd, 1)
            If dp >= "0" And dp <= "9" Then lEnd = lEnd + 1 Else Exit Do
        Loop
        If lEnd > lP Then nPredictor = CLng(Val(Mid$(sHeader, lP, lEnd - lP)))
    End If

    lC = InStr(1, sHeader, "/Columns", vbBinaryCompare)
    If lC > 0 Then
        lC = lC + 8
        Do While lC <= Len(sHeader)
            Dim cc As String: cc = Mid$(sHeader, lC, 1)
            If cc = " " Or cc = Chr(9) Or cc = Chr(10) Or cc = Chr(13) Then
                lC = lC + 1
            Else
                Exit Do
            End If
        Loop
        lEnd = lC
        Do While lEnd <= Len(sHeader)
            Dim dc As String: dc = Mid$(sHeader, lEnd, 1)
            If dc >= "0" And dc <= "9" Then lEnd = lEnd + 1 Else Exit Do
        Loop
        If lEnd > lC Then nColumns = CLng(Val(Mid$(sHeader, lC, lEnd - lC)))
    End If

    Dim lE As Long
    lE = InStr(1, sHeader, "/EarlyChange", vbBinaryCompare)
    If lE > 0 Then
        lE = lE + 12
        Do While lE <= Len(sHeader)
            Dim ce As String: ce = Mid$(sHeader, lE, 1)
            If ce = " " Or ce = Chr(9) Or ce = Chr(10) Or ce = Chr(13) Then
                lE = lE + 1
            Else
                Exit Do
            End If
        Loop
        lEnd = lE
        Do While lEnd <= Len(sHeader)
            Dim de As String: de = Mid$(sHeader, lEnd, 1)
            If de >= "0" And de <= "9" Then lEnd = lEnd + 1 Else Exit Do
        Loop
        If lEnd > lE Then nEarlyChange = CLng(Val(Mid$(sHeader, lE, lEnd - lE)))
    End If
End Sub

Private Function PDF_ApplyPredictor(bData() As Byte, _
                                     ByVal nPredictor As Long, _
                                     ByVal nColumns   As Long) As Byte()
    ' Reverses FlateDecode PNG (10-15) and TIFF (2) predictor encoding.
    Dim nLen   As Long
    Dim bOut() As Byte
    Dim i      As Long
    Dim row    As Long
    Dim col    As Long
    Dim nRows  As Long
    Dim ftype  As Long
    Dim inOff  As Long
    Dim outOff As Long
    Dim a      As Long  ' left (reconstructed)
    Dim b      As Long  ' above (prev row reconstructed)
    Dim c      As Long  ' upper-left (prev row, left col)
    Dim raw    As Long

    nLen = UBound(bData) + 1
    If nLen = 0 Or nPredictor <= 1 Or nColumns <= 0 Then
        PDF_ApplyPredictor = bData
        Exit Function
    End If

    If nPredictor = 2 Then
        bOut = bData
        For i = 0 To nLen - 1
            If (i Mod nColumns) > 0 Then
                bOut(i) = CByte((CLng(bOut(i)) + CLng(bOut(i - 1))) And 255)
            End If
        Next i
        PDF_ApplyPredictor = bOut
        Exit Function
    End If

    If nPredictor >= 10 And nPredictor <= 15 Then
        Dim rowStride As Long
        rowStride = nColumns + 1
        If nLen Mod rowStride <> 0 Then
            ' Malformed stream  - return as-is and let text extraction fail naturally
            PDF_ApplyPredictor = bData
            Exit Function
        End If
        nRows = nLen \ rowStride
        ReDim bOut(0 To nRows * nColumns - 1)

        Dim prevRow() As Long
        ReDim prevRow(0 To nColumns - 1)  ' all zeros for first row

        For row = 0 To nRows - 1
            inOff  = row * rowStride
            outOff = row * nColumns
            ftype  = CLng(bData(inOff))  ' filter type byte

            For col = 0 To nColumns - 1
                raw = CLng(bData(inOff + 1 + col))
                a = IIf(col > 0, CLng(bOut(outOff + col - 1)), 0)
                b = prevRow(col)
                c = IIf(col > 0, prevRow(col - 1), 0)

                Select Case ftype
                    Case 0  ' None
                        bOut(outOff + col) = CByte(raw And 255)
                    Case 1  ' Sub: add left
                        bOut(outOff + col) = CByte((raw + a) And 255)
                    Case 2  ' Up: add above
                        bOut(outOff + col) = CByte((raw + b) And 255)
                    Case 3  ' Average: add floor((left + above) / 2)
                        bOut(outOff + col) = CByte((raw + (a + b) \ 2) And 255)
                    Case 4  ' Paeth
                        Dim p  As Long: p  = a + b - c
                        Dim pa As Long: pa = Abs(p - a)
                        Dim pb As Long: pb = Abs(p - b)
                        Dim pc As Long: pc = Abs(p - c)
                        Dim pr As Long
                        If pa <= pb And pa <= pc Then
                            pr = a
                        ElseIf pb <= pc Then
                            pr = b
                        Else
                            pr = c
                        End If
                        bOut(outOff + col) = CByte((raw + pr) And 255)
                    Case Else  ' Unknown filter  - treat as None
                        bOut(outOff + col) = CByte(raw And 255)
                End Select
            Next col

            For col = 0 To nColumns - 1
                prevRow(col) = CLng(bOut(outOff + col))
            Next col
        Next row

        PDF_ApplyPredictor = bOut
        Exit Function
    End If

    ' Unknown predictor - return unchanged.
    PDF_ApplyPredictor = bData
End Function

' Resolves an indirect /Filter reference (e.g. SAP NetWeaver: "/Filter 6 0 R").
' Returns sHeader unchanged if no indirect reference is found.
Private Function PDF_ResolveFilterRef(ByVal sHeader As String, _
                                      ByVal sRaw    As String) As String
    Dim lF   As Long
    Dim lN   As Long
    Dim lEnd As Long
    Dim sRef As String
    Dim sObj As String
    Dim lO   As Long
    lF = InStr(1, sHeader, "/Filter", vbBinaryCompare)
    If lF = 0 Then PDF_ResolveFilterRef = sHeader: Exit Function
    lN = lF + 7
    Do While lN <= Len(sHeader)
        Dim c As String: c = Mid$(sHeader, lN, 1)
        If c = " " Or c = Chr(9) Or c = Chr(10) Or c = Chr(13) Then
            lN = lN + 1
        Else
            Exit Do
        End If
    Loop
    If lN > Len(sHeader) Then PDF_ResolveFilterRef = sHeader: Exit Function
    c = Mid$(sHeader, lN, 1)
    If c = "/" Or c = "[" Then PDF_ResolveFilterRef = sHeader: Exit Function
    lEnd = lN
    Do While lEnd <= Len(sHeader)
        If Mid$(sHeader, lEnd, 1) >= "0" And Mid$(sHeader, lEnd, 1) <= "9" Then
            lEnd = lEnd + 1
        Else
            Exit Do
        End If
    Loop
    If lEnd = lN Then PDF_ResolveFilterRef = sHeader: Exit Function
    Dim sObjNum As String: sObjNum = Mid$(sHeader, lN, lEnd - lN)
    Dim sFind As String: sFind = Chr(10) & sObjNum & " "
    lO = InStr(1, sRaw, sFind, vbBinaryCompare)
    If lO = 0 Then sFind = Chr(13) & sObjNum & " ": lO = InStr(1, sRaw, sFind, vbBinaryCompare)
    If lO = 0 Then PDF_ResolveFilterRef = sHeader: Exit Function
    Dim lOBody As Long: lOBody = InStr(lO, sRaw, " obj", vbBinaryCompare)
    If lOBody = 0 Then PDF_ResolveFilterRef = sHeader: Exit Function
    lOBody = lOBody + 4
    sObj = Mid$(sRaw, lOBody, 128)
    PDF_ResolveFilterRef = sHeader & " " & sObj
End Function

Private Function PDF_IsContentStream(ByVal sHeader As String) As Boolean
    If InStr(1, sHeader, "/Subtype /Image",    vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype/Image",     vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/DCTDecode",         vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/CCITTFaxDecode",    vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/JBIG2Decode",       vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/JPXDecode",         vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/EmbeddedFile",      vbBinaryCompare) > 0 Then Exit Function
    ' /FontFile and /Length1 reliably identify font binary streams.
    If InStr(1, sHeader, "/FontFile",          vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Length1",           vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Length2",           vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype/Type1C",    vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype /Type1C",   vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype/CIDFontType0C",  vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype /CIDFontType0C", vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/ICCBased",          vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype/XML",       vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Subtype /XML",      vbBinaryCompare) > 0 Then Exit Function
    ' /Type/ObjStm: handled by PDF_ExtractObjStmCMaps. Re-processing causes
    ' garbage output from /Names arrays in XFA/SAP StructTree objects.
    If InStr(1, sHeader, "/Type/ObjStm",       vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Type /ObjStm",      vbBinaryCompare) > 0 Then Exit Function
    ' /Type/XRef: compressed cross-reference tables - binary data, never text.
    If InStr(1, sHeader, "/Type/XRef",         vbBinaryCompare) > 0 Then Exit Function
    If InStr(1, sHeader, "/Type /XRef",        vbBinaryCompare) > 0 Then Exit Function
    PDF_IsContentStream = True
End Function

Private Function PDF_DecompressDeflate(bIn() As Byte) As Byte()
    On Error GoTo Fail
    Dim lSkip        As Long
    Dim bEmpty(0)    As Byte
    lSkip = 0
    If UBound(bIn) >= 1 Then
        If (bIn(0) And &H0F) = 8 Then lSkip = 2  ' strip zlib 2-byte header (CM=8 = deflate)
    End If
    If UBound(bIn) - lSkip < 0 Then GoTo Fail
    PDF_DecompressDeflate = VBA_Inflate(bIn, lSkip)
    Exit Function
Fail:
    PDF_DecompressDeflate = bEmpty
End Function

Private Function PDF_DecodeASCII85(bIn() As Byte) As Byte()
    ' ASCII85 decoder (PDF spec sec. 7.4.3).
    ' Uses Double (not Long) - group values reach ~4.44e9, exceeding VBA Long range.
    On Error GoTo Fail
    Dim nIn      As Long
    Dim bOut()   As Byte
    Dim nOutSize As Long
    Dim nOutPos  As Long
    Dim i        As Long
    Dim b        As Long
    Dim nGroup   As Long
    Dim grp(4)   As Long   ' up to 5 digits per group
    Dim v        As Double
    Dim bFail(0) As Byte

    nIn = UBound(bIn) + 1
    If nIn = 0 Then PDF_DecodeASCII85 = bFail: Exit Function

    nOutSize = (nIn \ 5 + 1) * 4 + 16
    ReDim bOut(0 To nOutSize - 1)
    nOutPos = 0
    nGroup  = 0
    i = 0

    Do While i < nIn
        b = bIn(i)

        If b = 126 Then   ' '~'
            If i + 1 < nIn Then
                If bIn(i + 1) = 62 Then   ' '>'
                    Exit Do
                End If
            End If
            i = i + 1
            GoTo NextByte
        End If

        If b = 32 Or b = 9 Or b = 10 Or b = 13 Then
            i = i + 1
            GoTo NextByte
        End If

        If b = 122 And nGroup = 0 Then  ' 'z'
            If nOutPos + 4 > nOutSize Then
                nOutSize = nOutSize + 1024
                ReDim Preserve bOut(0 To nOutSize - 1)
            End If
            bOut(nOutPos) = 0: bOut(nOutPos + 1) = 0
            bOut(nOutPos + 2) = 0: bOut(nOutPos + 3) = 0
            nOutPos = nOutPos + 4
            i = i + 1
            GoTo NextByte
        End If

        If b >= 33 And b <= 117 Then
            grp(nGroup) = b - 33
            nGroup = nGroup + 1
            If nGroup = 5 Then
                v = CDbl(grp(0)) * 52200625# + _
                    CDbl(grp(1)) *   614125# + _
                    CDbl(grp(2)) *     7225# + _
                    CDbl(grp(3)) *       85# + _
                    CDbl(grp(4))
                If nOutPos + 4 > nOutSize Then
                    nOutSize = nOutSize + 4096
                    ReDim Preserve bOut(0 To nOutSize - 1)
                End If
                bOut(nOutPos)     = CByte(Int(v / 16777216#) And 255)
                bOut(nOutPos + 1) = CByte(Int(v / 65536#)    And 255)
                bOut(nOutPos + 2) = CByte(Int(v / 256#)      And 255)
                bOut(nOutPos + 3) = CByte(v                  And 255)
                nOutPos = nOutPos + 4
                nGroup = 0
            End If
        End If

        i = i + 1
NextByte:
    Loop

    If nGroup > 1 Then
        Dim k As Long
        For k = nGroup To 4
            grp(k) = 84
        Next k
        v = CDbl(grp(0)) * 52200625# + _
            CDbl(grp(1)) *   614125# + _
            CDbl(grp(2)) *     7225# + _
            CDbl(grp(3)) *       85# + _
            CDbl(grp(4))
        Dim nBytes As Long: nBytes = nGroup - 1
        If nOutPos + nBytes > nOutSize Then
            nOutSize = nOutPos + nBytes + 16
            ReDim Preserve bOut(0 To nOutSize - 1)
        End If
        If nBytes >= 1 Then bOut(nOutPos)     = CByte(Int(v / 16777216#) And 255)
        If nBytes >= 2 Then bOut(nOutPos + 1) = CByte(Int(v / 65536#)    And 255)
        If nBytes >= 3 Then bOut(nOutPos + 2) = CByte(Int(v / 256#)      And 255)
        nOutPos = nOutPos + nBytes
    End If

    If nOutPos = 0 Then
        PDF_DecodeASCII85 = bFail
    Else
        ReDim Preserve bOut(0 To nOutPos - 1)
        PDF_DecodeASCII85 = bOut
    End If
    Exit Function
Fail:
    PDF_DecodeASCII85 = bFail
End Function

Private Function PDF_DecompressLZW(bIn() As Byte, ByVal bEarlyChange As Boolean) As Byte()
    ' LZW decoder (PDF sec. 7.4.4). MSB-first, 9-12 bit codes.
    ' EarlyChange=1 (PDF default): code added before width bump.
    ' EarlyChange=0: standard TIFF late-change (/DecodeParms /EarlyChange 0).
    ' Table: 0-255 = literals, 256 = CLEAR, 257 = EOI, 258+ = dynamic entries.
    ' KwKwK edge case: code == next available slot -> entry = prevStr + prevStr[0].
    On Error GoTo Fail
    Const LZW_CLEAR As Long = 256
    Const LZW_EOI   As Long = 257
    Const LZW_FIRST As Long = 258
    Const LZW_MAXCODES As Long = 4096  ' 12-bit cap

    Dim lBitBuf  As Long
    Dim lBitCnt  As Long
    Dim lInPos   As Long
    Dim bFail(0) As Byte

    Dim tPrefix(LZW_MAXCODES - 1) As Long  ' -1 = literal (no prefix)
    Dim tSuffix(LZW_MAXCODES - 1) As Long
    Dim tLen(LZW_MAXCODES - 1)    As Long  ' cached string length

    Dim bOut()   As Byte
    Dim lOutSize As Long
    Dim lOutPos  As Long
    lOutSize = 65536
    ReDim bOut(0 To lOutSize - 1)
    lOutPos = 0

    Dim codeSize As Long
    Dim nextCode As Long
    Dim code     As Long
    Dim prevCode As Long
    Dim i        As Long
    Dim j        As Long
    Dim stackLen As Long
    Dim stackBuf(LZW_MAXCODES) As Long  ' temp stack for string reversal

    For i = 0 To 255
        tPrefix(i) = -1
        tSuffix(i) = i
        tLen(i)    = 1
    Next i
    codeSize = 9
    nextCode = LZW_FIRST
    prevCode = -1

    lInPos  = 0
    lBitBuf = 0
    lBitCnt = 0

    Do
        Dim lCode As Long: lCode = 0
        Dim bitsLeft As Long: bitsLeft = codeSize
        Do While bitsLeft > 0
            If lBitCnt = 0 Then
                If lInPos > UBound(bIn) Then GoTo EmitResult
                lBitBuf = bIn(lInPos): lInPos = lInPos + 1: lBitCnt = 8
            End If
            Dim take As Long: take = IIf(bitsLeft < lBitCnt, bitsLeft, lBitCnt)
            lCode = (lCode << take) Or (lBitBuf >> (lBitCnt - take))
            lCode = lCode And ((1 << codeSize) - 1)   ' keep only codeSize bits
            lBitBuf = lBitBuf And ((1 << (lBitCnt - take)) - 1)
            lBitCnt = lBitCnt - take
            bitsLeft = bitsLeft - take
        Loop
        code = lCode

        If code = LZW_EOI Then Exit Do
        If code = LZW_CLEAR Then
            codeSize = 9
            nextCode = LZW_FIRST
            prevCode = -1
        Else
            Dim firstByte As Long
            If code < nextCode And tLen(code) > 0 Then
                stackLen = 0
                Dim chainCode As Long: chainCode = code
                Do While chainCode >= 0
                    stackBuf(stackLen) = tSuffix(chainCode)
                    stackLen = stackLen + 1
                    chainCode = tPrefix(chainCode)
                Loop
                firstByte = stackBuf(stackLen - 1)
            ElseIf code = nextCode And prevCode >= 0 Then
                ' KwKwK: new code not yet in table; string = prev + prev[0]
                stackLen = 0
                chainCode = prevCode
                Do While chainCode >= 0
                    stackBuf(stackLen) = tSuffix(chainCode)
                    stackLen = stackLen + 1
                    chainCode = tPrefix(chainCode)
                Loop
                firstByte = stackBuf(stackLen - 1)
                stackBuf(stackLen) = firstByte
                stackLen = stackLen + 1
            Else
                GoTo Fail  ' corrupt stream
            End If

            If lOutPos + stackLen > lOutSize Then
                lOutSize = lOutPos + stackLen + 65536
                ReDim Preserve bOut(0 To lOutSize - 1)
            End If
            For j = stackLen - 1 To 0 Step -1
                bOut(lOutPos) = CByte(stackBuf(j) And 255)
                lOutPos = lOutPos + 1
            Next j

            If prevCode >= 0 And nextCode < LZW_MAXCODES Then
                tPrefix(nextCode) = prevCode
                tSuffix(nextCode) = firstByte
                tLen(nextCode)    = tLen(prevCode) + 1
                nextCode = nextCode + 1

                If bEarlyChange Then
                    ' PDF default: bump when table fills INCLUDING the slot just added
                    If nextCode = (1 << codeSize) And codeSize < 12 Then
                        codeSize = codeSize + 1
                    End If
                Else
                    ' Late change: bump when table fills BEFORE next add
                    If nextCode - 1 = (1 << codeSize) And codeSize < 12 Then
                        codeSize = codeSize + 1
                    End If
                End If
            End If

            prevCode = code
        End If
    Loop

EmitResult:
    If lOutPos = 0 Then
        PDF_DecompressLZW = bFail
    Else
        ReDim Preserve bOut(0 To lOutPos - 1)
        PDF_DecompressLZW = bOut
    End If
    Exit Function
Fail:
    PDF_DecompressLZW = bFail
End Function

Private Function PDF_DecodeASCIIHex(bIn() As Byte) As Byte()
    ' ASCIIHexDecode decoder (PDF sec. 7.4.2). Hex pairs -> bytes, '>' terminates.
    On Error GoTo Fail
    Dim bFail(0) As Byte
    Dim nIn      As Long: nIn = UBound(bIn) + 1
    If nIn = 0 Then PDF_DecodeASCIIHex = bFail: Exit Function

    Dim bOut()   As Byte
    ReDim bOut(0 To nIn \ 2 + 1)
    Dim nOutPos  As Long: nOutPos = 0
    Dim i        As Long
    Dim c        As Long
    Dim hi       As Long: hi = -1
    Dim nibble   As Long

    For i = 0 To nIn - 1
        c = bIn(i)
        If c = 62 Then Exit For   ' '>' end marker
        If c >= 48 And c <= 57 Then       ' '0'-'9'
            nibble = c - 48
        ElseIf c >= 65 And c <= 70 Then   ' 'A'-'F'
            nibble = c - 55
        ElseIf c >= 97 And c <= 102 Then  ' 'a'-'f'
            nibble = c - 87
        ElseIf c = 32 Or c = 9 Or c = 10 Or c = 13 Then
            GoTo NextHexByte  ' skip whitespace
        Else
            GoTo NextHexByte  ' ignore invalid chars
        End If

        If hi = -1 Then
            hi = nibble
        Else
            bOut(nOutPos) = CByte((hi * 16 + nibble) And 255)
            nOutPos = nOutPos + 1
            hi = -1
        End If
NextHexByte:
    Next i

    If hi >= 0 Then
        bOut(nOutPos) = CByte((hi * 16) And 255)
        nOutPos = nOutPos + 1
    End If

    If nOutPos = 0 Then
        PDF_DecodeASCIIHex = bFail
    Else
        ReDim Preserve bOut(0 To nOutPos - 1)
        PDF_DecodeASCIIHex = bOut
    End If
    Exit Function
Fail:
    PDF_DecodeASCIIHex = bFail
End Function

Private Function PDF_DecodeRunLength(bIn() As Byte) As Byte()
    ' RunLengthDecode decoder (PDF sec. 7.4.5, PackBits).
    On Error GoTo Fail
    Dim bFail(0) As Byte
    Dim nIn      As Long: nIn = UBound(bIn) + 1
    If nIn = 0 Then PDF_DecodeRunLength = bFail: Exit Function

    Dim bOut()   As Byte
    Dim lOutSize As Long: lOutSize = nIn * 2 + 16
    ReDim bOut(0 To lOutSize - 1)
    Dim lOutPos  As Long: lOutPos = 0
    Dim i        As Long: i = 0
    Dim length   As Long
    Dim n        As Long
    Dim j        As Long

    Do While i < nIn
        length = bIn(i): i = i + 1
        If length = 128 Then Exit Do  ' EOD

        If length < 128 Then
            n = length + 1
            If lOutPos + n > lOutSize Then
                lOutSize = lOutPos + n + 16384
                ReDim Preserve bOut(0 To lOutSize - 1)
            End If
            For j = 0 To n - 1
                If i > UBound(bIn) Then GoTo EmitRL
                bOut(lOutPos) = bIn(i): lOutPos = lOutPos + 1: i = i + 1
            Next j
        Else
            n = 257 - length
            If i > UBound(bIn) Then GoTo EmitRL
            Dim repByte As Byte: repByte = bIn(i): i = i + 1
            If lOutPos + n > lOutSize Then
                lOutSize = lOutPos + n + 16384
                ReDim Preserve bOut(0 To lOutSize - 1)
            End If
            For j = 0 To n - 1
                bOut(lOutPos) = repByte: lOutPos = lOutPos + 1
            Next j
        End If
    Loop

EmitRL:
    If lOutPos = 0 Then
        PDF_DecodeRunLength = bFail
    Else
        ReDim Preserve bOut(0 To lOutPos - 1)
        PDF_DecodeRunLength = bOut
    End If
    Exit Function
Fail:
    PDF_DecodeRunLength = bFail
End Function

Private Function VBA_Inflate(bIn() As Byte, ByVal lSkip As Long) As Byte()
    On Error GoTo Fail

    Dim lBitBuf   As Long  ' bit buffer (up to 32 bits, LSB-first per DEFLATE)
    Dim lBitCnt   As Long  ' valid bits in lBitBuf
    Dim lInPos    As Long  ' current read position in bIn

    Dim bOut()    As Byte
    Dim lOutSize  As Long
    Dim lOutPos   As Long
    lOutSize = 65536
    ReDim bOut(0 To lOutSize - 1)

    ' RFC 1951 length/distance tables
    Dim LEN_EXTRA(30)  As Long
    Dim LEN_BASE(30)   As Long
    Dim DIST_EXTRA(29) As Long
    Dim DIST_BASE(29)  As Long

    LEN_EXTRA(0)=0:  LEN_BASE(0)=3
    LEN_EXTRA(1)=0:  LEN_BASE(1)=4
    LEN_EXTRA(2)=0:  LEN_BASE(2)=5
    LEN_EXTRA(3)=0:  LEN_BASE(3)=6
    LEN_EXTRA(4)=0:  LEN_BASE(4)=7
    LEN_EXTRA(5)=0:  LEN_BASE(5)=8
    LEN_EXTRA(6)=0:  LEN_BASE(6)=9
    LEN_EXTRA(7)=0:  LEN_BASE(7)=10
    LEN_EXTRA(8)=1:  LEN_BASE(8)=11
    LEN_EXTRA(9)=1:  LEN_BASE(9)=13
    LEN_EXTRA(10)=1: LEN_BASE(10)=15
    LEN_EXTRA(11)=1: LEN_BASE(11)=17
    LEN_EXTRA(12)=2: LEN_BASE(12)=19
    LEN_EXTRA(13)=2: LEN_BASE(13)=23
    LEN_EXTRA(14)=2: LEN_BASE(14)=27
    LEN_EXTRA(15)=2: LEN_BASE(15)=31
    LEN_EXTRA(16)=3: LEN_BASE(16)=35
    LEN_EXTRA(17)=3: LEN_BASE(17)=43
    LEN_EXTRA(18)=3: LEN_BASE(18)=51
    LEN_EXTRA(19)=3: LEN_BASE(19)=59
    LEN_EXTRA(20)=4: LEN_BASE(20)=67
    LEN_EXTRA(21)=4: LEN_BASE(21)=83
    LEN_EXTRA(22)=4: LEN_BASE(22)=99
    LEN_EXTRA(23)=4: LEN_BASE(23)=115
    LEN_EXTRA(24)=5: LEN_BASE(24)=131
    LEN_EXTRA(25)=5: LEN_BASE(25)=163
    LEN_EXTRA(26)=5: LEN_BASE(26)=195
    LEN_EXTRA(27)=5: LEN_BASE(27)=227
    LEN_EXTRA(28)=0: LEN_BASE(28)=258
    LEN_EXTRA(29)=0: LEN_BASE(29)=0
    LEN_EXTRA(30)=0: LEN_BASE(30)=0

    DIST_EXTRA(0)=0:  DIST_BASE(0)=1
    DIST_EXTRA(1)=0:  DIST_BASE(1)=2
    DIST_EXTRA(2)=0:  DIST_BASE(2)=3
    DIST_EXTRA(3)=0:  DIST_BASE(3)=4
    DIST_EXTRA(4)=1:  DIST_BASE(4)=5
    DIST_EXTRA(5)=1:  DIST_BASE(5)=7
    DIST_EXTRA(6)=2:  DIST_BASE(6)=9
    DIST_EXTRA(7)=2:  DIST_BASE(7)=13
    DIST_EXTRA(8)=3:  DIST_BASE(8)=17
    DIST_EXTRA(9)=3:  DIST_BASE(9)=25
    DIST_EXTRA(10)=4: DIST_BASE(10)=33
    DIST_EXTRA(11)=4: DIST_BASE(11)=49
    DIST_EXTRA(12)=5: DIST_BASE(12)=65
    DIST_EXTRA(13)=5: DIST_BASE(13)=97
    DIST_EXTRA(14)=6: DIST_BASE(14)=129
    DIST_EXTRA(15)=6: DIST_BASE(15)=193
    DIST_EXTRA(16)=7: DIST_BASE(16)=257
    DIST_EXTRA(17)=7: DIST_BASE(17)=385
    DIST_EXTRA(18)=8: DIST_BASE(18)=513
    DIST_EXTRA(19)=8: DIST_BASE(19)=769
    DIST_EXTRA(20)=9: DIST_BASE(20)=1025
    DIST_EXTRA(21)=9: DIST_BASE(21)=1537
    DIST_EXTRA(22)=10: DIST_BASE(22)=2049
    DIST_EXTRA(23)=10: DIST_BASE(23)=3073
    DIST_EXTRA(24)=11: DIST_BASE(24)=4097
    DIST_EXTRA(25)=11: DIST_BASE(25)=6145
    DIST_EXTRA(26)=12: DIST_BASE(26)=8193
    DIST_EXTRA(27)=12: DIST_BASE(27)=12289
    DIST_EXTRA(28)=13: DIST_BASE(28)=16385
    DIST_EXTRA(29)=13: DIST_BASE(29)=24577

    ' Code-length alphabet reorder (RFC 1951 sec. 3.2.7)
    Dim CL_ORDER(18) As Long
    CL_ORDER(0)=16: CL_ORDER(1)=17: CL_ORDER(2)=18: CL_ORDER(3)=0
    CL_ORDER(4)=8:  CL_ORDER(5)=7:  CL_ORDER(6)=9:  CL_ORDER(7)=6
    CL_ORDER(8)=10: CL_ORDER(9)=5:  CL_ORDER(10)=11: CL_ORDER(11)=4
    CL_ORDER(12)=12: CL_ORDER(13)=3: CL_ORDER(14)=13: CL_ORDER(15)=2
    CL_ORDER(16)=14: CL_ORDER(17)=1: CL_ORDER(18)=15

    Dim HT_CODE(287) As Long
    Dim HT_BLEN(287) As Long
    Dim HT_SYM(287)  As Long
    Dim HT_MAX       As Long
    Dim HT_SIZE      As Long

    Dim lengths(287) As Long
    Dim bl_count(15) As Long
    Dim next_code(15) As Long

    lInPos = lSkip
    lBitBuf = 0
    lBitCnt = 0

    Dim bFinal   As Long
    Dim bType    As Long
    Dim i        As Long
    Dim j        As Long
    Dim sym      As Long
    Dim lLen     As Long
    Dim lDist    As Long
    Dim lStart   As Long
    Dim hlit     As Long
    Dim hdist    As Long
    Dim hclen    As Long
    Dim nCodes   As Long
    Dim rep      As Long
    Dim prev     As Long
    Dim blkLen   As Long
    Dim distSym  As Long

    ' Fixed Huffman code lengths (RFC 1951 sec. 3.2.6)
    Dim fixLL(287) As Long
    For i = 0 To 143:   fixLL(i) = 8: Next i
    For i = 144 To 255: fixLL(i) = 9: Next i
    For i = 256 To 279: fixLL(i) = 7: Next i
    For i = 280 To 287: fixLL(i) = 8: Next i

    Dim fixDist(31) As Long
    For i = 0 To 31: fixDist(i) = 5: Next i

    Dim LL_CODE(287) As Long
    Dim LL_BLEN(287) As Long
    Dim LL_SYM(287)  As Long
    Dim LL_MAX       As Long
    Dim LL_SIZE      As Long
    Dim DS_CODE(31)  As Long
    Dim DS_BLEN(31)  As Long
    Dim DS_SYM(31)   As Long
    Dim DS_MAX       As Long
    Dim DS_SIZE      As Long

    Do  ' block loop
        bFinal = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 1)
        bType  = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 2)

        If bType = 0 Then  ' stored block
            lBitBuf = 0: lBitCnt = 0  ' byte-align
            blkLen = bIn(lInPos) + bIn(lInPos + 1) * 256
            lInPos = lInPos + 4  ' skip LEN + NLEN
            If lOutPos + blkLen > lOutSize Then
                lOutSize = lOutPos + blkLen + 65536
                ReDim Preserve bOut(0 To lOutSize - 1)
            End If
            For i = 0 To blkLen - 1
                bOut(lOutPos) = bIn(lInPos)
                lOutPos = lOutPos + 1
                lInPos = lInPos + 1
            Next i

        ElseIf bType = 1 Then  ' fixed Huffman
            INF_BuildTable fixLL, 288, LL_CODE, LL_BLEN, LL_SYM, LL_MAX, LL_SIZE
            INF_BuildTable fixDist, 32, DS_CODE, DS_BLEN, DS_SYM, DS_MAX, DS_SIZE

            Do
                sym = INF_DecodeHuff(bIn, lInPos, lBitBuf, lBitCnt, LL_CODE, LL_BLEN, LL_SYM, LL_MAX, LL_SIZE)
                If sym < 256 Then
                    If lOutPos >= lOutSize Then
                        lOutSize = lOutSize + 65536
                        ReDim Preserve bOut(0 To lOutSize - 1)
                    End If
                    bOut(lOutPos) = CByte(sym): lOutPos = lOutPos + 1
                ElseIf sym = 256 Then
                    Exit Do
                Else
                    i = sym - 257
                    lLen = LEN_BASE(i) + INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, LEN_EXTRA(i))
                    distSym = INF_DecodeHuff(bIn, lInPos, lBitBuf, lBitCnt, DS_CODE, DS_BLEN, DS_SYM, DS_MAX, DS_SIZE)
                    lDist = DIST_BASE(distSym) + INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, DIST_EXTRA(distSym))
                    lStart = lOutPos - lDist
                    If lOutPos + lLen > lOutSize Then
                        lOutSize = lOutPos + lLen + 65536
                        ReDim Preserve bOut(0 To lOutSize - 1)
                    End If
                    For j = 0 To lLen - 1
                        bOut(lOutPos) = bOut(lStart + (j Mod lDist))
                        lOutPos = lOutPos + 1
                    Next j
                End If
            Loop

        ElseIf bType = 2 Then  ' dynamic Huffman
            hlit  = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 5) + 257
            hdist = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 5) + 1
            hclen = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 4) + 4

            Dim cl_lens(18) As Long
            For i = 0 To 18: cl_lens(i) = 0: Next i
            For i = 0 To hclen - 1
                cl_lens(CL_ORDER(i)) = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 3)
            Next i

            Dim CL_CODE(18) As Long
            Dim CL_BLEN(18) As Long
            Dim CL_SYM(18)  As Long
            Dim CL_MAX      As Long
            Dim CL_SIZE     As Long
            INF_BuildTable cl_lens, 19, CL_CODE, CL_BLEN, CL_SYM, CL_MAX, CL_SIZE

            nCodes = hlit + hdist
            Dim all_lens(575) As Long  ' max 288+32 but use 576 for safety
            i = 0
            Do While i < nCodes
                sym = INF_DecodeHuff(bIn, lInPos, lBitBuf, lBitCnt, CL_CODE, CL_BLEN, CL_SYM, CL_MAX, CL_SIZE)
                If sym < 16 Then
                    all_lens(i) = sym: i = i + 1
                ElseIf sym = 16 Then
                    rep = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 2) + 3
                    prev = all_lens(i - 1)
                    For j = 0 To rep - 1: all_lens(i) = prev: i = i + 1: Next j
                ElseIf sym = 17 Then
                    rep = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 3) + 3
                    For j = 0 To rep - 1: all_lens(i) = 0: i = i + 1: Next j
                ElseIf sym = 18 Then
                    rep = INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, 7) + 11
                    For j = 0 To rep - 1: all_lens(i) = 0: i = i + 1: Next j
                End If
            Loop

            Dim ll_lens(287) As Long
            Dim dt_lens(31)  As Long
            For i = 0 To hlit - 1:  ll_lens(i) = all_lens(i):        Next i
            For i = hlit To 287:    ll_lens(i) = 0:                   Next i
            For i = 0 To hdist - 1: dt_lens(i) = all_lens(hlit + i):  Next i
            For i = hdist To 31:    dt_lens(i) = 0:                    Next i

            INF_BuildTable ll_lens, hlit,  LL_CODE, LL_BLEN, LL_SYM, LL_MAX, LL_SIZE
            INF_BuildTable dt_lens, hdist, DS_CODE, DS_BLEN, DS_SYM, DS_MAX, DS_SIZE

            Do
                sym = INF_DecodeHuff(bIn, lInPos, lBitBuf, lBitCnt, LL_CODE, LL_BLEN, LL_SYM, LL_MAX, LL_SIZE)
                If sym < 256 Then
                    If lOutPos >= lOutSize Then
                        lOutSize = lOutSize + 65536
                        ReDim Preserve bOut(0 To lOutSize - 1)
                    End If
                    bOut(lOutPos) = CByte(sym): lOutPos = lOutPos + 1
                ElseIf sym = 256 Then
                    Exit Do
                Else
                    i = sym - 257
                    lLen = LEN_BASE(i) + INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, LEN_EXTRA(i))
                    distSym = INF_DecodeHuff(bIn, lInPos, lBitBuf, lBitCnt, DS_CODE, DS_BLEN, DS_SYM, DS_MAX, DS_SIZE)
                    lDist = DIST_BASE(distSym) + INF_ReadBits(bIn, lInPos, lBitBuf, lBitCnt, DIST_EXTRA(distSym))
                    lStart = lOutPos - lDist
                    If lOutPos + lLen > lOutSize Then
                        lOutSize = lOutPos + lLen + 65536
                        ReDim Preserve bOut(0 To lOutSize - 1)
                    End If
                    For j = 0 To lLen - 1
                        bOut(lOutPos) = bOut(lStart + (j Mod lDist))
                        lOutPos = lOutPos + 1
                    Next j
                End If
            Loop
        End If
    Loop While bFinal = 0

    If lOutPos = 0 Then
        Dim bZero(0) As Byte
        VBA_Inflate = bZero
    Else
        ReDim Preserve bOut(0 To lOutPos - 1)
        VBA_Inflate = bOut
    End If
    Exit Function
Fail:
    Dim bFail(0) As Byte
    VBA_Inflate = bFail
End Function

Private Function INF_ReadBits(bIn() As Byte, lInPos As Long, _
                               lBitBuf As Long, lBitCnt As Long, _
                               ByVal n As Long) As Long
    If n = 0 Then INF_ReadBits = 0: Exit Function
    Do While lBitCnt < n
        lBitBuf = lBitBuf Or (CLng(bIn(lInPos)) * (2 ^ lBitCnt))
        lBitCnt = lBitCnt + 8
        lInPos = lInPos + 1
    Loop
    INF_ReadBits = lBitBuf And ((2 ^ n) - 1)
    lBitBuf = lBitBuf \ (2 ^ n)
    lBitCnt = lBitCnt - n
End Function

Private Sub INF_BuildTable(lengths() As Long, ByVal nSyms As Long, _
                            CODE() As Long, BLEN() As Long, SYM() As Long, _
                            maxBits As Long, tSize As Long)
    Dim bl_count(15) As Long
    Dim next_code(15) As Long
    Dim i            As Long
    Dim bits         As Long
    Dim lCode        As Long

    maxBits = 0
    For i = 0 To nSyms - 1
        If lengths(i) > maxBits Then maxBits = lengths(i)
    Next i
    If maxBits = 0 Then tSize = 0: Exit Sub

    For i = 0 To nSyms - 1
        If lengths(i) > 0 Then bl_count(lengths(i)) = bl_count(lengths(i)) + 1
    Next i

    lCode = 0
    bl_count(0) = 0
    For bits = 1 To maxBits
        lCode = (lCode + bl_count(bits - 1)) * 2
        next_code(bits) = lCode
    Next bits

    tSize = 0
    For i = 0 To nSyms - 1
        If lengths(i) > 0 Then
            CODE(tSize) = next_code(lengths(i))
            BLEN(tSize) = lengths(i)
            SYM(tSize)  = i
            next_code(lengths(i)) = next_code(lengths(i)) + 1
            tSize = tSize + 1
        End If
    Next i
End Sub

Private Function INF_DecodeHuff(bIn() As Byte, lInPos As Long, _
                                 lBitBuf As Long, lBitCnt As Long, _
                                 CODE() As Long, BLEN() As Long, SYM() As Long, _
                                 ByVal maxBits As Long, ByVal tSize As Long) As Long
    Dim lCode   As Long
    Dim bits    As Long
    Dim i       As Long

    lCode = 0
    For bits = 1 To maxBits
        If lBitCnt = 0 Then
            lBitBuf = CLng(bIn(lInPos))
            lBitCnt = 8
            lInPos = lInPos + 1
        End If
        lCode = lCode * 2 + (lBitBuf And 1)
        lBitBuf = lBitBuf \ 2
        lBitCnt = lBitCnt - 1

        For i = 0 To tSize - 1
            If BLEN(i) = bits And CODE(i) = lCode Then
                INF_DecodeHuff = SYM(i)
                Exit Function
            End If
        Next i
    Next bits
    INF_DecodeHuff = -1  ' error
End Function

Private Function PDF_ExtractTextOps(ByVal sStream As String, ByVal sCMap As String) As String
    Dim i        As Long
    Dim lLen     As Long
    Dim c        As String
    Dim result   As String
    Dim tokens() As String
    Dim tCount   As Long
    Dim cABT     As String
    Dim cAET     As String
    Dim sLit     As String
    Dim depth    As Long
    Dim cl       As String
    Dim cE       As String
    Dim sO       As String
    Dim o2       As String
    Dim o3       As String
    Dim sHx      As String
    Dim hc       As String
    Dim op       As String
    Dim k        As Long
    Dim curX     As Double
    Dim curY     As Double
    Dim curLead  As Double
    Dim nTok     As Long
    Dim sTok     As String
    Dim nc       As String
    Dim sNum     As String
    Dim trimP    As Long
    Dim sRunOut  As String
    Dim emitY    As Long
    Dim sTokArr  As Variant
    Dim nArr     As Long
    Dim tokIsHex() As Boolean  ' True when token came from <hex> string (needs CMap); False for (literal)

    curLead = 12
    lLen = Len(sStream)
    ReDim tokens(0 To 1023)
    ReDim tokIsHex(0 To 1023)
    tCount = 0
    i = 1

    Do While i <= lLen
        c = Mid$(sStream, i, 1)

        Select Case c

        Case "B"
            If i + 1 <= lLen Then
                If Mid$(sStream, i, 2) = "BT" Then
                    If i + 2 <= lLen Then cABT = Mid$(sStream, i + 2, 1) Else cABT = " "
                    If cABT = " " Or cABT = Chr(9) Or cABT = Chr(10) Or cABT = Chr(13) Then
                        tCount = 0: curX = 0: curY = 0: curLead = 12
                        nTok = 0: sTok = ""
                        i = i + 2: GoTo NextChar
                    End If
                End If
            End If

        Case "E"
            If i + 1 <= lLen Then
                If Mid$(sStream, i, 2) = "ET" Then
                    If i + 2 <= lLen Then cAET = Mid$(sStream, i + 2, 1) Else cAET = " "
                    If cAET = " " Or cAET = Chr(9) Or cAET = Chr(10) Or cAET = Chr(13) Then
                        tCount = 0: i = i + 2: GoTo NextChar
                    End If
                End If
            End If

        Case "("
            depth = 1: i = i + 1: sLit = ""
            Do While i <= lLen And depth > 0
                cl = Mid$(sStream, i, 1)
                If cl = "\" And i + 1 <= lLen Then
                    cE = Mid$(sStream, i + 1, 1)
                    Select Case cE
                        Case "n":  sLit = sLit & Chr(10): i = i + 2
                        Case "r":  sLit = sLit & Chr(13): i = i + 2
                        Case "t":  sLit = sLit & Chr(9):  i = i + 2
                        Case "(":  sLit = sLit & "(":     i = i + 2
                        Case ")":  sLit = sLit & ")":     i = i + 2
                        Case "\": sLit = sLit & "\":    i = i + 2
                        Case Else
                            If cE >= "0" And cE <= "7" Then
                                sO = cE
                                If i + 2 <= lLen Then
                                    o2 = Mid$(sStream, i + 2, 1)
                                    If o2 >= "0" And o2 <= "7" Then
                                        sO = sO & o2
                                        If i + 3 <= lLen Then
                                            o3 = Mid$(sStream, i + 3, 1)
                                            If o3 >= "0" And o3 <= "7" Then
                                                sO = sO & o3: i = i + 4
                                            Else: i = i + 3
                                            End If
                                        Else: i = i + 3
                                        End If
                                    Else: i = i + 2
                                    End If
                                Else: i = i + 2
                                End If
                                sLit = sLit & Chr(Val("&O" & sO))
                            Else
                                sLit = sLit & cE: i = i + 2
                            End If
                    End Select
                ElseIf cl = "(" Then
                    depth = depth + 1: sLit = sLit & cl: i = i + 1
                ElseIf cl = ")" Then
                    depth = depth - 1
                    If depth > 0 Then sLit = sLit & cl
                    i = i + 1
                Else
                    sLit = sLit & cl: i = i + 1
                End If
            Loop
            If tCount > UBound(tokens) Then
                ReDim Preserve tokens(0 To tCount + 1023)
                ReDim Preserve tokIsHex(0 To tCount + 1023)
            End If
            tokens(tCount) = sLit: tokIsHex(tCount) = False: tCount = tCount + 1
            GoTo NextChar

        Case "<"
            If i + 1 <= lLen Then
                If Mid$(sStream, i + 1, 1) = "<" Then
                    i = i + 2: GoTo NextChar
                End If
            End If
            sHx = ""
            i = i + 1
            Do While i <= lLen
                hc = Mid$(sStream, i, 1)
                If hc = ">" Then i = i + 1: Exit Do
                sHx = sHx & hc: i = i + 1
            Loop
            sHx = Replace(Replace(Replace(sHx, " ", ""), Chr(10), ""), Chr(13), "")
            If tCount > UBound(tokens) Then
                ReDim Preserve tokens(0 To tCount + 1023)
                ReDim Preserve tokIsHex(0 To tCount + 1023)
            End If
            tokens(tCount) = PDF_HexDecode(sHx): tokIsHex(tCount) = True: tCount = tCount + 1
            GoTo NextChar

        Case "T"
            If i + 1 <= lLen Then
                op = Mid$(sStream, i + 1, 1)
                Select Case op
                Case "j"
                    If tCount > 0 Then
                        If Len(sCMap) > 0 And tokIsHex(tCount - 1) Then
                            sRunOut = PDF_ApplyCMap(sCMap, tokens(tCount - 1))
                        Else
                            sRunOut = tokens(tCount - 1)
                        End If
                        If Len(Trim$(sRunOut)) > 0 Then
                            emitY = CLng(curY * 100)
                            result = result & CStr(emitY) & "|" & CStr(CLng(curX * 100)) & "|" & sRunOut & Chr(2)
                        End If
                    End If
                    tCount = 0: i = i + 2: GoTo NextChar
                Case "J"
                    sRunOut = ""
                    For k = 0 To tCount - 1
                        If Len(sCMap) > 0 And tokIsHex(k) Then
                            sRunOut = sRunOut & PDF_ApplyCMap(sCMap, tokens(k))
                        Else
                            sRunOut = sRunOut & tokens(k)
                        End If
                    Next k
                    If Len(Trim$(sRunOut)) > 0 Then
                        emitY = CLng(curY * 100)
                        result = result & CStr(emitY) & "|" & CStr(CLng(curX * 100)) & "|" & sRunOut & Chr(2)
                    End If
                    tCount = 0: i = i + 2: GoTo NextChar
                Case "m"  ' Tm: set text matrix; last two numerics are X Y
                    If Len(sTok) > 0 Then sTok = Left$(sTok, Len(sTok) - 1)  ' strip trailing Chr(3)
                    sTokArr = Split(sTok, Chr(3))
                    nArr = UBound(sTokArr)
                    If nArr >= 1 Then
                        curY = Val(sTokArr(nArr))
                        curX = Val(sTokArr(nArr - 1))
                    End If
                    nTok = 0: sTok = "": tCount = 0
                    i = i + 2: GoTo NextChar
                Case "d", "D"  ' Td/TD: move text position by (dX, dY)
                    If Len(sTok) > 0 Then sTok = Left$(sTok, Len(sTok) - 1)
                    sTokArr = Split(sTok, Chr(3))
                    nArr = UBound(sTokArr)
                    If nArr >= 1 Then
                        curY = curY + Val(sTokArr(nArr))
                        curX = curX + Val(sTokArr(nArr - 1))
                        If op = "D" Then curLead = Abs(Val(sTokArr(nArr)))
                    End If
                    nTok = 0: sTok = "": tCount = 0
                Case "L"  ' TL: set text leading
                    If Len(sTok) > 0 Then sTok = Left$(sTok, Len(sTok) - 1)
                    sTokArr = Split(sTok, Chr(3))
                    nArr = UBound(sTokArr)
                    If nArr >= 0 Then curLead = Abs(Val(sTokArr(nArr)))
                    nTok = 0: sTok = "": tCount = 0
                Case "*"
                    curY = curY - curLead
                    tCount = 0
                Case Else
                    tCount = 0
                End Select
            End If

        Case "'"
            If tCount > 0 Then
                curY = curY - curLead
                If Len(sCMap) > 0 And tokIsHex(tCount - 1) Then
                    sRunOut = PDF_ApplyCMap(sCMap, tokens(tCount - 1))
                Else
                    sRunOut = tokens(tCount - 1)
                End If
                If Len(Trim$(sRunOut)) > 0 Then
                    emitY = CLng(curY * 100)
                    result = result & CStr(emitY) & "|" & CStr(CLng(curX * 100)) & "|" & sRunOut & Chr(2)
                End If
                tCount = 0
            End If
            i = i + 1: GoTo NextChar

        Case Chr(34)  ' " operator: aw ac string " = set word/char spacing, move to next line, show string
            ' aw/ac (word/char spacing) are in sTok - ignored for extraction.
            If tCount > 0 Then
                curY = curY - curLead
                If Len(sCMap) > 0 And tokIsHex(tCount - 1) Then
                    sRunOut = PDF_ApplyCMap(sCMap, tokens(tCount - 1))
                Else
                    sRunOut = tokens(tCount - 1)
                End If
                If Len(Trim$(sRunOut)) > 0 Then
                    emitY = CLng(curY * 100)
                    result = result & CStr(emitY) & "|" & CStr(CLng(curX * 100)) & "|" & sRunOut & Chr(2)
                End If
                tCount = 0
            End If
            i = i + 1: GoTo NextChar

        Case "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
            sNum = ""
            Do While i <= lLen
                nc = Mid$(sStream, i, 1)
                If (nc >= "0" And nc <= "9") Or nc = "." Or (nc = "-" And Len(sNum) = 0) Then
                    sNum = sNum & nc: i = i + 1
                Else
                    Exit Do
                End If
            Loop
            If Len(sNum) > 0 Then
                sTok = sTok & sNum & Chr(3)
                nTok = nTok + 1
                If nTok > 8 Then
                    trimP = InStr(sTok, Chr(3))
                    sTok = Mid$(sTok, trimP + 1)
                    nTok = nTok - 1
                End If
            End If
            GoTo NextChar

        Case Chr(10), Chr(13)
            If tCount > 512 Then tCount = 0

        End Select

        i = i + 1
NextChar:
    Loop

    PDF_ExtractTextOps = result
End Function
Private Function PDF_HexDecode(ByVal sHex As String) As String
    Dim i      As Long
    Dim result As String
    Dim cp     As Long
    Dim b      As Long

    If Len(sHex) Mod 2 = 1 Then sHex = sHex & "0"

    If Len(sHex) >= 4 Then
        If Left$(sHex, 4) = "FEFF" Then  ' UTF-16 BOM  - decode as 2-byte codepoints
            For i = 5 To Len(sHex) - 3 Step 4
                cp = Val("&H" & Mid$(sHex, i, 2)) * 256 + Val("&H" & Mid$(sHex, i + 2, 2))
                If cp > 0 Then result = result & ChrW(cp)
            Next i
            PDF_HexDecode = result
            Exit Function
        End If
    End If

    For i = 1 To Len(sHex) - 1 Step 2
        b = Val("&H" & Mid$(sHex, i, 2))
        result = result & Chr(b)  ' preserve nulls  - 2-byte CID pairs need both bytes
    Next i
    PDF_HexDecode = result
End Function
Private Function PDF_CleanText(ByVal s As String) As String
    Dim i      As Long
    Dim c      As Long
    Dim result As String

    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        Select Case c
            Case 9, 10, 13:    result = result & Chr(c)
            Case 32 To 126:    result = result & Chr(c)
            Case 160 To 65535: result = result & ChrW(c)
        End Select
    Next i

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    Do While InStr(result, Chr(10) & Chr(10) & Chr(10)) > 0
        result = Replace(result, Chr(10) & Chr(10) & Chr(10), Chr(10) & Chr(10))
    Loop

    PDF_CleanText = Trim$(result)
End Function

Private Function PDF_ParseCMap(ByVal sCMap As String) As String
    Dim result  As String
    Dim lPos    As Long
    Dim lEnd    As Long
    Dim sBlock  As String
    Dim lA      As Long
    Dim lAE     As Long
    Dim lB      As Long
    Dim lBE     As Long
    Dim lC      As Long
    Dim lCE     As Long
    Dim srcLo   As Long
    Dim srcHi   As Long
    Dim dstBase As Long
    Dim k       As Long
    Dim sSrc    As String
    Dim sDst    As String

    lPos = 1
    Do
        lPos = InStr(lPos, sCMap, "beginbfchar", vbBinaryCompare)
        If lPos = 0 Then Exit Do
        lEnd = InStr(lPos, sCMap, "endbfchar", vbBinaryCompare)
        If lEnd = 0 Then Exit Do
        sBlock = Mid$(sCMap, lPos + 11, lEnd - lPos - 11)
        lA = 1
        Do
            lA  = InStr(lA,  sBlock, "<", vbBinaryCompare): If lA  = 0 Then Exit Do
            lAE = InStr(lA,  sBlock, ">", vbBinaryCompare): If lAE = 0 Then Exit Do
            lB  = InStr(lAE, sBlock, "<", vbBinaryCompare): If lB  = 0 Then Exit Do
            lBE = InStr(lB,  sBlock, ">", vbBinaryCompare): If lBE = 0 Then Exit Do
            sSrc = Mid$(sBlock, lA  + 1, lAE - lA  - 1)
            sDst = Mid$(sBlock, lB  + 1, lBE - lB  - 1)
            If Len(sSrc) > 0 And Len(sDst) > 0 Then
                If Len(result) > 0 Then result = result & Chr(1)
                result = result & sSrc & ">" & sDst
            End If
            lA = lBE + 1
        Loop
        lPos = lEnd + 9
    Loop

    lPos = 1
    Do
        lPos = InStr(lPos, sCMap, "beginbfrange", vbBinaryCompare)
        If lPos = 0 Then Exit Do
        lEnd = InStr(lPos, sCMap, "endbfrange", vbBinaryCompare)
        If lEnd = 0 Then Exit Do
        sBlock = Mid$(sCMap, lPos + 12, lEnd - lPos - 12)
        lA = 1
        Do
            lA  = InStr(lA,  sBlock, "<", vbBinaryCompare): If lA  = 0 Then Exit Do
            lAE = InStr(lA,  sBlock, ">", vbBinaryCompare): If lAE = 0 Then Exit Do
            lB  = InStr(lAE, sBlock, "<", vbBinaryCompare): If lB  = 0 Then Exit Do
            lBE = InStr(lB,  sBlock, ">", vbBinaryCompare): If lBE = 0 Then Exit Do
            lC  = InStr(lBE, sBlock, "<", vbBinaryCompare): If lC  = 0 Then Exit Do
            lCE = InStr(lC,  sBlock, ">", vbBinaryCompare): If lCE = 0 Then Exit Do
            srcLo   = CLng(Val("&H" & Mid$(sBlock, lA + 1, lAE - lA - 1)))
            srcHi   = CLng(Val("&H" & Mid$(sBlock, lB + 1, lBE - lB - 1)))
            dstBase = CLng(Val("&H" & Mid$(sBlock, lC + 1, lCE - lC - 1)))
            For k = 0 To srcHi - srcLo
                If Len(result) > 0 Then result = result & Chr(1)
                result = result & Hex(srcLo + k) & ">" & Hex(dstBase + k)
            Next k
            lA = lCE + 1
        Loop
        lPos = lEnd + 10
    Loop

    PDF_ParseCMap = result
End Function

Private Function PDF_ApplyCMap(ByVal sCMapData As String, ByVal sRaw As String) As String
    ' Auto-detects 1-byte vs 2-byte CID mode.
    Dim result  As String
    Dim i       As Long
    Dim cid     As Long
    Dim j       As Long
    Dim found   As Boolean
    Dim nPairs  As Long
    Dim pairs() As String
    Dim src()   As Long
    Dim dst()   As String
    Dim sepPos  As Long
    Dim dstHex  As String
    Dim bTwoByte As Boolean

    If Len(sCMapData) = 0 Or Len(sRaw) = 0 Then
        PDF_ApplyCMap = sRaw: Exit Function
    End If

    pairs = Split(sCMapData, Chr(1))
    nPairs = UBound(pairs) + 1
    ReDim src(0 To nPairs - 1)
    ReDim dst(0 To nPairs - 1)
    bTwoByte = False
    For j = 0 To nPairs - 1
        sepPos = InStr(pairs(j), ">")
        If sepPos > 0 Then
            src(j) = CLng(Val("&H" & Left$(pairs(j), sepPos - 1)))
            dstHex = Mid$(pairs(j), sepPos + 1)
            dst(j) = dstHex
            If src(j) > 255 Then bTwoByte = True
        End If
    Next j

    If bTwoByte And Len(sRaw) < 2 Then
        PDF_ApplyCMap = sRaw: Exit Function
    End If

    i = 1
    Do While i <= Len(sRaw)
        If bTwoByte Then
            If i + 1 > Len(sRaw) Then Exit Do
            cid = Asc(Mid$(sRaw, i, 1)) * 256 + Asc(Mid$(sRaw, i + 1, 1))
        Else
            cid = Asc(Mid$(sRaw, i, 1))
        End If
        found = False
        For j = 0 To nPairs - 1
            If src(j) = cid Then
                If Len(dst(j)) = 8 Then
                    result = result & ChrW(Val("&H" & Left$(dst(j), 4))) & ChrW(Val("&H" & Right$(dst(j), 4)))
                ElseIf Len(dst(j)) > 0 Then
                    If Val("&H" & dst(j)) > 0 Then result = result & ChrW(Val("&H" & dst(j)))
                End If
                found = True: Exit For
            End If
        Next j
        If Not found Then
            ' Fallback: emit raw byte if printable ASCII
            If bTwoByte Then
                If cid >= 32 And cid <= 126 Then result = result & Chr(cid)
            Else
                If cid >= 32 And cid <= 255 Then result = result & Chr(cid)
            End If
        End If
        i = i + IIf(bTwoByte, 2, 1)
    Loop
    PDF_ApplyCMap = result
End Function

Private Function PDF_SortAndJoin(ByVal sRuns As String) As String
    ' Y_TOL: runs within 8 pts are on the same line (coords stored *100).
    Const Y_TOL   As Long = 800
    Dim runs()    As String
    Dim n         As Long
    Dim i         As Long
    Dim j         As Long
    Dim arrY()    As Long
    Dim arrX()    As Long
    Dim arrT()    As String
    Dim pipeA     As Long
    Dim pipeB     As Long
    Dim result    As String
    Dim lineY     As Long
    Dim lineText  As String
    Dim swapped   As Boolean
    Dim tmpL      As Long
    Dim tmpS      As String
    Dim doSwap    As Boolean
    Dim maxY      As Long
    Dim bDescend  As Boolean

    If Len(sRuns) = 0 Then PDF_SortAndJoin = "": Exit Function

    runs = Split(sRuns, Chr(2))
    n = 0
    For i = 0 To UBound(runs)
        If Len(Trim$(runs(i))) > 0 Then n = n + 1
    Next i
    If n = 0 Then PDF_SortAndJoin = "": Exit Function

    ReDim arrY(0 To n - 1)
    ReDim arrX(0 To n - 1)
    ReDim arrT(0 To n - 1)

    Dim idx As Long: idx = 0
    For i = 0 To UBound(runs)
        If Len(Trim$(runs(i))) = 0 Then GoTo NextRun
        pipeA = InStr(runs(i), "|")
        If pipeA = 0 Then GoTo NextRun
        pipeB = InStr(pipeA + 1, runs(i), "|")
        If pipeB = 0 Then GoTo NextRun
        arrY(idx) = CLng(Val(Left$(runs(i), pipeA - 1)))
        arrX(idx) = CLng(Val(Mid$(runs(i), pipeA + 1, pipeB - pipeA - 1)))
        arrT(idx) = Mid$(runs(i), pipeB + 1)
        idx = idx + 1
NextRun:
    Next i
    n = idx
    If n = 0 Then PDF_SortAndJoin = "": Exit Function

    ' Y>20000 = standard PDF (bottom-left origin, desc. sort). Y<=20000 = flipped CTM (asc. sort).
    maxY = 0
    For i = 0 To n - 1
        If arrY(i) > maxY Then maxY = arrY(i)
    Next i
    bDescend = (maxY > 20000)   ' 200 points * 100

    Do  ' bubble sort
        swapped = False
        For i = 0 To n - 2
            If bDescend Then
                doSwap = arrY(i) < arrY(i + 1)
            Else
                doSwap = arrY(i) > arrY(i + 1)
            End If
            If arrY(i) = arrY(i + 1) Then
                doSwap = arrX(i) > arrX(i + 1)
            End If
            If doSwap Then
                tmpL = arrY(i): arrY(i) = arrY(i + 1): arrY(i + 1) = tmpL
                tmpL = arrX(i): arrX(i) = arrX(i + 1): arrX(i + 1) = tmpL
                tmpS = arrT(i): arrT(i) = arrT(i + 1): arrT(i + 1) = tmpS
                swapped = True
            End If
        Next i
    Loop While swapped

    lineY = arrY(0)
    lineText = arrT(0)
    For i = 1 To n - 1
        If Abs(arrY(i) - lineY) <= Y_TOL Then
            lineText = lineText & Chr(9) & arrT(i)
        Else
            If Len(result) > 0 Then result = result & Chr(10)
            result = result & lineText
            lineY = arrY(i)
            lineText = arrT(i)
        End If
    Next i
    If Len(lineText) > 0 Then
        If Len(result) > 0 Then result = result & Chr(10)
        result = result & lineText
    End If

    PDF_SortAndJoin = result
End Function

Public Function PDF_DiagnoseStreams(ByVal sFilePath As String) As String
    Dim bFile()      As Byte
    Dim sRaw         As String
    Dim out          As String
    Dim lSearch      As Long
    Dim lPos         As Long
    Dim lStart       As Long
    Dim lEnd         As Long
    Dim lLen         As Long
    Dim n            As Long
    Dim pre3         As String
    Dim lScan        As Long
    Dim lDO          As Long
    Dim lHS          As Long
    Dim sHdr         As String
    Dim isContent    As Boolean
    Dim hasFlate     As Boolean

    bFile = PDF_ReadFileBytes(sFilePath)
    sRaw = PDF_BytesToLatin1(bFile)
    out = "FileSize=" & (UBound(bFile) + 1) & " Encrypted=" & _
          (InStr(1, sRaw, "/Encrypt", vbBinaryCompare) > 0) & vbLf

    lSearch = 1
    Do
        lPos = InStr(lSearch, sRaw, "stream", vbBinaryCompare)
        If lPos = 0 Then Exit Do
        If lPos >= 4 Then pre3 = Mid$(sRaw, lPos - 3, 3) Else pre3 = ""
        If pre3 = "end" Then
            lSearch = lPos + 6
        Else
            n = n + 1
            lStart = lPos + 6
            Do While lStart <= Len(sRaw)
                If bFile(lStart - 1) = 32 Or bFile(lStart - 1) = 9 Then
                    lStart = lStart + 1
                Else
                    Exit Do
                End If
            Loop
            If bFile(lStart - 1) = 13 Then lStart = lStart + 1
            If bFile(lStart - 1) = 10 Then lStart = lStart + 1
            lEnd = InStr(lStart, sRaw, "endstream", vbBinaryCompare)
            If lEnd = 0 Then out = out & "#" & n & " no endstream" & vbLf: Exit Do
            lLen = lEnd - lStart

            lScan = lPos - 1: lDO = 0
            Do While lScan > 0
                If Mid$(sRaw, lScan, 2) = "<<" Then lDO = lScan: Exit Do
                If Mid$(sRaw, lScan, 6) = "endobj" Then Exit Do
                lScan = lScan - 1
            Loop
            If lDO = 0 Then lHS = IIf(lPos - 512 < 1, 1, lPos - 512) Else lHS = lDO
            sHdr = Mid$(sRaw, lHS, lPos - lHS)

            isContent = PDF_IsContentStream(sHdr)
            hasFlate = InStr(1, sHdr, "/FlateDecode", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdr, "/Fl ", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdr, "/Fl>", vbBinaryCompare) > 0 Or _
                       InStr(1, sHdr, "/Fl/", vbBinaryCompare) > 0
            Dim isObjStm As Boolean
            isObjStm = InStr(1, sHdr, "/Type /ObjStm",  vbBinaryCompare) > 0 Or _
                       InStr(1, sHdr, "/Type/ObjStm",   vbBinaryCompare) > 0

            out = out & "#" & n & " pos=" & lPos & " len=" & lLen & _
                        " content=" & isContent & " flate=" & hasFlate & _
                        " objstm=" & isObjStm & vbLf
            out = out & "  hdr=[" & Left$(sHdr, 100) & "]" & vbLf

            lSearch = lEnd + 9
        End If
    Loop
    PDF_DiagnoseStreams = out
End Function
