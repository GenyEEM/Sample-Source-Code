Attribute VB_Name = "MdlShare"
Option Explicit

Public IsStop As Boolean    '是否停止运行标志

Public Declare Function GetTickCount Lib "kernel32" () As Long   '得到系统时间

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliSeconds As Long)  '使系统休眠时间（毫秒）

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal Length As Long)

Declare Function WritePrivateProfileString _
        Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString _
        Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long

Public ProgramPath As String

Dim i(15)          As Integer

Dim m(15)          As Long

Public blnStop    As Boolean

Public Function ByteToSingie(intByte1 As Integer, _
                             intByte2 As Integer, _
                             intByte3 As Integer, _
                             intByte4 As Integer, _
                             Optional intLength As Integer = 10) As Single
    Dim sngData         As Single
    Dim bytData(0 To 3) As Byte
    Dim varTmp          As Variant

    bytData(0) = intByte1
    bytData(1) = intByte2
    bytData(2) = intByte3
    bytData(3) = intByte4

    CopyMemory sngData, bytData(0), 4

    'varTmp = CDec(sngData)

    ByteToSingie = Left$(Val(sngData), intLength)
End Function

Public Sub Delay(dblTimeS As Double, Optional blnIsRun As Boolean = True)
    Dim start1 As Double
    Dim s2     As Double

restat:
    start1 = GetTickCount

    Do
        DoEvents

        If GetTickCount - start1 < 0 Then
            GoTo restat
        End If

        Sleep 5
        s2 = GetTickCount

        If blnIsRun = False Or blnStop Then Exit Sub
    Loop Until (s2 - start1) >= dblTimeS * 1000

End Sub

Public Function getHexData(strSource As String) As String
    Dim strData()  As String

    Dim i          As Integer

    Dim strTmp     As String

    Dim strGetData As String

    strData = Split(Trim$(strSource))

    For i = 0 To UBound(strData)

        If Len(strData(i)) Then

            Select Case Left$(strData(i), 1)

                Case "d"
                    strTmp = dblToHex(CDbl(Mid$(strData(i), 2)))

                Case "f"
                    strTmp = sngToHex(CSng(Mid$(strData(i), 2)))

                Case "i"
                    strTmp = intToHex(CInt(Mid$(strData(i), 2)))

                Case "l"
                    strTmp = lngToHex(CLng(Mid$(strData(i), 2)))

                Case Else
                    strTmp = Format$(strData(i), "00")
            End Select

        End If

        strGetData = Trim$(strGetData & " " & strTmp)
    Next i

    getHexData = strGetData
End Function

Function GetIniInfo(ByVal filename As String, _
                    ByVal Section As String, _
                    ByVal KeyName As String, _
                    Optional ByVal Default As Variant, _
                    Optional ByVal ByValue As Boolean) As Variant
    Dim strDefault As String
    Dim Result     As String
    Dim ValueLen   As Long
    Dim MSG        As String
    On Error Resume Next
    strDefault = Default
    ValueLen = 4096
    Result = Space$(ValueLen)
    ValueLen = GetPrivateProfileString(Section, KeyName, strDefault, Result, ValueLen, filename)

    If ByValue Then
        GetIniInfo = Val(Result)
    Else
        Result = Trim$(Result)

        If Asc(Right$(Result, 1)) = 0 Then Result = Left$(Result, Len(Result) - 1)
        GetIniInfo = Trim$(Result)
    End If

End Function

Function GetItem(ByVal MSG As String, _
                 ByVal Split As String, _
                 ByVal Index As Long, _
                 Optional ByVal ByValue As Boolean) As Variant
    '取指定项,EX: GetItem("1A,5A,10A,20A",",",2) = "10A"
    'Index = -1 , Get Items Count
    Dim SplitLen As Long
    Dim S        As Long
    Dim N        As Long
    Dim Count    As Long
    Dim item     As String

    SplitLen = Len(Split)

    If Len(MSG) * SplitLen > 0 Then   '有效的字符串和分隔符
        S = 1

        If Index < 0 Then   '取项数

            Do
                N = InStr(S, MSG, Split)
                Count = Count + 1

                If N > 0 Then S = N + SplitLen
            Loop Until (N = 0)

            GetItem = Count
        Else                '取指定项

            Do
                N = InStr(S, MSG, Split)

                If Count = Index Then
                    item = Mid$(MSG, S, IIf(N = 0, Len(MSG), N - S))
                    Exit Do
                Else
                    Count = Count + 1

                    If N > 0 Then S = N + SplitLen
                End If

            Loop Until (N = 0)

            GetItem = IIf(ByValue, Val(item), item)
        End If
    End If

End Function

Function GetItemNo(ByVal MSG As String, Split As String, item As String) As Long
    '取指定项的序号(0..N),找不到返回-1
    Dim SplitLen As Long
    Dim S        As Long
    Dim N        As Long
    Dim Count    As Long

    GetItemNo = -1
    SplitLen = Len(Split)

    If SplitLen > 0 Then  '有效的分隔符
        S = 1

        Do
            N = InStr(S, MSG, Split)

            If N = 0 Then
                If Mid$(MSG, S) = item Then GetItemNo = Count
            Else

                If Mid$(MSG, S, N - S) = item Then GetItemNo = Count
                S = N + SplitLen
                Count = Count + 1
            End If

        Loop Until (N = 0)

    End If

End Function

Public Function getReceiveData(bytData() As Byte) As String

    Dim i      As Integer

    Dim strTmp As String

    '将字符串转换为字符型数组
    '将十进制转换为十六进制
    For i = 0 To UBound(bytData)
        strTmp = strTmp & Format$(Hex$(bytData(i)), "00") & " "
    Next i

    getReceiveData = strTmp
End Function

Public Sub getSendByte(bytData() As Byte, strSource As String)

    Dim strData() As String

    Dim i         As Integer

    '将字符串转换为字符型数组
    strData = Split(strSource)
    '重新定义数组长度
    ReDim bytData(0 To UBound(strData)) As Byte

    '将十六进制转换为十进制
    For i = 0 To UBound(strData)
        bytData(i) = "&H" & (strData(i))
    Next i

End Sub

Function SaveIniInfo(ByVal filename As String, _
                     ByVal Section As String, _
                     ByVal KeyName As String, _
                     ByVal Value As Variant)
    On Error Resume Next
    Dim strValue As String
    strValue = Value
    WritePrivateProfileString Section, KeyName, strValue, filename
End Function

Public Function TenToEighteen(ByVal ten As Long) As String   '十进制转八进制

    Dim i, j, x As Integer '除8后的商

    Dim ix As Integer  '除8后的余数

    Dim iy As Integer  '余数除以8后的商

    Dim iz As Integer  '除数除以8后的余数

    If ten >= 64 Then
        j = Int(Int(ten / 64) / 8)
        x = Int(ten / 64) Mod 8

        If j Mod 8 = 0 Then
            i = (j / 8) * 10 * 10 + x
        Else
            i = (Int(j / 8) * 10 + j Mod 8) * 10 + x
        End If

        ix = ten Mod 64
        iy = Int(ix / 8)
        iz = ix Mod 8
        TenToEighteen = i * 100 + iy * 10 + iz
    ElseIf 8 < ten < 64 Then
        iy = Int(ten / 8)
        iz = ten Mod 8
        TenToEighteen = iy * 10 + iz
    Else
        TenToEighteen = ten
    End If

End Function

Public Function TenToSixteen(ten As Long) As String     '十进制转十六进制

    Dim i As Integer

    Dim j As Integer

    Dim m As Integer

    Dim x As Long

    Dim y As String

    Dim z As String

    Dim W As String

    i = Int(Len(TenToTwo(ten)) / 4)
    j = Len(TenToTwo(ten)) Mod 4
    z = ""

    For m = 1 To i
        x = Mid$(TenToTwo(ten), Len(TenToTwo(ten)) - (m * 4) + 1, 4)
        y = TwoToTen(x)

        Select Case y

            Case 10
                y = "A"

            Case 11
                y = "B"

            Case 12
                y = "C"

            Case 13
                y = "D"

            Case 14
                y = "E"

            Case 15
                y = "F"
        End Select

        z = y & z
    Next m

    If j = 0 Then
        TenToSixteen = z
    Else
        W = TwoToTen(Mid$(TenToTwo(ten), 1, j))
        TenToSixteen = W & z
    End If

End Function

Public Function TenToTwo(ten As Long) As String     '十进制转二进制
    m(0) = 0
    m(1) = 0
    m(2) = 0
    m(3) = 0
    m(4) = 0
    m(5) = 0
    m(6) = 0
    m(7) = 0
    m(8) = 0
    m(9) = 0
    m(10) = 0
    m(11) = 0
    m(12) = 0
    i(0) = 0
    i(1) = 0
    i(2) = 0
    i(3) = 0
    i(4) = 0
    i(5) = 0
    i(6) = 0
    i(7) = 0
    i(8) = 0
    i(9) = 0
    i(10) = 0
    i(11) = 0
    i(12) = 0

    Do While 2 ^ (i(0) + 1) <= ten
        i(0) = i(0) + 1
    Loop

    m(0) = ten - 2 ^ i(0)

    If m(0) = 0 Then
        TenToTwo = 10 ^ i(0)
    Else

        Do While 2 ^ (i(1) + 1) <= m(0)
            i(1) = i(1) + 1
        Loop

        m(1) = m(0) - 2 ^ (i(1))

        '*****************************
        If m(1) = 0 Then
            TenToTwo = 10 ^ i(0) + 10 ^ i(1)
        Else

            Do While 2 ^ (i(2) + 1) <= m(1)
                i(2) = i(2) + 1
            Loop

            m(2) = m(1) - 2 ^ (i(2))

            '*****************************
            If m(2) = 0 Then
                TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2)
            Else

                Do While 2 ^ (i(3) + 1) <= m(2)
                    i(3) = i(3) + 1
                Loop

                m(3) = m(2) - 2 ^ (i(3))

                '*****************************
                If m(3) = 0 Then
                    TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3)
                Else

                    Do While 2 ^ (i(4) + 1) <= m(3)
                        i(4) = i(4) + 1
                    Loop

                    m(4) = m(3) - 2 ^ (i(4))

                    '*****************************
                    If m(4) = 0 Then
                        TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4)
                    Else

                        Do While 2 ^ (i(5) + 1) <= m(4)
                            i(5) = i(5) + 1
                        Loop

                        m(5) = m(4) - 2 ^ (i(5))

                        '*****************************
                        If m(5) = 0 Then
                            TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5)
                        Else

                            Do While 2 ^ (i(6) + 1) <= m(5)
                                i(6) = i(6) + 1
                            Loop

                            m(6) = m(5) - 2 ^ (i(6))

                            '*****************************
                            If m(6) = 0 Then
                                TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6)
                            Else

                                Do While 2 ^ (i(7) + 1) <= m(6)
                                    i(7) = i(7) + 1
                                Loop

                                m(7) = m(6) - 2 ^ (i(7))

                                '*****************************
                                If m(7) = 0 Then
                                    TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7)
                                Else

                                    Do While 2 ^ (i(8) + 1) <= m(7)
                                        i(8) = i(8) + 1
                                    Loop

                                    m(8) = m(7) - 2 ^ (i(8))

                                    '*****************************
                                    If m(8) = 0 Then
                                        TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7) + 10 ^ i(8)
                                    Else

                                        Do While 2 ^ (i(9) + 1) <= m(8)
                                            i(9) = i(9) + 1
                                        Loop

                                        m(9) = m(8) - 2 ^ (i(9))

                                        '*****************************
                                        If m(9) = 0 Then
                                            TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7) + 10 ^ i(8) + 10 ^ i(9)
                                        Else

                                            Do While 2 ^ (i(10) + 1) <= m(9)
                                                i(10) = i(10) + 1
                                            Loop

                                            m(10) = m(9) - 2 ^ (i(10))

                                            '*****************************
                                            If m(10) = 0 Then
                                                TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7) + 10 ^ i(8) + 10 ^ i(9) + 10 ^ i(10)
                                            Else

                                                Do While 2 ^ (i(11) + 1) <= m(10)
                                                    i(11) = i(11) + 1
                                                Loop

                                                m(11) = m(10) - 2 ^ (i(11))

                                                '*****************************
                                                If m(11) = 0 Then
                                                    TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7) + 10 ^ i(8) + 10 ^ i(9) + 10 ^ i(10) + 10 ^ i(11)
                                                Else

                                                    Do While 2 ^ (i(12) + 1) <= m(10)
                                                        i(12) = i(12) + 1
                                                    Loop

                                                    m(12) = m(11) - 2 ^ (i(12))

                                                    '*****************************
                                                    If m(12) = 0 Then
                                                        TenToTwo = 10 ^ i(0) + 10 ^ i(1) + 10 ^ i(2) + 10 ^ i(3) + 10 ^ i(4) + 10 ^ i(5) + 10 ^ i(6) + 10 ^ i(7) + 10 ^ i(8) + 10 ^ i(9) + 10 ^ i(10) + 10 ^ i(11) + 10 ^ i(12)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Public Function TwoToTen(ten) As Long     '二进制转十进制

    Dim a    As Integer

    Dim N    As Integer

    Static W As Long

    W = 0
    a = Len(ten)

    For N = 1 To a
        W = W + (Mid$(ten, N, 1) * 2 ^ (a - N))
    Next N

    TwoToTen = W
End Function

Private Function dblToHex(dblData As Double) As String
    Dim i         As Integer
    Dim hexData   As String
    Dim Buffer(7) As Byte

    CopyMemory Buffer(0), dblData, 8

    '低字节在前
    For i = 7 To 0 Step -1

        If Len(Hex$(Buffer(i))) = 1 Then
            hexData = "0" & Hex$(Buffer(i)) + " " + hexData
        Else
            hexData = Hex$(Buffer(i)) + " " + hexData
        End If

    Next

    dblToHex = hexData
End Function

Private Function intToHex(intData As Integer) As String
    Dim i         As Integer
    Dim hexData   As String
    Dim Buffer(1) As Byte

    CopyMemory Buffer(0), intData, 2

    '低字节在前
    For i = 1 To 0 Step -1

        If Len(Hex$(Buffer(i))) = 1 Then
            hexData = "0" & Hex$(Buffer(i)) + " " + hexData
        Else
            hexData = Hex$(Buffer(i)) + " " + hexData
        End If

    Next

    intToHex = hexData
End Function

Private Function lngToHex(lngData As Long) As String
    Dim i         As Integer
    Dim hexData   As String
    Dim Buffer(3) As Byte

    CopyMemory Buffer(0), lngData, 4

    '低字节在前
    For i = 3 To 0 Step -1

        If Len(Hex$(Buffer(i))) = 1 Then
            hexData = "0" & Hex$(Buffer(i)) + " " + hexData
        Else
            hexData = Hex$(Buffer(i)) + " " + hexData
        End If

    Next

    lngToHex = hexData
End Function

Private Function sngToHex(sngData As Single) As String
    Dim i         As Integer
    Dim hexData   As String
    Dim Buffer(3) As Byte

    CopyMemory Buffer(0), sngData, 4

    '低字节在前
    For i = 3 To 0 Step -1

        If Len(Hex$(Buffer(i))) = 1 Then
            hexData = "0" & Hex$(Buffer(i)) + " " + hexData
        Else
            hexData = Hex$(Buffer(i)) + " " + hexData
        End If

    Next

    sngToHex = hexData
End Function

Public Function SecondToTime(fSecond As Double) As String
    '秒时间 -->“时时分分秒秒” 格式表示的字符串
    Dim hh As String, mm As String, ss As String
    hh = Trim$(Str(Fix(Val(fSecond) / 3600)))
    mm = Trim$(Str(Fix((Val(fSecond) - Val(hh) * 3600) / 60)))
    ss = Trim$(Str((Val(fSecond) - Val(hh) * 3600) Mod 60))

    If Len(hh) = 1 Then hh = "0" & hh
    If Len(mm) = 1 Then mm = "0" & mm
    If Len(ss) = 1 Then ss = "0" & ss
    If Len(hh) > 2 Then hh = Right$(hh, 2)
    SecondToTime = hh & ":" & mm & ":" & ss
End Function

Sub LogWrite(Text As String, filename As String, Optional AppendMode As Boolean)
    Dim fnum As Integer, isOpen As Boolean
    On Error GoTo Error_Handler
    
    fnum = FreeFile()
    If AppendMode Then
        Open filename For Append As #fnum
    Else
        Open filename For Output As #fnum
    End If
    isOpen = True
    Print #fnum, Text
Error_Handler:
    If isOpen Then Close #fnum
    If Err Then Err.Raise Err.Number, , Err.Description
End Sub
