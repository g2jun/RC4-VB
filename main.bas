Option Explicit

Sub Text2BytArray(ByVal strFName As String, bytArray() As Byte)
    
    Dim intFNo As Integer
    intFNo = FreeFile
    
    Open strFName For Binary As #intFNo
    ReDim bytArray(LOF(intFNo) - 1)
    Get #intFNo, , bytArray()
    Close #intFNo
    
End Sub

Sub BytArray2Text(ByVal strFName As String, bytArray() As Byte)
    
    Dim intFNo As Integer
    intFNo = FreeFile
    
    Open strFName For Output As #intFNo
    Close #intFNo
    
    intFNo = FreeFile
    Open strFName For Binary As #intFNo
    Put #intFNo, , bytArray()
    Close #intFNo
    
End Sub

Sub Hex2BytArray(ByVal strFName As String, bytArray() As Byte)
    
    Dim i As Long
    Dim j As Integer
    Dim strHex As String
    Dim lonLen As Long
    Dim intFNo As Integer
    Dim strHex4Byte As String
    
    intFNo = FreeFile
    Open strFName For Input As #intFNo
    
    Line Input #intFNo, strHex
    lonLen = Len(strHex)
    
    j = 0
    For i = 1 To lonLen - 1 Step 2
        strHex4Byte = Mid(strHex, i, 2)
        ReDim Preserve bytArray(j)
        bytArray(j) = CByte("&H" & strHex4Byte)
        j = j + 1
    Next i
        
    Close #intFNo
    
End Sub

Sub BytArray2Hex(ByVal strFName As String, bytArray() As Byte)
    
    Dim i As Long
    Dim strHex As String
    Dim intFNo As Integer
    
    For i = LBound(bytArray()) To UBound(bytArray) Step 1
        'strHex = strHex & CStr(bytArray(i)) & " "
        strHex = strHex & Right("0" & Hex(bytArray(i)), 2)
    Next i
    
    intFNo = FreeFile
    Open strFName For Output As #intFNo
    Print #intFNo, strHex
    Close #intFNo
    
End Sub

Sub Crypt(bytPlainArray() As Byte, bytKeyArray() As Byte, bytCipherArray() As Byte, bytKeyStreamArray() As Byte)
    
    Dim i As Integer
    Dim j As Integer
    Dim S(255) As Integer
    Dim intKeyLen As Integer
    Dim intTemp As Integer
    Dim m As Integer
    Dim intPlainLen As Integer
    Dim k As Integer
    
    intKeyLen = UBound(bytKeyArray) - LBound(bytKeyArray) + 1
    intPlainLen = UBound(bytPlainArray) - LBound(bytPlainArray) + 1
    ReDim bytCipherArray(intPlainLen - 1)
    ReDim bytKeyStreamArray(intPlainLen - 1)
    
    For i = 0 To 255 Step 1
        S(i) = i
    Next i
    
    j = 0
    For i = 0 To 255 Step 1
        j = (j + S(i) + bytKeyArray(i Mod intKeyLen)) Mod 256
        intTemp = S(i)
        S(i) = S(j)
        S(j) = intTemp
    Next i
    
    i = 0
    j = 0
    For m = 0 To intPlainLen - 1 Step 1
        i = (i + 1) Mod 256
        j = (j + S(i)) Mod 256
        intTemp = S(i)
        S(i) = S(j)
        S(j) = intTemp
        k = S((S(i) + S(j)) Mod 256)
        bytKeyStreamArray(m) = k
        bytCipherArray(m) = CByte(bytPlainArray(m) Xor k)
    Next m
    
End Sub

Sub main()
    Dim bytPlainArray() As Byte
    Dim bytKeyArray() As Byte
    Dim bytCipherArray() As Byte
    Dim bytKeyStreamArray() As Byte
        
    Dim strFName4Plain As String
    Dim strFName4Key As String
    Dim strFName4Hex As String
    Dim strFName4KeyStream As String
    
    Dim n As Integer
    
    strFName4Plain = "E:\VBSamples\RC4\Plain.txt"
    strFName4Key = "E:\VBSamples\RC4\Key.txt"
    strFName4Hex = "E:\VBSamples\RC4\Hex.txt"
    strFName4KeyStream = "E:\VBSamples\RC4\KeyStream.txt"
    
    Text2BytArray strFName:=strFName4Key, bytArray:=bytKeyArray()
    
    n = InputBox(Prompt:="Encrypt: 1 Decrypt: 2")
    If n = 1 Then
        Text2BytArray strFName:=strFName4Plain, bytArray:=bytPlainArray()
        Crypt bytPlainArray:=bytPlainArray(), bytKeyArray:=bytKeyArray(), bytCipherArray:=bytCipherArray(), bytKeyStreamArray:=bytKeyStreamArray()
        BytArray2Hex strFName:=strFName4Hex, bytArray:=bytCipherArray()
    ElseIf n = 2 Then
        Hex2BytArray strFName:=strFName4Hex, bytArray:=bytCipherArray()
        Crypt bytPlainArray:=bytCipherArray(), bytKeyArray:=bytKeyArray(), bytCipherArray:=bytPlainArray(), bytKeyStreamArray:=bytKeyStreamArray()
        BytArray2Text strFName:=strFName4Plain, bytArray:=bytPlainArray()
    Else
        Exit Sub
    End If
    
    BytArray2Hex strFName:=strFName4KeyStream, bytArray:=bytKeyStreamArray()
        
End Sub
