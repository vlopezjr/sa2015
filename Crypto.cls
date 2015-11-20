VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Crypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Microsoft VB4 Developer's Workshop (Chapter 14 & 30), John Craig
' slight modification to extend the range of encryption


Private Const BIGNUM As Long = 32768


Public Function encrypt(plaintext As String) As String
    Dim n1, n2, n3
    Dim s As String
    Dim i As Integer
    
    s = plaintext
    For i = 1 To Len(g_CryptoPassword)
        n1 = n1 + Asc(Mid$(g_CryptoPassword, i, 1))
        n1 = (n1 * 367 + 331) Mod &HFFF
        n2 = ((n2 + n1) * 743 + 599) Mod &HFF
        n3 = ((n3 + n2) * 563 + 787) Mod &HFFF
    Next i
    Cipher s, n1, n2, n3
    encrypt = s
End Function


Public Function decrypt(ciphertext As String) As String
    Dim n1, n2, n3
    Dim s As String
    Dim i As Integer
    
    s = ciphertext
    For i = 1 To Len(g_CryptoPassword)
        n1 = n1 + Asc(Mid$(g_CryptoPassword, i, 1))
        n1 = (n1 * 367 + 331) Mod &HFFF
        n2 = ((n2 + n1) * 743 + 599) Mod &HFF
        n3 = ((n3 + n2) * 563 + 787) Mod &HFFF
    Next i
    Cipher s, n1, n2, n3
    decrypt = s
End Function


Private Sub Cipher(Txt As String, Optional Rvalue As Variant, Optional A As Variant, Optional B As Variant)
    Static R As Long
    Static M As Long
    Static N As Long
    Dim i As Long, c As Long, d As Long
    
    If IsMissing(Rvalue) = False Then
        R = Rvalue
    End If
    
    If IsMissing(A) Then
        If M = 0 Then M = 69
    Else
        M = (A * 4 + 1) Mod BIGNUM
    End If
    
    If IsMissing(B) Then
        If N = 0 Then N = 47
    Else
        N = (B * 2 + 1) Mod BIGNUM
    End If
    
    For i = 1 To Len(Txt)
        c = Asc(Mid$(Txt, i, 1))
        Select Case c
            Case 40 To 57   ' numeric (extended to include the 8 ASCII characters before "0")
                d = c - 40
            Case 63 To 90   ' ?, @, A-Z
                d = c - 45
            Case 97 To 122  ' a-z
                d = c - 51
            Case Else
                d = -1
        End Select
        If d >= 0 Then
            R = (R * M + N) Mod BIGNUM
            d = (R And 63) Xor d
            Select Case d
                Case 0 To 17
                    c = d + 40
                Case 18 To 45
                    c = d + 45
                Case 46 To 71
                    c = d + 51
            End Select
            Mid$(Txt, i, 1) = chr$(c)
        End If
    Next i
End Sub


