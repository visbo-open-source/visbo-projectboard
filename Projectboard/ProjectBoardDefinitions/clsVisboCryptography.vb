Imports System.Security.Cryptography
Public NotInheritable Class clsVisboCryptography

    Private TripleDes As New TripleDESCryptoServiceProvider


    ''' <summary>
    ''' erzeugt einen verschlüsselten String aus Username, Pwd
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <param name="pwd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function verschluessleUserPwd(ByVal userName As String, ByVal pwd As String) As String

        Dim completeString As String = My.Computer.Name & vbLf & _
                                       "visbo" & vbLf & _
                                       userName & vbLf & _
                                       "h0lzk1rch3n" & vbLf & _
                                       pwd

        Dim verschluesselterString = Me.EncryptData(completeString)
        verschluessleUserPwd = verschluesselterString

    End Function

    ''' <summary>
    ''' gibt von dem verschlüsselten UserName PWD den Username zurück
    ''' </summary>
    ''' <param name="verschluesselterText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getUserNameFromCipher(ByVal verschluesselterText As String) As String

        Dim tmpName As String = ""
        Dim completeString As String = Me.DecryptData(verschluesselterText)
        Dim tmpStr() As String = completeString.Split(New Char() {CChar(vbLf)})

        If tmpStr.Length = 5 Then
            ' alles in Ordnung 
            If tmpStr(0) = My.Computer.Name And tmpStr(1) = "visbo" And tmpStr(3) = "h0lzk1rch3n" Then
                ' es kann weitergemacht werden 
                tmpName = tmpStr(2)
            End If

        End If

        getUserNameFromCipher = tmpName

    End Function

    ''' <summary>
    ''' gibt von dem verschlüsselten UserName PWD das pwd zurück
    ''' </summary>
    ''' <param name="verschluesselterText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getPwdFromCipher(ByVal verschluesselterText As String) As String

        Dim tmpPWD As String = ""
        Dim completeString As String = Me.DecryptData(verschluesselterText)
        Dim tmpStr() As String = completeString.Split(New Char() {CChar(vbLf)})

        If tmpStr.Length = 5 Then
            ' alles in Ordnung 
            If tmpStr(0) = My.Computer.Name And tmpStr(1) = "visbo" And tmpStr(3) = "h0lzk1rch3n" Then
                ' es kann weitergemacht werden 
                tmpPWD = tmpStr(4)
            End If

        End If

        getPwdFromCipher = tmpPWD

    End Function

    ''' <summary>
    ''' method that creates a byte array of a specified length from the hash of the specified key.
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key.
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(length - 1)
        Return hash
    End Function

    ''' <summary>
    ''' method that encrypts a string.
    ''' </summary>
    ''' <param name="plaintext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EncryptData(ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim plaintextBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function

    ''' <summary>
    ''' method that decrypts a string.
    ''' </summary>
    ''' <param name="encryptedtext"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function DecryptData(ByVal encryptedtext As String) As String

        ' Convert the encrypted text string to a byte array.
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

    ''' <summary>
    ''' constructor to initialize the 3DES cryptographic service provider.
    ''' </summary>
    ''' <remarks></remarks>
    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

End Class
