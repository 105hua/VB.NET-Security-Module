Imports System.Security.Cryptography
Imports System.Text
Module securityModule

    Public Function MD5Hash(ByVal stringToHash As String)
        Using hash As MD5 = MD5.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the MD5 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new string builder to convert the MD5 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the MD5 Hash into the string builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2"))
            Next
            Return stringBuilder.ToString() ' This function will return the stringbuilder as a string.
        End Using
    End Function

    Public Function SHA1Hash(ByVal stringToHash As String)
        Using hash As SHA1 = SHA1.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-1 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new string builder to convert the SHA-1 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-1 Hash into the string builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2"))
            Next
            Return stringBuilder.ToString() ' This function will return the stringbuilder as a string.
        End Using
    End Function

    Public Function SHA256Hash(ByVal stringToHash As String)
        Using hash As SHA256 = SHA256.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-256 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new string builder to convert the SHA-256 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-256 Hash into the string builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2"))
            Next
            Return stringBuilder.ToString() ' This function will return the stringbuilder as a string.
        End Using
    End Function

    Public Function SHA384Hash(ByVal stringToHash As String)
        Using hash As SHA384 = SHA384.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-384 Hash as Bytes.
            Dim stringBuilder As New StringBuilder
            For i As Integer = 0 To stringtohashBytes.Length - 1
                stringBuilder.Append(stringtohashBytes(i).ToString("X2"))
            Next
            Return stringBuilder.ToString()
        End Using
    End Function

    Public Function SHA512Hash(ByVal stringToHash As String)
        Using hash As SHA512 = SHA512.Create()
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash))
            Dim stringBuilder As New StringBuilder
            For i As Integer = 0 To stringtohashBytes.Length - 1
                stringBuilder.Append(stringtohashBytes(i).ToString("X2"))
            Next
            Return stringBuilder.ToString()
        End Using
    End Function

End Module
