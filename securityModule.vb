Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports SHA3


Module securityModule

    Public Function MD5Hash(ByVal stringToHash As String)
        Using hash As MD5 = MD5.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the MD5 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new String Builder to convert the MD5 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the MD5 Hash into the string builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2")) ' This will append each character of the Byte() variable that holds the hash into the String Builder
            Next
            Return stringBuilder.ToString() ' This function will return the String Builder as a string.
        End Using
    End Function

    Public Function SHA1Hash(ByVal stringToHash As String)
        Using hash As SHA1 = SHA1.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-1 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new String Builder to convert the SHA-1 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-1 Hash into the String Builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2")) ' This will append each character of the Byte() variable that holds the hash into the String Builder
            Next
            Return stringBuilder.ToString() ' This function will return the String Builder as a string.
        End Using
    End Function

    Public Function SHA2_256Hash(ByVal stringToHash As String)
        Using hash As SHA256 = SHA256.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-256 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new String Builder to convert the SHA-256 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-256 Hash into the String Builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2")) ' This will append each character of the Byte() variable that holds the hash into the String Builder
            Next
            Return stringBuilder.ToString() ' This function will return the String Builder as a string.
        End Using
    End Function

    Public Function SHA2_384Hash(ByVal stringToHash As String)
        Using hash As SHA384 = SHA384.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-384 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new String Builder to convert the SHA-384 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-256 Hash into the String Builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2")) ' This will append each character of the Byte() variable that holds the hash into the String Builder
            Next
            Return stringBuilder.ToString() ' This function will return the String Builder as a string.
        End Using
    End Function

    Public Function SHA2_512Hash(ByVal stringToHash As String)
        Using hash As SHA512 = SHA512.Create() ' Create the hashing object for use in computing the hash.
            Dim stringtohashBytes As Byte() = hash.ComputeHash(Encoding.UTF8.GetBytes(stringToHash)) ' Compute the SHA-512 Hash as Bytes.
            Dim stringBuilder As New StringBuilder ' Define a new String Builder to convert the SHA-512 Hash into a string.
            For i As Integer = 0 To stringtohashBytes.Length - 1 ' In a for loop, we append each character of the SHA-512 Hash into the String Builder
                stringBuilder.Append(stringtohashBytes(i).ToString("X2")) ' This will append each character of the Byte() variable that holds the hash into the String Builder
            Next
            Return stringBuilder.ToString() ' This function will return the String Builder as a string.
        End Using
    End Function

    Public Function SHA3_224_Hash(ByVal stringToHash As String)
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(stringToHash) ' Seperate each character of the string that we are hashing into a Byte array
        Dim hasher As New SHA3.SHA3Managed(224) ' Define the SHA3Managed variable that will be computing our hash
        Dim hashedString As Byte() = hasher.ComputeHash(bytes) ' Compute the hash and store it as a new Byte array
        Dim stringBuilder As New StringBuilder ' Define a string builder so we can convert back to a string
        For i As Integer = 0 To hashedString.Length - 1
            stringBuilder.Append(hashedString(i).ToString("X2")) ' Append each character of the hash to the string builder
        Next
        Return stringBuilder.ToString ' Return the string builder converted to a string
    End Function

    Public Function SHA3_256_Hash(ByVal stringToHash As String)
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(stringToHash) ' Seperate each character of the string that we are hashing into a Byte array
        Dim hasher As New SHA3.SHA3Managed(256) ' Define the SHA3Managed variable that will be computing our hash
        Dim hashedString As Byte() = hasher.ComputeHash(bytes) ' Compute the hash and store it as a new Byte array
        Dim stringBuilder As New StringBuilder ' Define a string builder so we can convert back to a string
        For i As Integer = 0 To hashedString.Length - 1
            stringBuilder.Append(hashedString(i).ToString("X2")) ' Append each character of the hash to the string builder
        Next
        Return stringBuilder.ToString ' Return the string builder converted to a string
    End Function

    Public Function SHA3_384_Hash(ByVal stringToHash As String)
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(stringToHash) ' Seperate each character of the string that we are hashing into a Byte array
        Dim hasher As New SHA3.SHA3Managed(384) ' Define the SHA3Managed variable that will be computing our hash
        Dim hashedString As Byte() = hasher.ComputeHash(bytes) ' Compute the hash and store it as a new Byte array
        Dim stringBuilder As New StringBuilder ' Define a string builder so we can convert back to a string
        For i As Integer = 0 To hashedString.Length - 1
            stringBuilder.Append(hashedString(i).ToString("X2")) ' Append each character of the hash to the string builder
        Next
        Return stringBuilder.ToString ' Return the string builder converted to a string
    End Function

    Public Function SHA3_512_Hash(ByVal stringToHash As String)
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(stringToHash) ' Seperate each character of the string that we are hashing into a Byte array
        Dim hasher As New SHA3.SHA3Managed(512) ' Define the SHA3Managed variable that will be computing our hash
        Dim hashedString As Byte() = hasher.ComputeHash(bytes) ' Compute the hash and store it as a new Byte array
        Dim stringBuilder As New StringBuilder ' Define a string builder so we can convert back to a string
        For i As Integer = 0 To hashedString.Length - 1
            stringBuilder.Append(hashedString(i).ToString("X2")) ' Append each character of the hash to the string builder
        Next
        Return stringBuilder.ToString ' Return the string builder converted to a string
    End Function

    Public Function TripleDES(ByVal fileToEncrypt As String, ByVal choice As String)
        If choice = "encrypt" Then ' If the user chooses to encrypt...

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' We first get the filename from the directory. I don't believe this line is necessary, however I've done it anyway.
            Dim outputFileName As String = fileName & ".3des" ' We then add the .3des extension to the files name. This will help users identify that the file is encrypted with Triple DES.
            Dim encryptor As New TripleDESCryptoServiceProvider ' Unlike hashing, we need a CryptoServiceProvider to help us encrypt the file given. CryptoServiceProvider can also be used for hashes, however it is not the best way.
            encryptor.Key = ASCIIEncoding.ASCII.GetBytes("LCxS9JAVvVkbRhmLvga5nQQB") ' This defines the key that will be used to encrypt the file. This string MUST be 24 characters long, otherwise the program will produce an error.
            encryptor.IV = ASCIIEncoding.ASCII.GetBytes("jbeQ7FLM") ' This defines the Initialization Vector that will be used to encrypt the file. This string MUST be 8 characters long, otherwise the program will produce an error.
            Dim cryptotransform As ICryptoTransform = encryptor.CreateEncryptor(encryptor.Key, encryptor.IV) ' This line of code does exactly what you think it would do. It will create the encryptor that will be used to transform the plain file into a cipher file.
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' Of course, we need the file that we are going to encrypt. We will define this file as a FileStream.
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' We also need to create the cipher file, which we will also define as a FileStream.
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' Finally, this CryptoStream will tell the program where to write the file and with what CryptoTransform. We are writing to the cipher file, so pur Stream Mode will be write.
            Dim inputArray(sourceFile.Length - 1) As Byte ' This defines a Byte() array that will help to read all of the data from our source file and then write it to the cipher file.
            sourceFile.Read(inputArray, 0, inputArray.Length) ' We then read the file to get all of the data to encrypt it in the next line.
            encrprocess.Write(inputArray, 0, inputArray.Length) ' We then write all of the encrypted data to the cipher file.
            encrprocess.Close()
            sourceFile.Close()  ' These three lines are here to make sure that all the streams are closed, so they cannot produce errors later on.
            outputFile.Close()
            Dim returnString As String = "File Encrypted!"
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' This will delete the file we started with for security reasons.
            Return returnString ' We then return "File Encrypted" back to the program that called it.

        ElseIf choice = "decrypt" Then ' If the user chooses to decrypt...

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' We first get the filename from the directory. I don't believe this line is necessary, however I've done it anyway.
            Dim outputFileName As String = Replace(fileName, ".3des", "") ' We will also remove the .3des extension to indicate that the file is not encryted.
            Dim decryptor As New TripleDESCryptoServiceProvider ' The CryptoServiceProvider is also used for decryption, so we will define it in a similar way we did with the encryption process.
            decryptor.Key = ASCIIEncoding.ASCII.GetBytes("LCxS9JAVvVkbRhmLvga5nQQB") ' For testing purposes, we have given the decryptor the same Key and Initialization Vector as we did when encrypting our file.
            decryptor.IV = ASCIIEncoding.ASCII.GetBytes("jbeQ7FLM")                  ' Remember: The Key MUST be 24 characters long and the Initialization Vector MUST be 8 characters long, otherwise the program will produce an error.
            Dim cryptotransform As ICryptoTransform = decryptor.CreateDecryptor(decryptor.Key, decryptor.IV) ' Like in the encryption process, we create the decryptor via a CryptoTransform variable. We will be using this to transform the cipher file into a plain file later in the function.
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' We will also need to define the cipher file as our source file to decrypt.
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' We will also need to create the plain file as the file to feed the transformed data into.
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' The encryption process is handled via the CryptoStream, only with slightly different variables than the encryption process.
            Dim inputArray(sourceFile.Length - 1) As Byte ' This Byte variable is defined so that the function knows how much data is in the source file.
            sourceFile.Read(inputArray, 0, inputArray.Length) ' We then read the cipher file so that the program has the cipher data to decrypt.
            encrprocess.Write(inputArray, 0, inputArray.Length) ' After that, we write all of the transformed data into our plain file.
            encrprocess.Close()
            sourceFile.Close() ' These three lines are here to make sure that all of the streams are closed, so that they cannot produce errors later on.
            outputFile.Close()
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' Finally, we delete the cipher file for security reasons.
            Dim returnString As String = "File Decrypted!"
            Return returnString ' We then return "File Decrypted" back to the program that called it.

        Else ' If the user puts neither "encrypt" nor "decrypt"

            Dim returnString As String = "Invalid Options."
            Return returnString ' We return "Invalid Options" back to the program that called it.

        End If

    End Function

    Public Function DES(ByVal fileToEncrypt As String, ByVal choice As String)
        If choice = "encrypt" Then

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' Define the file name
            Dim outputFileName = fileName & ".des" ' Add the .des extension to the file name, indicating that the file is encrypted with DES
            Dim encryptor As New DESCryptoServiceProvider ' Define a new DESCryptoServiceProvider
            encryptor.Key = ASCIIEncoding.ASCII.GetBytes("M3k5jOa1") ' Define the encryption key. MUST be 8 characters long.
            encryptor.IV = ASCIIEncoding.ASCII.GetBytes("jbeQ7FLM") ' Define the Initialization Vector. MUST be 8 characters long.
            Dim cryptotransform As ICryptoTransform = encryptor.CreateEncryptor(encryptor.Key, encryptor.IV) ' Define a new ICryptoTransform and create a new encryptor with the Key and Initialization Vector
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' Define the file to encrypt.
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' Define the file to feed the ciphered data to.
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' Define the crytostream that will start the process of encryption.
            Dim inputArray(sourceFile.Length - 1) As Byte ' Define the source files data as a Byte variable
            sourceFile.Read(inputArray, 0, inputArray.Length) ' Read all of the data from the source file
            encrprocess.Write(inputArray, 0, inputArray.Length) ' Write the ciphered data to the cipher file.
            encrprocess.Close()
            sourceFile.Close() ' Close all file streams
            outputFile.Close()
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' Delete the plain file.
            Dim returnString As String = "File Encrypted!"
            Return returnString ' Return "File Encrypted" to whatever called the function.

        ElseIf choice = "decrypt" Then

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' Define the file to decrypt
            Dim outputFileName As String = Replace(fileName, ".des", "") ' Remove the .des file extension
            Dim decryptor As New DESCryptoServiceProvider ' Define a new DESCryptoServiceProvider
            decryptor.Key = ASCIIEncoding.ASCII.GetBytes("M3k5jOa1") ' Define the decryption key
            decryptor.IV = ASCIIEncoding.ASCII.GetBytes("jbeQ7FLM") ' Define the decryption initialization vector
            Dim cryptotransform As ICryptoTransform = decryptor.CreateDecryptor(decryptor.Key, decryptor.IV) ' Define a new ICryptoTransform variable and create a decryptor with the key and initialization vector.
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' Define the file to decrypt
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' Define the plain file to feed the plain data into
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' Define the cryptostream that will start the process of decryption
            Dim inputArray(sourceFile.Length - 1) As Byte ' Define the source files data as a Byte variable
            sourceFile.Read(inputArray, 0, inputArray.Length) ' Read the data from the source file
            encrprocess.Write(inputArray, 0, inputArray.Length) ' Write the plain data to the output file
            encrprocess.Close()
            sourceFile.Close() ' Close all file streams
            outputFile.Close()
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' Delete the encrypted file
            Dim returnString As String = "File Decrypted!"
            Return returnString ' Return "File Decrypted to whatever called the function

        Else

            Dim returnString As String = "Invalid Options."
            Return returnString ' Return "Invalid Options" to whatever called the function

        End If
    End Function

    Public Function AES(ByVal fileToEncrypt As String, ByVal choice As String)
        If choice = "encrypt" Then

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' Get the file name
            Dim outputFileName As String = fileName & ".aes" ' Add the .aes file extension to indicate that the file is encrypted with AES.
            Dim encryptor As New AesCryptoServiceProvider ' Define the encryptor as an AesCryptoServiceProvider
            encryptor.Key = ASCIIEncoding.ASCII.GetBytes("dRgUkXp2s5v8y/B?E(G+KbPeShVmYq3t") ' Define the key and convert it to a Byte() variable
            encryptor.IV = ASCIIEncoding.ASCII.GetBytes("D(G+KbPeSgVkYp3s") ' Define the initialization vector and convert it to a Byte() variable
            Dim cryptotransform As ICryptoTransform = encryptor.CreateEncryptor(encryptor.Key, encryptor.IV) ' Create the encryptor with an ICryptoTransform variable
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' Define the source file to encrypt
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' Create the file that will hold the encrypted data
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' Define the encryption process that will write the encrypted data
            Dim inputArray(sourceFile.Length - 1) As Byte ' Define the source file length as a Byte variable.
            sourceFile.Read(inputArray, 0, inputArray.Length) ' Read the source files data
            encrprocess.Write(inputArray, 0, inputArray.Length) ' Encrypt the data and write it to the cipher file.
            encrprocess.Close()
            sourceFile.Close() ' Close all file streams
            outputFile.Close()
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' Delete the plain file.
            Dim returnString As String = "File Encrypted!"
            Return returnString ' Return "File Encrypted" to whatever called the function.

        ElseIf choice = "decrypt" Then

            Dim fileName As String = Path.GetFileName(fileToEncrypt) ' Get the file name
            Dim outputFileName As String = Replace(fileName, ".aes", "") ' Remove the .aes file to indicated that it is a plain file.
            Dim decryptor As New AesCryptoServiceProvider ' Define the decryptor as an AesCryptoServiceProvider
            decryptor.Key = ASCIIEncoding.ASCII.GetBytes("dRgUkXp2s5v8y/B?E(G+KbPeShVmYq3t") ' Give the same key and initialization vector that was used to encrypt.
            decryptor.IV = ASCIIEncoding.ASCII.GetBytes("D(G+KbPeSgVkYp3s") '                   ^^^^^^^^^^^^^^
            Dim cryptotransform As ICryptoTransform = decryptor.CreateDecryptor(decryptor.Key, decryptor.IV) ' Create the decryptor with an ICryptoTransform variable
            Dim sourceFile As FileStream = New FileStream(fileToEncrypt, FileMode.Open, FileAccess.Read) ' Define the source file to decrypt
            Dim outputFile As FileStream = New FileStream(outputFileName, FileMode.Create, FileAccess.Write) ' Create the plain file
            Dim encrprocess As CryptoStream = New CryptoStream(outputFile, cryptotransform, CryptoStreamMode.Write) ' Define the encryption process that will write the plain data
            Dim inputArray(sourceFile.Length - 1) As Byte ' Define the source file length as a Byte variable
            sourceFile.Read(inputArray, 0, inputArray.Length) ' Read the source files data
            encrprocess.Write(inputArray, 0, inputArray.Length) ' Encrypt the data and write it to the cipher file
            encrprocess.Close()
            sourceFile.Close() ' Close all file streams
            outputFile.Close()
            My.Computer.FileSystem.DeleteFile(fileToEncrypt) ' Delete the encrypted file
            Dim returnString As String = "File Decrypted!"
            Return returnString ' Return "File Decrypted" to whatever called the function

        Else

            Dim returnString As String = "Invalid Options."
            Return returnString ' Return "Invalid Options"

        End If

    End Function

End Module
