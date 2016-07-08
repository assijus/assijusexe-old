Imports System
Imports System.Security.Cryptography
Imports System.Security.Permissions
Imports System.IO
Imports System.Security.Cryptography.X509Certificates

Imports System.Text.RegularExpressions
Imports System.Security.Cryptography.Pkcs
Imports System.Runtime.InteropServices

Module BluC
    Private getCertificateTitle As String = "Assinatura Digital"
    Private getCertificateMessage As String = "Escolha o certificado que será utilizado na assinatura."

    Private WithEvents ni As New NotifyIcon
    Private Const WM_HOTKEY As Integer = &H312
    Private Const MOD_CONTROL As UInteger = &H2
    Private Const WM_SYSCOMMAND As Integer = &H112
    Private Const SC_MINIMIZE As Integer = &HF020
    Private Const SW_RESTORE As Integer = &H9

    <DllImport("user32.dll", EntryPoint:="RegisterHotKey")> _
    Private Function RegisterHotKey(ByVal hWnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As UInteger, ByVal vk As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="UnregisterHotKey")> _
    Private Function UnregisterHotKey(ByVal hWnd As IntPtr, ByVal id As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="SetForegroundWindow")> _
    Private Function SetForegroundWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindowW")> _
    Private Function FindWindowW(<MarshalAs(UnmanagedType.LPTStr)> ByVal lpClassName As String, <MarshalAs(UnmanagedType.LPTStr)> ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="IsWindowVisible")> _
    Private Function IsWindowVisible(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="ShowWindow")> _
    Private Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="IsIconic")> _
    Private Function IsIconic(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function


    Private Sub activateCertificateSelectionWindow()
        Threading.Thread.Sleep(500) ' 500 milliseconds = 0.5 seconds

        Dim hwnd As IntPtr = FindWindowW(Nothing, getCertificateTitle) 'Find the window handle (Works even if the app is hidden and not shown in taskbar)
        If hwnd <> IntPtr.Zero Then
            If Not IsWindowVisible(hwnd) Or IsIconic(hwnd) Then 'If the window is minimized or hidden then Restore and Show the window
                ShowWindow(hwnd, SW_RESTORE)
                ni.Visible = False
            End If
            SetForegroundWindow(hwnd) 'Set the window as the foreground window
        End If
    End Sub

    Private certificate As X509Certificate2

    Public Sub clearCurrentCertificate()
        certificate = Nothing
    End Sub

    Private Function getCertificateList(subjectRegex As String, issuerRegex As String) As X509Certificate2Collection
        Dim store As X509Store = New X509Store(StoreName.My, StoreLocation.CurrentUser)
        store.Open(OpenFlags.OpenExistingOnly)
        Dim certificates As X509Certificate2Collection = store.Certificates
        Dim certificatesFiltered As X509Certificate2Collection = New X509Certificate2Collection()
        Dim enumCert As X509Certificate2Enumerator = certificates.GetEnumerator()
        While (enumCert.MoveNext())
            Dim certificateTmp As X509Certificate2 = enumCert.Current

            Dim subjectOk As Boolean = True
            If subjectRegex.Length > 0 Then
                subjectOk = certificateTmp.Subject = subjectRegex
                If (Not subjectOk) Then
                    Dim matchSubject As Match = Regex.Match(certificateTmp.Subject, subjectRegex, RegexOptions.IgnoreCase)
                    subjectOk = matchSubject.Success
                End If
            End If

            Dim issuerOk As Boolean = True
            If issuerRegex.Length > 0 Then
                Dim matchIssuer As Match = Regex.Match(certificateTmp.Issuer, issuerRegex, RegexOptions.IgnoreCase)
                issuerOk = matchIssuer.Success
            End If

            Dim dateOk As Boolean = Now > certificateTmp.NotBefore AndAlso Now < certificateTmp.NotAfter

            If subjectOk AndAlso issuerOk AndAlso dateOk AndAlso certificateTmp.HasPrivateKey Then
                certificatesFiltered.Add(certificateTmp)
            End If
        End While
        Return certificatesFiltered
    End Function

    Public Function getCertificate(subjectRegex As String, issuerRegex As String) As String
        Dim ret As String = ""
        Dim certificatesFiltered As X509Certificate2Collection = getCertificateList(subjectRegex, issuerRegex)

        If certificatesFiltered.Count = 0 Then
            certificate = Nothing
            Return ""
        ElseIf certificatesFiltered.Count = 1 Then
            certificate = certificatesFiltered(0)
        Else
            Dim ti As Threading.Thread = New Threading.Thread(AddressOf activateCertificateSelectionWindow)
            ti.Start()
            Dim certificateSel As X509Certificate2Collection = X509Certificate2UI.SelectFromCollection(certificatesFiltered, getCertificateTitle, getCertificateMessage, X509SelectionFlag.SingleSelection)
            If certificateSel.Count > 0 Then
                certificate = certificateSel(0)
            End If
        End If
        Dim certAsByte As Byte() = certificate.Export(X509ContentType.Cert)
        Dim certAsString As String = Convert.ToBase64String(certAsByte)
        ret = certAsString
        Return ret
    End Function

    Public Function getCertificateBySubject(subject As String) As String
        Dim certificatesFiltered As X509Certificate2Collection = getCertificateList(subject, "")
        If certificatesFiltered.Count > 0 Then
            certificate = certificatesFiltered(0)
            Dim certAsByte As Byte() = certificate.Export(X509ContentType.Cert)
            Dim certAsString As String = Convert.ToBase64String(certAsByte)
            Return certAsString
        End If
        Return Nothing
    End Function

    Public Function getSubject() As String
        If certificate Is Nothing Then
            Return ""
        End If
        Return certificate.Subject
    End Function

    Public Function getKeySize() As Integer
        Dim publicKey As RSACryptoServiceProvider = DirectCast(certificate.PublicKey.Key, RSACryptoServiceProvider)
        Return publicKey.KeySize
    End Function

    Public Function sign(hashAlg As String, contentB64 As String) As String
        Return sign(convertHashAlg(hashAlg), contentB64)
    End Function

    Public Function sign(hashAlg As Integer, contentB64 As String) As String
        Dim content As Byte() = Convert.FromBase64String(contentB64)
        Return Convert.ToBase64String(sign(hashAlg, content))
    End Function

    Public Function SignMsg(hashAlg As Integer, msg As Byte(), signerCert As X509Certificate2) As Byte()
        Dim contentInfo As ContentInfo = New ContentInfo(msg)
        Dim signedCms As SignedCms = New SignedCms(contentInfo, True)
        Dim cmsSigner As CmsSigner = New CmsSigner(signerCert)
        signedCms.ComputeSignature(cmsSigner, False)
        Dim ab As Byte() = signedCms.Encode()

        ' Dim signedCms2 As SignedCms = New SignedCms(contentInfo)
        ' signedCms2.Decode(ab)
        ' Console.WriteLine(signedCms2.Detached)
        ' Dim ab2 As Byte() = signedCms2.Encode()
        ' Dim dettachedB64 As String = Convert.ToBase64String(ab)
        ' Dim attachedB64 As String = Convert.ToBase64String(ab2)
        ' Dim msgB64 As String = Convert.ToBase64String(msg)

        Return ab
    End Function

    Public Function sign(hashAlg As Integer, content As Byte()) As Byte()
        Dim hash As HashAlgorithm = Nothing
        Dim signature As Byte()

        If hashAlg = 99 Then
            signature = SignMsg(hashAlg, content, certificate)
        Else
            Select Case hashAlg
                Case 0
                    hash = New SHA1Managed()
                Case 1
                    Throw New Exception("Algoritmo nao suportado.")
                Case 2
                    hash = New SHA256Managed()
                Case 3
                    hash = New SHA384Managed()
                Case 4
                    hash = New SHA512Managed()
            End Select

            Dim privateKey As RSACryptoServiceProvider = DirectCast(certificate.PrivateKey, RSACryptoServiceProvider)
            Dim publicKey As RSACryptoServiceProvider = DirectCast(certificate.PublicKey.Key, RSACryptoServiceProvider)

            Dim verify As Boolean = False
            signature = privateKey.SignData(content, hash)
            verify = publicKey.VerifyData(content, hash, signature)
            If Not verify Then
                Throw New Exception("Erro verificando assinatura com a chave publica")
            End If
        End If
        Return signature
    End Function

    Public Function convertHashAlg(hashAlg As String) As Integer
        Dim tmp As String = hashAlg.ToUpper()
        Select Case tmp
            Case "SHA1", "0"
                Return 0
            Case "SHA256", "2"
                Return 2
            Case "SHA384", "3"
                Return 3
            Case "SHA512", "4"
                Return 4
            Case "PKCS7", "99"
                Return 99
        End Select
        Throw New Exception("Algoritmo de hash nao reconhecido: " + hashAlg)
    End Function
End Module
