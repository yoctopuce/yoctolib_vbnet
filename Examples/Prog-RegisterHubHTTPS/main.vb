Imports System.Security.Policy
Imports ConsoleApplication1.yocto_api
Imports System.Threading
Imports System.IO

Module Module1
  Private Function load_cert_from_file(url As String) As String
    Dim Path As String = url.Replace("/", "_").Replace(":", "_") + ".crt"
    If (File.Exists(Path)) Then
      Return File.ReadAllText(Path)
    End If
    Return ""
  End Function

  Private Sub save_cert_to_file(url As String, trustedCert As String)
    Dim Path As String = url.Replace("/", "_").Replace(":", "_") + ".crt"
    File.WriteAllText(Path, trustedCert)
  End Sub


  Sub Main()

    Dim m As YModule
    Dim errmsg As String = ""
    Dim username As String = "admin"
    Dim password As String = "1234"
    Dim host As String = "localhost"
    Dim url As String = "secure://" + username + ":" + password + "@" + host

    ' load known TLS certificate into the API
    Dim trusted_cert As String = load_cert_from_file(host)
    If (trusted_cert <> "") Then
      Dim err As String = YAPI.AddTrustedCertificates(trusted_cert)
      If (err <> "") Then
        Console.WriteLine(err)
        Environment.Exit(0)
      End If
    End If
    ' test connection with VirtualHub
    Dim Res As Integer = YAPI.TestHub(Url, 1000, errmsg)
    If (Res = YAPI.SSL_UNK_CERT) Then
      ' remote TLS certificate Is unknown ask user what to do
      Console.WriteLine("Remote SSL/TLS certificate is unknown")
      Console.WriteLine("You can...")
      Console.WriteLine(" -(A)dd certificate to the API")
      Console.WriteLine(" -(I)gnore this error and continue")
      Console.WriteLine(" -(E)xit")
      Console.Write("Your choice: ")
      Dim line As String = Console.ReadLine().ToLower()
      If (line.StartsWith("a")) Then
        ' download remote certificate And save it locally
        trusted_cert = YAPI.DownloadHostCertificate(Url, 5000)
        If (trusted_cert.StartsWith("error")) Then
          Console.WriteLine(trusted_cert)
          Environment.Exit(0)
        End If
        save_cert_to_file(host, trusted_cert)
        Dim err As String = YAPI.AddTrustedCertificates(trusted_cert)
        If (err <> "") Then
          Console.WriteLine(err)
          Environment.Exit(0)
        End If
      ElseIf (line.StartsWith("i")) Then
        YAPI.SetNetworkSecurityOptions(YAPI.NO_HOSTNAME_CHECK Or YAPI.NO_TRUSTED_CA_CHECK Or
                                           YAPI.NO_EXPIRATION_CHECK)
      Else
        Environment.Exit(0)
      End If
    ElseIf (Res <> YAPI.SUCCESS) Then
      Console.WriteLine("YAPI.TestHub failed:" + errmsg)
      Environment.Exit(0)
    End If


    If (YAPI.RegisterHub(Url, errmsg) <> YAPI.SUCCESS) Then
      Console.WriteLine("YAPI.RegisterHub failed:" + errmsg)
      Environment.Exit(0)
    End If

    Console.WriteLine("Device list")
      m = YModule.FirstModule()
    While m IsNot Nothing
      Console.WriteLine(m.get_serialNumber() + " (" + m.get_productName() + ")")
        m = m.nextModule()
    End While
    YAPI.FreeAPI()

  End Sub

End Module
