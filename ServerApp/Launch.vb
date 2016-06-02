Module Launch
    Public Sub Main(ByVal cmdArgs() As String)

        'Dim regStartUp As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)

        'Dim value As String

        'value = regStartUp.GetValue("BluCRESTSigner")

        'If value <> Application.ExecutablePath.ToString() Then

        '    regStartUp.CreateSubKey("BluCRESTSigner")
        '    regStartUp.SetValue("BluCRESTSigner", Application.ExecutablePath.ToString())

        'End If

        Application.EnableVisualStyles()
        Application.Run(New AppContext)


    End Sub
End Module
