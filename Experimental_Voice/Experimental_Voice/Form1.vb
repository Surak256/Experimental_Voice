Public Class frmMonitor
    Dim mySpeech As x32.Speech

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mySpeech = New x32.Speech(Me)
    End Sub

    Public Sub Write(ByVal text As String)
        txtOut.AppendText(text)
    End Sub

    Public Sub writeLine(ByVal text As String)
        txtOut.AppendText(text & vbNewLine)
    End Sub

    Public Sub frmMonitor_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        mySpeech.LoadSpeech()
        mySpeech.addCommand("end program", New x32.CommandHandler(AddressOf EndProgram))
    End Sub

    Public Sub EndProgram()
        Application.Exit()
    End Sub

End Class
