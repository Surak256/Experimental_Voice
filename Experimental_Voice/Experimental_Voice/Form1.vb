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
        mySpeech.addCommand("maximize window", New x32.CommandHandler(AddressOf Maximise))
        mySpeech.addCommand("restore window", New x32.CommandHandler(AddressOf Restore))
    End Sub



    Public Sub EndProgram()
        Application.Exit()
    End Sub

    Public Sub Maximise()
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Public Sub Restore()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Public Sub DisplayMessage(ByVal phrase As SpeechLib.ISpeechPhraseInfo)
        MsgBox(phrase.Rule.Children.Item(0).Name)
    End Sub

    Private Sub btnNewCommand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewCommand.Click
        Dim name As String = InputBox("Input a command to be spoken.")
        'Dim message As String = InputBox("Input a message to display when this command is executed")
        mySpeech.addCommand(name, New x32.ComplexCommandHandler(AddressOf DisplayMessage), name)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Dim name As String = InputBox("Name of command to be removed.")
        mySpeech.removeCommand(name)
    End Sub
End Class
