Public Class frmMonitor
    Dim mySpeech As x32.Speech

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mySpeech = New x32.Speech(Me)
    End Sub

    Public Sub WriteIf(ByVal condition As Boolean, ByVal text As String)
        If condition Then
            txtOut.AppendText(text)
            checkError(text)
        End If
    End Sub

    Public Sub writeLineIf(ByVal condition As Boolean, ByVal text As String)
        If condition Then
            txtOut.AppendText(text & vbNewLine)
            checkError(text)
        End If
    End Sub

    Private Sub checkError(ByVal text As String)
        If text.Contains("Critical error") Then
            txtOut.BackColor = Color.Red
        End If
    End Sub

    Public Sub frmMonitor_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        mySpeech.LoadSpeech()
        mySpeech.addCommand("end program", New x32.CommandHandler(AddressOf EndProgram))
        mySpeech.addCommand("maximize window", New x32.CommandHandler(AddressOf Maximise))
        mySpeech.addCommand("restore window", New x32.CommandHandler(AddressOf Restore))
        mySpeech.addCommand("enable continuous commands", New x32.CommandHandler(AddressOf ActivateContinuousCommands))
        mySpeech.addCommand("disable continous commands", New x32.CommandHandler(AddressOf DeactiateContinuousCommands))
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
        MsgBox(phrase.Rule.Children.Item(0).Name & vbNewLine & phrase.GetText())
    End Sub

    Private Sub btnNewCommand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewCommand.Click
        Dim name As String = InputBox("Input the command's name. (Must be unique)")
        Dim text As String = InputBox("Input the text of the command")
        mySpeech.addCommand(name, New x32.ComplexCommandHandler(AddressOf DisplayMessage), text)
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Dim name As String = InputBox("Name of command to be removed.")
        mySpeech.removeCommand(name)
    End Sub

    Private Sub btnDisplay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplay.Click
        mySpeech.DisplayGrammar()
    End Sub

    Private Sub btnSubRule_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubRule.Click
        Dim name As String = InputBox("Input the subrule's name. (Must be unique)")
        Dim text As String = InputBox("Input the text of the subrule.")
        mySpeech.addSubRule(name, text)
    End Sub

    Private Sub btnValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnValidate.Click
        Dim text As String = InputBox("Enter the text of the rule to validate")
        Dim isValid As x32.CommandTextValidationError = mySpeech.IsValidCommand(text)
        If isValid = x32.CommandTextValidationError.NO_ERROR Then
            MsgBox("Valid!")
        Else
            MsgBox("Invalid." & vbNewLine & isValid.ToString())
        End If
    End Sub

    Private Sub ActivateContinuousCommands()
        mySpeech.ContinuousCommands = True
    End Sub

    Private Sub DeactiateContinuousCommands()
        mySpeech.ContinuousCommands = False
    End Sub
End Class
