Namespace x32
    Public Class SimpleSpeechCommand
        Inherits ComplexSpeechCommand
        Private _Command As CommandHandler

        Public Sub New(ByVal Name As String, ByVal Command As CommandHandler)
            MyBase.New(Name, Name, Nothing)
            _Command = Command
        End Sub
        Public Overrides Sub executeCommand(ByVal recoResult As SpeechLib.ISpeechPhraseInfo)
            _Command.Invoke()
        End Sub
    End Class

    Public Class ComplexSpeechCommand
        Private _Name As String
        Private _Text As String
        Private Command As ComplexCommandHandler

        Public Sub New(ByVal Name As String, ByVal Text As String, ByVal command As ComplexCommandHandler)
            _Name = Name
            _Text = Text
            Me.Command = command
        End Sub

        Public Overridable Sub executeCommand(ByVal recoResult As SpeechLib.ISpeechPhraseInfo)
            Command.Invoke(recoResult)
        End Sub

        Public ReadOnly Property Name() As String
            Get
                Return _Name
            End Get
        End Property
    End Class

End Namespace

