Imports SpeechLib

Namespace x32
    Public Delegate Sub CommandHandler()

    Public Class Speech
        'Debug variables
        Private monitor As frmMonitor

        'Functional variables
        Private _commandMap As New Dictionary(Of Integer, SpeechCommand)

        Private speechEngine As SpInprocRecognizer
        Private listener As SpInProcRecoContext
        Private builder As ISpGrammarBuilder
        Private grammar As ISpeechRecoGrammar

        Private commandListPtr As IntPtr
        Private i As Integer = 2

        Public Sub New(ByVal monitor As frmMonitor)
            Me.monitor = monitor
        End Sub

        Public Sub LoadSpeech()
            Try
                'Load recognizer
                monitor.writeLine("Loading recognizer")
                speechEngine = New SpInprocRecognizer
                listener = speechEngine.CreateRecoContext()
                monitor.writeLine("Setting audio inputs")
                Dim inputs As ISpeechObjectTokens = listener.Recognizer.GetAudioInputs()
                If (inputs.Count < 1) Then
                    MsgBox("No audio inputs found!")
                End If
                listener.Recognizer.AudioInput = inputs.Item(0)
                AddHandler listener.Recognition, AddressOf onRecognition
                monitor.writeLine("Setting up grammar object")
                grammar = listener.CreateGrammar(1)
                grammar.DictationLoad()
                grammar.DictationSetState(SpeechRuleState.SGDSInactive)
                grammar.State = SpeechGrammarState.SGSEnabled
                builder = grammar ' I have no idea why this works. It does.
                'Load grammar
                monitor.writeLine("Loading grammar")
                'Main rule
                Dim hStateMain As IntPtr
                builder.GetRule("Main", 0, SpeechLib.SpeechRuleAttributes.SRATopLevel, True, hStateMain)
                'Interim state
                Dim hStateMain2 As IntPtr
                builder.CreateNewState(hStateMain, hStateMain2)
                'Add word transition for init command ("Computer")
                builder.AddWordTransition(hStateMain, hStateMain2, "Computer,", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                commandListPtr = hStateMain2
                'Cancel command
                'Required to have a full path to Nothing
                Dim hStateCancel As IntPtr
                builder.GetRule("Cancel", 1, 0, True, hStateCancel)
                builder.AddRuleTransition(hStateMain2, Nothing, hStateCancel, 1, Nothing)
                builder.AddWordTransition(hStateCancel, Nothing, "cancel", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Commit and set active
                monitor.writeLine("Committing changes")
                builder.Commit(0)
                grammar.CmdSetRuleState("Main", SpeechRuleState.SGDSActive)
                monitor.writeLine("Initialization complete. Listening.")
            Catch ex As Exception
                monitor.writeLine("Critical error:")
                monitor.writeLine(ex.ToString())
                monitor.writeLine(ex.StackTrace())
            End Try
        End Sub

        ''' <summary>
        ''' Adds a command to the grammer in the format "Computer, [command]"
        ''' </summary>
        ''' <param name="commandName">Command to be spoken</param>
        ''' <param name="command">Sub to be run when the command is spoken</param>
        ''' <remarks></remarks>
        Public Sub addCommand(ByVal commandName As String, ByRef command As CommandHandler)
            monitor.writeLine("Adding command: " & commandName)
            Try
                Dim myCommand As SpeechCommand
                myCommand.Name = commandName
                myCommand.ID = i
                i += 1 'Increment the id for the next command
                myCommand.Command = command
                _commandMap.Add(myCommand.ID, myCommand)
                Dim hStateNewCommand As IntPtr
                builder.GetRule(commandName, myCommand.ID, SpeechLib.SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                builder.AddWordTransition(hStateNewCommand, Nothing, commandName, " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.Commit(0)
            Catch ex As Exception
                monitor.writeLine("Critical error adding command:")
                monitor.writeLine(ex.ToString())
                monitor.writeLine(ex.StackTrace)
            End Try
        End Sub


        Public Sub removeCommand(ByVal Name As String)
            If Name = "Main" Then
                monitor.writeLine("Cannot remove root command")
            Else
                Dim hStateRemoved As IntPtr
                'Creates rule if it doesn't exist, otherwise returns current rule
                builder.GetRule(Name, 0, 0, True, hStateRemoved)
                builder.ClearRule(hStateRemoved)
                builder.Commit(0)
            End If
        End Sub
        Private Sub onRecognition(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
            monitor.writeLine("Recognized: " & Result.PhraseInfo.GetText())
            monitor.writeLine("    Rule name: " & Result.PhraseInfo.Rule.Name)
            monitor.writeLine("    Rule ID: " & Result.PhraseInfo.Rule.Id)
            If Result.PhraseInfo.Rule.Children.Count > 0 Then
                For x As Integer = 0 To Result.PhraseInfo.Rule.Children.Count - 1
                    monitor.writeLine(" Child rule: " & Result.PhraseInfo.Rule.Children.Item(x).Name)
                Next
            End If
            If _commandMap.ContainsKey(Result.PhraseInfo.Rule.Children.Item(0).Id) Then
                _commandMap.Item(Result.PhraseInfo.Rule.Children.Item(0).Id).Command.Invoke()
            End If
        End Sub
    End Class
End Namespace
