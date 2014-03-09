Imports SpeechLib

Namespace x32
    Public Delegate Sub CommandHandler()

    Public Delegate Sub ComplexCommandHandler(ByVal result As ISpeechPhraseInfo)

    Public Class Speech
        'Debug variables
        Private monitor As frmMonitor

        'Functional variables
        Private _commandMap As New Dictionary(Of Integer, ComplexSpeechCommand)

        Private speechEngine As SpInprocRecognizer
        Private listener As SpInProcRecoContext
        Private builder As ISpGrammarBuilder
        Private grammar As ISpeechRecoGrammar

        Private commandListPtr As IntPtr
        Private hStateDigit As IntPtr
        Private hStateGreekLetter

        Private index As Integer = 3

        'Constants
        Public Const SPRULETRANS_DICTATION As Int32 = -3

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
                'Command rule
                builder.GetRule("Command", 0, SpeechRuleAttributes.SRATopLevel, True, commandListPtr)
                builder.AddRuleTransition(hStateMain2, Nothing, commandListPtr, 1, Nothing)
                'Epsilon transition
                'Required to have a full path to Nothing
                builder.AddWordTransition(commandListPtr, Nothing, Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Utility rule for numbers
                builder.GetRule("Digit", 1, SpeechRuleAttributes.SRATopLevel, True, hStateDigit)
                builder.AddWordTransition(hStateDigit, Nothing, "zero", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "one", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "two", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "three", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "four", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "five", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "six", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "seven", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "eight", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateDigit, Nothing, "nine", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Utility rule for Greek Letters
                builder.GetRule("Greek Letter", 2, SpeechRuleAttributes.SRATopLevel, True, hStateGreekLetter)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "alpha", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "beta", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "gamma", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "delta", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "epsilon", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "zeta", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "eta", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "theta", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "iota", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "kappa", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "lambda", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "mu", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "nu", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "xi", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "omicron", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "pi", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "rho", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "sigma", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "tau", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "upsilon", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "phi", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "chi", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "psi", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.AddWordTransition(hStateGreekLetter, Nothing, "omega", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Commit and set active
                monitor.writeLine("Committing changes")
                builder.Commit(0)
                grammar.CmdSetRuleState("Main", SpeechRuleState.SGDSActive)
                grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSInactive)
                grammar.CmdSetRuleState("Digit", SpeechRuleState.SGDSInactive)
                grammar.CmdSetRuleState("Greek Letter", SpeechRuleState.SGDSInactive)
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
        Public Sub addCommand(ByVal CommandName As String, ByRef Command As CommandHandler)
            monitor.writeLine("Adding simple command: " & CommandName)
            Try
                Dim myCommand As New SimpleSpeechCommand(commandName, index, command)
                index += 1 'Increment the id for the next command
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

        ''' <summary>
        ''' Adds a command to the grammar in the format "Computer, [command]"
        ''' </summary>
        ''' <param name="CommandName">Command to be spoken</param>
        ''' <param name="Command">Sub to be run when the command is spoken</param>
        ''' <remarks>
        ''' This overload of addCommand supports a complex command handler. Other than that,
        ''' it functions exactly the same as the overload that accepts a standard command handler.
        ''' </remarks>
        Public Sub addCommand(ByVal CommandName As String, ByRef Command As ComplexCommandHandler)
            monitor.writeLine("Adding command with complex handler: " & CommandName)
            Try
                Dim myCommand As New ComplexSpeechCommand(CommandName, index, CommandName, Command)
                index += 1
                _commandMap.Add(myCommand.ID, myCommand)
                Dim hStateNewCommand As IntPtr
                builder.GetRule(CommandName, myCommand.ID, SpeechLib.SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                builder.AddWordTransition(hStateNewCommand, Nothing, CommandName, " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.Commit(0)
            Catch ex As Exception
                monitor.writeLine("Critical error adding command:")
                monitor.writeLine(ex.ToString())
                monitor.writeLine(ex.StackTrace)
            End Try
        End Sub

        ''' <summary>
        ''' Adds a complex command to the grammer.
        ''' </summary>
        ''' <param name="Name">Name of the command</param>
        ''' <param name="Command">Sub to run when the command is spoken</param>
        ''' <param name="Text">Text of the command</param>
        ''' <remarks>
        ''' <para>Complex commands recieve full phrase information when recognized. This information can
        ''' be ignored by the referenced sub, but will be helpful in determining exactly what
        ''' action should be taken.</para>
        ''' 
        ''' <para>The text of the command is interpreted as follows:</para>
        ''' 
        ''' <para>Unmarked text (Not enclosed in brackets of any kind) will be inserted directly in the 
        ''' rule. Brackets are used to mark nodes that require special processing. Ex: "Hello world!"
        ''' </para>
        ''' 
        ''' <para>A bracket must have an initial character inside it that identifies the form of 
        ''' special processing it will recieve. Depending on the element's type, other elements 
        ''' may be nested inside it. The enclosed elements are read immediately after the identifying
        ''' character. A space should follow the identifying character.</para>
        ''' 
        ''' <para>A "L" character indicates the brackets contain a list of pipe-delimited elements
        ''' that should be treated as interchangeable. This is equivalent to the &lt;L&gt; element
        ''' in the text XML grammar format. Ex: [L red|green|blue]</para>
        ''' 
        ''' <para>A question mark ("?") indicates that the brackets contain an element that is optional.
        ''' An epsilon transition will be inserted along with the enclosed element to accomplish
        ''' this. Ex: [? please]</para>
        ''' 
        ''' <para>An asterisk ("*") indicates that dictation will be used. The brackets must contain two
        ''' numbers, the upper and lower bounds for number of words recognized through dictation.
        ''' Ex: [* 1 4]</para>
        ''' 
        ''' <para>A pound sign ("#") indicates that the rule references the digit rule. The brackets must
        ''' contain two numbers, the upper and lower bounds for digits recognized.
        ''' Ex: [# 1 3]</para>
        ''' 
        ''' <para>A "G" character indicates that the rule references the Greek letters rule. The brackets
        ''' must contain two numbers, the upper and lower bounds for the number of letters recognized.
        ''' Ex: [G 0 5]</para>
        ''' 
        ''' <para>An "A" character indicates that the rule references the phonetic alphabet rule. The 
        ''' brackets must contain two numbers, the upper and lower bounds for the number of letters 
        ''' recognized. Ex: [A 0 9]</para>
        ''' 
        ''' <para>These elements can be combined to create highly adaptive rules. The List and Optional
        ''' elements can contain within them other elements of any type. For example the rule:
        ''' "Run program [* 1 5] [? in [L normal|maximized|minimized|fullscreen] format]" requires
        ''' the speaker to specify a program between one and five words long to run, but optionally
        ''' allows options for how to display that program's window. The handler for this rule would
        ''' need to be similarly adaptive when interpreting the recognized text.</para>
        ''' </remarks>
        Public Sub addCommand(ByVal Name As String, ByVal Command As ComplexCommandHandler, ByVal Text As String)
            monitor.writeLine("Adding complex command: " & Name)
            monitor.writeLine("    Command text: " & Text)
            Try
                Dim myCommand As New ComplexSpeechCommand(Name, index, Text, Command)
                index += 1
                _commandMap.Add(myCommand.ID, myCommand)
                Dim hStateNewCommand As IntPtr
                builder.GetRule(Name, myCommand.ID, SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                ParseCommandText(Text, hStateNewCommand, Nothing)
                builder.Commit(0)
            Catch ex As Exception
                monitor.writeLine("Critical error adding command:")
                monitor.writeLine(ex.ToString())
                monitor.writeLine(ex.StackTrace)
            End Try
        End Sub

        Private Sub ParseCommandText(ByVal Text As String, ByRef hStateBefore As IntPtr, ByRef hStateAfter As IntPtr)
            monitor.writeLine("Parsing command: " & Text)
            Dim nodes As String() = SeparateIntoNodes(Text)
            Dim hStates(nodes.Length) As IntPtr
            hStates(0) = hStateBefore
            hStates(hStates.Length - 1) = hStateAfter
            For i As Integer = 1 To hStates.Length - 2
                builder.CreateNewState(hStateBefore, hStates(i))
            Next
            For i As Integer = 0 To nodes.Length - 1
                monitor.writeLine("Node: " & nodes(i))
                monitor.Write("Node type: ")
                If nodes(i).Substring(0, 1) = "[" Then
                    Select Case nodes(i).Substring(1, 1)
                        Case "L"
                            monitor.writeLine("List")
                            Dim items As String() = nodes(i).Substring(3, nodes(i).Length - 4).Split("|")
                            For Each item As String In items
                                monitor.writeLine("    Adding item: " & item)
                                ParseCommandText(item, hStates(i), hStates(i + 1))
                            Next
                        Case "?"
                            monitor.writeLine("Optional")
                            builder.AddWordTransition(hStates(i), hStates(i + 1), Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                            ParseCommandText(nodes(i).Substring(3, nodes(i).Length - 4), hStates(i), hStates(i + 1))
                        Case "*"
                            monitor.writeLine("Dictation")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), New IntPtr(SPRULETRANS_DICTATION))
                        Case "#"
                            monitor.writeLine("Number")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), hStateDigit)
                        Case "G"
                            monitor.writeLine("Greek Alphabet")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), hStateGreekLetter)
                        Case "A"
                            monitor.writeLine("Phonetic alphabet - Not yet implemented")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            builder.AddWordTransition(hStates(i), hStates(i + 1), Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                            'Add phonetic alphabet rule reference
                        Case Else
                            monitor.writeLine("Unknown node type")
                    End Select
                Else
                    monitor.writeLine("Plaintext node")
                    builder.AddWordTransition(hStates(i), hStates(i + 1), nodes(i), " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                End If
            Next
        End Sub

        Public Function SeparateIntoNodes(ByVal Text As String) As String()
            If Text.Contains("[") Then
                Dim nodeList As New List(Of String)
                Dim Start As Integer = 0
                Dim nodeStack As New Stack(Of Integer)
                Dim limit As Integer = 0
                While ((Text.IndexOf("[", Start) <> -1 Or Text.IndexOf("]", Start) <> -1) And limit < 100)
                    Dim open As Integer = Text.IndexOf("[", Start)
                    Dim close As Integer = Text.IndexOf("]", Start)
                    If open < close And open <> -1 Then
                        'Check for a plaintext node in a top-level node
                        If open > Start And nodeStack.Count = 0 Then
                            Dim tryText = Text.Substring(Start, open - Start)
                            If tryText.Trim().Length > 0 Then
                                nodeList.Add(tryText.Trim())
                            End If
                        End If
                        nodeStack.Push(open)
                        Start = open + 1
                    Else
                        If nodeStack.Count > 1 Then
                            nodeStack.Pop()
                        Else
                            Dim nodestart = nodeStack.Pop()
                            nodeList.Add(Text.Substring(nodestart, close - nodestart + 1))
                        End If
                        Start = close + 1
                    End If
                    limit += 1
                End While
                If limit = 100 Then
                    monitor.writeLine("Limit exceeded. Error while parsing")
                End If
                Return nodeList.ToArray()
            Else 'Simple phrase
                monitor.writeLine("Simple phrase")
                Dim ret As String() = {Text}
                Return ret
            End If
        End Function

        Private Sub addMultipleTransition(ByVal min As Integer, ByVal max As Integer, ByRef hStateStart As IntPtr, ByRef hStateEnd As IntPtr, ByVal hStateTransition As IntPtr)
            Dim hStates(max) As IntPtr
            hStates(0) = hStateStart
            hStates(max) = hStateEnd
            For i As Integer = 1 To max - 1
                builder.CreateNewState(hStateStart, hStates(i))
            Next
            For i As Integer = 0 To max - 1
                builder.AddRuleTransition(hStates(i), hStates(i + 1), hStateTransition, 1, Nothing)
                If i >= min Then
                    builder.AddWordTransition(hStates(i), hStates(i + 1), Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                End If
            Next
        End Sub

        Public Sub removeCommand(ByVal Name As String)
            If Name.ToLower() = "main" Then
                monitor.writeLine("Cannot remove root command")
            Else
                monitor.writeLine("Removing command: " & Name)
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
            If Result.PhraseInfo.Rule.Name = "Main" Then
                If Result.PhraseInfo.Rule.Children.Item(0).Children Is Nothing Then
                    monitor.writeLine("Epsilon transition used" & vbNewLine & _
                                      "You do not need to pause for recogniton")
                Else
                    Dim ruleID As Integer = Result.PhraseInfo.Rule.Children.Item(0).Children.Item(0).Id
                    If _commandMap.ContainsKey(ruleID) Then
                        _commandMap.Item(ruleID).executeCommand(Result.PhraseInfo)
                    End If
                End If
            End If
        End Sub
        Public Sub DisplayGrammar()
            monitor.writeLine("Rule listing:")
            For i As Integer = 0 To grammar.Rules.Count - 1
                monitor.writeLine("    " & grammar.Rules.Item(i).Name)
            Next
        End Sub
    End Class
End Namespace
