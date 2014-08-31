Imports SpeechLib

'TODO:
'   Remove debug monitor
'   Add checks for circular rule references
'   Remove simple commands?

Namespace x32
    Public Delegate Sub CommandHandler()

    Public Delegate Sub ComplexCommandHandler(ByVal result As ISpeechPhraseInfo)

    ''' <summary>
    ''' Contains methods to create, modify, and use a dynamic grammar for interpreting speech.
    ''' </summary>
    Public Class Speech
        'Debug variables
        Private monitor As frmMonitor

        'Functional variables
        Private _commandMap As New Dictionary(Of String, ComplexSpeechCommand)
        Private _internalCommands As New List(Of String)

        Private speechEngine As SpInprocRecognizer
        Private listener As SpInProcRecoContext
        Private builder As ISpGrammarBuilder
        Private grammar As ISpeechRecoGrammar

        Private commandListPtr As IntPtr
        Private hStateDigit As IntPtr
        Private hStateGreekLetter As IntPtr

        'State
        Private _continuousCommands As Boolean = False

        'Logging state
        Private logRecognitions As Boolean = True
        Private logAdditions As Boolean = True
        Private logRemovals As Boolean = True
        Private logNodeSplitting As Boolean = False
        Private logParsing As Boolean = False
        Private logGeneral As Boolean = True
        Private logInitialization As Boolean = False
        Private logErrors As Boolean = True
        Private logRecognitionsDetail As Boolean = False

        'Constants
        ''' <summary>
        ''' Constant for a dictation transition.
        ''' </summary>
        Public Const SPRULETRANS_DICTATION As Int32 = -3

        ''' <summary>
        ''' Creates a new Speech object
        ''' </summary>
        ''' <param name="monitor">Debug monitor</param>
        Public Sub New(ByVal monitor As frmMonitor)
            Me.monitor = monitor
        End Sub

        ''' <summary>
        ''' Loads the speech recognizer and initializes internal commands.
        ''' </summary>
        Public Sub LoadSpeech()
            Try
                'Load recognizer
                monitor.writeLineIf(logInitialization, "Loading recognizer")
                speechEngine = New SpInprocRecognizer
                listener = speechEngine.CreateRecoContext()
                monitor.writeLineIf(logInitialization, "Setting audio inputs")
                Dim inputs As ISpeechObjectTokens = listener.Recognizer.GetAudioInputs()
                If (inputs.Count < 1) Then
                    MsgBox("No audio inputs found!")
                End If
                listener.Recognizer.AudioInput = inputs.Item(0)
                AddHandler listener.Recognition, AddressOf onRecognition
                monitor.writeLineIf(logInitialization, "Setting up grammar object")
                grammar = listener.CreateGrammar(1)
                grammar.DictationLoad()
                grammar.DictationSetState(SpeechRuleState.SGDSInactive)
                grammar.State = SpeechGrammarState.SGSEnabled
                builder = grammar ' I have no idea why this works. It does.
                'Load grammar
                monitor.writeLineIf(logInitialization, "Loading grammar")
                'Main rule
                Dim hStateMain As IntPtr
                builder.GetRule("Main", 0, SpeechLib.SpeechRuleAttributes.SRATopLevel, True, hStateMain)
                _internalCommands.Add("Main")
                'Interim state
                Dim hStateMain2 As IntPtr
                builder.CreateNewState(hStateMain, hStateMain2)
                'Add word transition for init command ("Computer")
                builder.AddWordTransition(hStateMain, hStateMain2, "Computer,", " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Command rule
                builder.GetRule("Command", 0, SpeechRuleAttributes.SRATopLevel, True, commandListPtr)
                _internalCommands.Add("Command")
                builder.AddRuleTransition(hStateMain2, Nothing, commandListPtr, 1, Nothing)
                'Epsilon transition
                'Required to have a full path to Nothing
                builder.AddWordTransition(commandListPtr, Nothing, Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                'Utility rule for numbers
                addSubRule("Digit", "[L zero|one|two|three|four|five|six|seven|eight|nine]")
                _internalCommands.Add("Digit")
                builder.GetRule("Digit", 0, SpeechRuleAttributes.SRATopLevel, False, hStateDigit)
                'Utility rule for Greek Letters
                addSubRule("Greek Letter", "[L alpha|beta|gamma|delta|epsilon|zeta|eta|theta|iota|kappa|lambda|mu|nu|xi|omicron|pi|rho|sigma|tau|upsilon|phi|chi|psi|omega]")
                _internalCommands.Add("Greek Letter")
                builder.GetRule("Greek Letter", 0, SpeechRuleAttributes.SRATopLevel, False, hStateGreekLetter)
                'Commit and set active
                monitor.writeLineIf(logInitialization, "Committing changes")
                builder.Commit(0)
                grammar.CmdSetRuleState("Main", SpeechRuleState.SGDSActive)
                grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSInactive)
                monitor.writeLineIf(logGeneral, "Initialization complete. Listening.")
            Catch ex As Exception
                monitor.writeLineIf(logErrors, "Critical error:")
                monitor.writeLineIf(logErrors, ex.ToString())
                monitor.writeLineIf(logErrors, ex.StackTrace())
            End Try
        End Sub

        ''' <summary>
        ''' Adds a command to the grammer in the format "Computer, [command]"
        ''' </summary>
        ''' <param name="commandName">Command to be spoken</param>
        ''' <param name="command">Sub to be run when the command is spoken</param>
        ''' <remarks></remarks>
        Public Sub addCommand(ByVal CommandName As String, ByRef Command As CommandHandler)
            monitor.writeLineIf(logAdditions, "Adding simple command: " & CommandName)
            Try
                Dim myCommand As New SimpleSpeechCommand(CommandName, Command)
                If _commandMap.ContainsKey(CommandName) Then
                    _commandMap.Remove(CommandName)
                End If
                _commandMap.Add(myCommand.Name, myCommand)
                Dim hStateNewCommand As IntPtr
                builder.GetRule(CommandName, 0, SpeechLib.SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                builder.ClearRule(hStateNewCommand)
                builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                builder.AddWordTransition(hStateNewCommand, Nothing, CommandName, " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.Commit(0)
            Catch ex As Exception
                monitor.writeLineIf(logErrors, "Critical error adding command:")
                monitor.writeLineIf(logErrors, ex.ToString())
                monitor.writeLineIf(logErrors, ex.StackTrace)
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
            monitor.writeLineIf(logAdditions, "Adding command with complex handler: " & CommandName)
            Try
                Dim myCommand As New ComplexSpeechCommand(CommandName, CommandName, Command)
                If _commandMap.ContainsKey(CommandName) Then
                    _commandMap.Remove(CommandName)
                End If
                _commandMap.Add(myCommand.Name, myCommand)
                Dim hStateNewCommand As IntPtr
                builder.GetRule(CommandName, 0, SpeechLib.SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                builder.ClearRule(hStateNewCommand)
                builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                builder.AddWordTransition(hStateNewCommand, Nothing, CommandName, " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                builder.Commit(0)
            Catch ex As Exception
                monitor.writeLineIf(logErrors, "Critical error adding command:")
                monitor.writeLineIf(logErrors, ex.ToString())
                monitor.writeLineIf(logErrors, ex.StackTrace)
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
        ''' <para>An "R" character indicates that the rule references another rule. The brackets
        ''' contain the case-sensitive rule name of the referenced rule. Any given rule may be referenced,
        ''' however the rule may not directly or indirectly reference itself. This precludes using the
        ''' "Main" rule as a subrule. Ex: [R Digit]</para>
        ''' 
        ''' <para>These elements can be combined to create highly adaptive rules. The List and Optional
        ''' elements can contain within them other elements of any type. For example the rule:
        ''' "Run program [* 1 5] [? in [L normal|maximized|minimized|fullscreen] format]" requires
        ''' the speaker to specify a program between one and five words long to run, but optionally
        ''' allows options for how to display that program's window. The handler for this rule would
        ''' need to be similarly adaptive when interpreting the recognized text.</para>
        ''' </remarks>
        Public Sub addCommand(ByVal Name As String, ByVal Command As ComplexCommandHandler, ByVal Text As String)
            monitor.writeLineIf(logAdditions, "Adding complex command: " & Name)
            monitor.writeLineIf(logParsing, "    Command text: " & Text)
            Dim isValid As CommandTextValidationError = IsValidCommand(Text)
            If isValid = CommandTextValidationError.NO_ERROR Then
                Try
                    Dim myCommand As New ComplexSpeechCommand(Name, Text, Command)
                    If _commandMap.ContainsKey(Name) Then
                        _commandMap.Remove(Name)
                    End If
                    _commandMap.Add(myCommand.Name, myCommand)
                    Dim hStateNewCommand As IntPtr
                    builder.GetRule(Name, 0, SpeechRuleAttributes.SRADynamic, True, hStateNewCommand)
                    builder.ClearRule(hStateNewCommand)
                    builder.AddRuleTransition(commandListPtr, Nothing, hStateNewCommand, 1, Nothing)
                    ParseCommandText(Text, hStateNewCommand, Nothing, myCommand)
                    builder.Commit(0)
                Catch ex As Exception
                    monitor.writeLineIf(logErrors, "Critical error adding command:")
                    monitor.writeLineIf(logErrors, ex.ToString())
                    monitor.writeLineIf(logErrors, ex.StackTrace)
                End Try
            Else
                'Throw an exception
                monitor.writeLineIf(logErrors, "    Invalid command text: " & isValid.ToString())
            End If
        End Sub

        ''' <summary>
        ''' Adds a subrule that can be referenced by other rules but not recognized on its own.
        ''' </summary>
        ''' <param name="Name">Name of the rule (Must be unique)</param>
        ''' <param name="Text">Text of the command</param>
        ''' <remarks>
        ''' Command text is interpreted exactly as in a command.
        ''' </remarks>
        Public Sub addSubRule(ByVal Name As String, ByVal Text As String)
            monitor.writeLineIf(logAdditions, "Adding subrule: " & Name)
            monitor.writeLineIf(logParsing, "    CommandText: " & Text)
            Dim isValid As CommandTextValidationError = IsValidCommand(Text)
            If isValid = CommandTextValidationError.NO_ERROR Then
                Try
                    Dim myRule As New ComplexSpeechCommand(Name, Text, Nothing)
                    If _commandMap.ContainsKey(Name) Then
                        _commandMap.Remove(Name)
                    End If
                    _commandMap.Add(myRule.Name, myRule)
                    Dim hStateNewRule As IntPtr
                    builder.GetRule(Name, 0, SpeechRuleAttributes.SRATopLevel, True, hStateNewRule)
                    builder.ClearRule(hStateNewRule)
                    ParseCommandText(Text, hStateNewRule, Nothing, myRule)
                    builder.Commit(0)
                    grammar.CmdSetRuleState(Name, SpeechRuleState.SGDSInactive)
                Catch ex As Exception
                    monitor.writeLineIf(logErrors, "Critical error adding rule:")
                    monitor.writeLineIf(logErrors, ex.ToString())
                    monitor.writeLineIf(logErrors, ex.StackTrace)
                End Try
            Else
                'Throw an exception
                monitor.writeLineIf(logErrors, "    Invalid command text: " & isValid.ToString())
            End If
        End Sub

        Private Sub ParseCommandText(ByVal Text As String, ByRef hStateBefore As IntPtr, ByRef hStateAfter As IntPtr, ByVal Command As ComplexSpeechCommand)
            monitor.writeLineIf(logParsing, "Parsing command: " & Text)
            Dim nodes As String() = SeparateIntoNodes(Text)
            Dim hStates(nodes.Length) As IntPtr
            hStates(0) = hStateBefore
            hStates(hStates.Length - 1) = hStateAfter
            For i As Integer = 1 To hStates.Length - 2
                builder.CreateNewState(hStateBefore, hStates(i))
            Next
            For i As Integer = 0 To nodes.Length - 1
                monitor.writeLineIf(logParsing, "Node: " & nodes(i))
                monitor.WriteIf(logParsing, "Node type: ")
                If nodes(i).Substring(0, 1) = "[" Then
                    Select Case nodes(i).Substring(1, 1)
                        Case "L"
                            monitor.writeLineIf(logParsing, "List")
                            Dim items As String() = SeparateListNodes(nodes(i).Substring(3, nodes(i).Length - 4))
                            For Each item As String In items
                                monitor.writeLineIf(logParsing, "    Adding item: " & item)
                                ParseCommandText(item, hStates(i), hStates(i + 1), Command)
                            Next
                        Case "?"
                            monitor.writeLineIf(logParsing, "Optional")
                            builder.AddWordTransition(hStates(i), hStates(i + 1), Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                            ParseCommandText(nodes(i).Substring(3, nodes(i).Length - 4), hStates(i), hStates(i + 1), Command)
                        Case "*"
                            monitor.writeLineIf(logParsing, "Dictation")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), New IntPtr(SPRULETRANS_DICTATION))
                        Case "#"
                            monitor.writeLineIf(logParsing, "Number")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), hStateDigit)
                        Case "G"
                            monitor.writeLineIf(logParsing, "Greek Alphabet")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            addMultipleTransition(min, max, hStates(i), hStates(i + 1), hStateGreekLetter)
                        Case "R"
                            monitor.writeLineIf(logParsing, "Rule reference")
                            Dim subRule As String = nodes(i).Substring(3, nodes(i).Length - 4)
                            If _commandMap.ContainsKey(subRule) Then
                                Dim hStateRule As IntPtr
                                builder.GetRule(subRule, 0, Nothing, False, hStateRule)
                                builder.AddRuleTransition(hStates(i), hStates(i + 1), hStateRule, 1, Nothing)
                                Command.AddDependency(subRule)
                            End If
                        Case "A"
                            monitor.writeLineIf(logParsing, "Phonetic alphabet - Not yet implemented")
                            Dim min As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(0)
                            Dim max As Integer = nodes(i).Substring(3, nodes(i).Length - 4).Split(" ")(1)
                            builder.AddWordTransition(hStates(i), hStates(i + 1), Nothing, Nothing, SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                            'Add phonetic alphabet rule reference
                        Case Else
                            monitor.writeLineIf(logErrors, "Unknown node type")
                    End Select
                Else
                    monitor.writeLineIf(logParsing, "Plaintext node")
                    builder.AddWordTransition(hStates(i), hStates(i + 1), nodes(i), " ", SPGRAMMARWORDTYPE.SPWT_LEXICAL, 1, Nothing)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Separates a command's text into its top-level nodes
        ''' </summary>
        ''' <param name="Text">Text of the command</param>
        ''' <returns>Array of strings for the command's nodes.</returns>
        ''' <remarks>
        ''' Optional and List nodes contain sub nodes that have to be processed separately. This 
        ''' function will not separate enclosed nodes.
        ''' </remarks>
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
                    monitor.writeLineIf(logNodeSplitting, "Limit exceeded. Error while parsing")
                End If
                Return nodeList.ToArray()
            Else 'Simple phrase
                monitor.writeLineIf(logNodeSplitting, "Simple phrase")
                Dim ret As String() = {Text}
                Return ret
            End If
        End Function

        ''' <summary>
        ''' Separates the elements of a list node
        ''' </summary>
        ''' <param name="Text">Text of the list node. Does not include outer brackets.</param>
        ''' <returns>Array of list elements as strings</returns>
        ''' <remarks>
        ''' The list nodes cannot simply be separated with <see cref="Split">String.Split()</see>
        ''' if the list contains a second list within it. This function solves that problem.
        ''' </remarks>
        Public Function SeparateListNodes(ByVal Text As String) As String()
            Dim splitList As New List(Of String)
            splitList.AddRange(Text.Split("|"))
            If Text.Contains("[L") Then
                Dim nodeList As New List(Of String)
                Dim openCount As Integer = 0
                Dim closeCount As Integer = 0
                Dim currentNode As String = ""
                For Each myString As String In splitList
                    If openCount <> closeCount Then
                        'Some previous node was a list, we need a separator
                        currentNode = currentNode & "|"
                    End If
                    If myString.Contains("[") Or myString.Contains("]") Then
                        'Count brackets
                        For Each c As Char In myString.ToCharArray
                            If c = "[" Then
                                openCount += 1
                            ElseIf c = "]" Then
                                closeCount += 1
                            End If
                        Next
                    End If
                    currentNode = currentNode & myString
                    If openCount = closeCount Then
                        nodeList.Add(currentNode)
                        currentNode = ""
                    End If
                Next
                Return nodeList.ToArray()
            Else
                'No subnodes with lists
                Return splitList.ToArray()
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

        ''' <summary>
        ''' Removes a command or rule from the grammar
        ''' </summary>
        ''' <param name="Name">Name of command or rule to remove</param>
        ''' <remarks>
        ''' The command or rule will not be removed in the following cases:
        ''' <list>
        ''' <item>The command or rule does not exist</item>
        ''' <item>The command or rule is internal</item>
        ''' <item>There are other commands or rules that depend on the rule to be removed.</item>
        ''' </list>
        ''' In these cases removeCommand() will fail silently.
        ''' </remarks>
        Public Sub removeCommand(ByVal Name As String)
            If _commandMap.ContainsKey(Name) And Not _internalCommands.Contains(Name) Then
                monitor.writeLineIf(logRemovals, "Removing command: " & Name)
                Dim found As Boolean = False
                For Each myRule As ComplexSpeechCommand In _commandMap.Values
                    If myRule.DependsOn(Name) Then found = True
                Next
                If Not found Then
                    Dim hStateRemoved As IntPtr
                    'Creates rule if it doesn't exist, otherwise returns current rule
                    builder.GetRule(Name, 0, 0, True, hStateRemoved)
                    builder.ClearRule(hStateRemoved)
                    builder.Commit(0)
                Else
                    monitor.writeLineIf(logRemovals, "Cannot remove rule while other rules depend on it.")
                End If
            Else
                monitor.writeLineIf(logRemovals, "Cannot remove command that is internal or nonexistant.")
            End If
        End Sub

        Private Sub onRecognition(ByVal StreamNumber As Integer, ByVal StreamPosition As Object, ByVal RecognitionType As SpeechRecognitionType, ByVal Result As ISpeechRecoResult)
            monitor.writeLineIf(logRecognitions, "Recognized: " & Result.PhraseInfo.GetText())
            monitor.writeLineIf(logRecognitionsDetail, "    Rule name: " & Result.PhraseInfo.Rule.Name)
            monitor.writeLineIf(logRecognitionsDetail, "    Rule ID: " & Result.PhraseInfo.Rule.Id)
            If Result.PhraseInfo.Rule.Name = "Main" Then
                If Result.PhraseInfo.Rule.Children.Item(0).Children Is Nothing Then
                    monitor.writeLineIf(logGeneral, "Listening for continuation of command.")
                    grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSActive)
                Else
                    If Not _continuousCommands Then
                        grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSInactive)
                    End If
                    Dim ruleName As String = Result.PhraseInfo.Rule.Children.Item(0).Children.Item(0).Name
                    If _commandMap.ContainsKey(ruleName) Then
                        _commandMap.Item(ruleName).executeCommand(Result.PhraseInfo)
                    End If
                End If
            ElseIf Result.PhraseInfo.Rule.Name = "Command" Then
                If Result.PhraseInfo.Rule.Children Is Nothing Then
                    monitor.writeLineIf(logErrors, "Something went wrong! Bad recognition structure.")
                Else
                    If Not _continuousCommands Then
                        grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSInactive)
                    End If
                    Dim ruleName As String = Result.PhraseInfo.Rule.Children.Item(0).Name
                    If _commandMap.ContainsKey(ruleName) Then
                        _commandMap.Item(ruleName).executeCommand(Result.PhraseInfo)
                    End If
                End If
            End If
        End Sub
        ''' <summary>
        ''' Temporary function for debugging. Lists all rules currently in the grammar.
        ''' </summary>
        Public Sub DisplayGrammar()
            monitor.writeLineIf(True, "Rule listing:")
            For i As Integer = 0 To grammar.Rules.Count - 1
                monitor.writeLineIf(True, "    " & grammar.Rules.Item(i).Name)
            Next
        End Sub

        ''' <summary>
        ''' Validates a command's text.
        ''' </summary>
        ''' <param name="CommandText">Text to validate</param>
        ''' <returns><see cref="CommandTextValidationError.NO_ERROR">NO_ERROR</see> if valid.</returns>
        ''' <remarks>
        ''' This function checks the command's validity in several ways:
        ''' 
        ''' Any brackets are checked to see if they are properly matched and contain node information
        ''' in the correct format.
        ''' 
        ''' Any nodes used are checked against the list of valid nodes. Their contents are also checked
        ''' for the correct format for their node type
        ''' 
        ''' Any complex nodes used have their contents independantly checked for validity.
        ''' 
        ''' Any rules referenced are checked to see if they exist.
        ''' </remarks>
        Public Function IsValidCommand(ByVal CommandText As String) As CommandTextValidationError
            'Check bracket balancing and format - Enough to pass to node separators
            Dim chars As Char() = CommandText.ToCharArray()
            Dim i As Integer = 0
            Dim openCount As Integer = 0
            Dim closeCount As Integer = 0
            Try
                While i < chars.Length
                    If chars(i) = "[" Then
                        openCount += 1
                        'Validate node format
                        If chars(i + 2) <> " " Then
                            Return CommandTextValidationError.ERROR_INVALID_NODE_FORMAT
                        End If
                    End If
                    If chars(i) = "]" Then
                        closeCount += 1
                    End If
                    i += 1
                End While
            Catch ex As IndexOutOfRangeException
                Return CommandTextValidationError.ERROR_INVALID_NODE_FORMAT
            End Try
            If openCount = closeCount Then
                'Brackets balance and have valid node formats
                For Each myNode As String In SeparateIntoNodes(CommandText)
                    If myNode.StartsWith("[") Then
                        'Special node, needs additional validation
                        Dim type As String = myNode.Substring(1, 1)
                        If type = "L" Then
                            'Separate list nodes and validate each one
                            For Each myListNode In SeparateListNodes(myNode.Substring(3, myNode.Length - 4))
                                Dim isValid As CommandTextValidationError = IsValidCommand(myListNode)
                                If Not isValid = CommandTextValidationError.NO_ERROR Then
                                    Return isValid
                                End If
                            Next
                        ElseIf type = "?" Then
                            'Check contents
                            Dim isValid As CommandTextValidationError = IsValidCommand(myNode.Substring(3, myNode.Length - 4))
                            If Not isValid = CommandTextValidationError.NO_ERROR Then
                                Return isValid
                            End If
                        ElseIf type = "*" Or type = "#" Or type = "G" Or type = "A" Then
                            Dim parts As String() = myNode.Substring(3, myNode.Length - 4).Split(" ")
                            If parts.Length <> 2 Then
                                Return CommandTextValidationError.ERROR_INCORRECT_NODE_CONTENTS
                            End If
                            If Not (IsNumeric(parts(0)) And IsNumeric(parts(1))) Then
                                Return CommandTextValidationError.ERROR_INCORRECT_NODE_CONTENTS
                            End If
                        ElseIf type = "R" Then
                            Dim rule As String = myNode.Substring(3, myNode.Length - 4)
                            If Not _commandMap.ContainsKey(rule) Then
                                Return CommandTextValidationError.ERROR_UNDEFINED_RULE_REFERENCED
                            End If
                        Else
                            'Unknown node type
                            Return CommandTextValidationError.ERROR_UNKNOWN_NODE_TYPE
                        End If
                    End If
                Next
            Else
                Return CommandTextValidationError.ERROR_BRACKET_MISMATCH
            End If
            'Everything passed
            Return CommandTextValidationError.NO_ERROR
        End Function

        Public Property ContinuousCommands() As Boolean
            Get
                Return _continuousCommands
            End Get
            Set(ByVal value As Boolean)
                _continuousCommands = value
                If _continuousCommands Then
                    grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSActive)
                    monitor.writeLineIf(logGeneral, "Continous commands active.")
                Else
                    grammar.CmdSetRuleState("Command", SpeechRuleState.SGDSInactive)
                    monitor.writeLineIf(logGeneral, "Continous commands disabled.")
                End If
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Possible values for a command text validation to produce
    ''' </summary>
    Public Enum CommandTextValidationError
        ''' <summary>
        ''' No errors; the command text is a valid rule.
        ''' </summary>
        NO_ERROR = 0
        ''' <summary>
        ''' Incorrect number of opening and closing brackets
        ''' </summary>
        ERROR_BRACKET_MISMATCH = 1
        ''' <summary>
        ''' Node does not follow correct format. (Bracket followed by one character and a space)
        ''' </summary>
        ERROR_INVALID_NODE_FORMAT = 2
        ''' <summary>
        ''' Unknown node identifier used.
        ''' </summary>
        ERROR_UNKNOWN_NODE_TYPE = 3
        ''' <summary>
        ''' Node contents are incorrect for the node type.
        ''' </summary>
        ERROR_INCORRECT_NODE_CONTENTS = 4
        ''' <summary>
        ''' A rule reference is to an undefined rule. (Case-sensitive)
        ''' </summary>
        ERROR_UNDEFINED_RULE_REFERENCED = 5
    End Enum
End Namespace
