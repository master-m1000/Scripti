Module ModInterpreter

    ''' <summary>
    ''' Interprets a line
    ''' </summary>
    ''' <param name="line">Line to interpret</param>
    ''' <returns>Success or an error</returns>
    Function InterpretLine(ByVal line As String) As ScriptEventInfo
        Try
            If line.StartsWith("'") = False And String.IsNullOrWhiteSpace(line) = False Then 'Ignore comments
                If line.Contains("%b") Or line.Contains("%i") Or line.Contains("%s") Then 'Only check lines if they contains the indicator for variables
                    line = ReplaceVariables(line)
                End If
                Dim pattern As String = "(\w+)\.(\w+):([^\n;]*);?([^\n;]*);?([^\n;]*);?([^\n;]*);?([^\n;]*)"
                Dim regexMC As Text.RegularExpressions.MatchCollection = Text.RegularExpressions.Regex.Matches(line, pattern, Text.RegularExpressions.RegexOptions.IgnoreCase)
                If regexMC.Count = 0 Then
                    Throw New Exception("Bad line")
                End If
                Select Case regexMC(0).Result("$1")
                    Case "me"
                        InterpretLineGroupMe(regexMC)
                    Case "io"
                        InterpretLineGroupIo(regexMC)
                    Case "var"
                        InterpretLineGroupVar(regexMC)
                    Case Else
                        Throw New Exception(regexMC(0).Result("Uknown group: $1"))
                End Select
            End If
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Interprets a line in the group "me"
    ''' </summary>
    ''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    Sub InterpretLineGroupMe(ByVal lineRegEx As Text.RegularExpressions.MatchCollection)
        Select Case lineRegEx(0).Result("$2")
            Case "clear"
                Console.Clear()
            Case "color"
                Console.BackgroundColor = CType(lineRegEx(0).Result("$3"), ConsoleColor)
                Console.ForegroundColor = CType(lineRegEx(0).Result("$4"), ConsoleColor)
            Case "pause"
                Pause()
            Case "resetcolor"
                Console.ResetColor()
            Case "write"
                Console.WriteLine(lineRegEx(0).Result("$3"))
            Case "exit"
                exitScript = True
            Case Else
                Throw New Exception(lineRegEx(0).Result("Uknown command: $2"))
        End Select
    End Sub

    ''' <summary>
    ''' Interprets a line in the group "io"
    ''' </summary>
    ''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    Sub InterpretLineGroupIo(ByVal lineRegEx As Text.RegularExpressions.MatchCollection)
        Select Case lineRegEx(0).Result("$2")
            Case "copyfile"
                If IO.File.Exists(lineRegEx(0).Result("$4")) = True Then 'If file already exist then ...
                    If Boolean.TryParse(lineRegEx(0).Result("$5"), New Boolean) = True Then '... check if the script has an advice to overwrite it.
                        If Boolean.Parse(lineRegEx(0).Result("$5")) = True Then 'If the script says to overwrite then override (in other case nothing happens)
                            IO.File.Copy(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"), True)
                        End If
                    Else 'If the script says nothing then ...
                        If AskYesNo("Overwrite """ & lineRegEx(0).Result("$4") & """ with a copy of """ & lineRegEx(0).Result("$3") & """?", YesNoQuestionDefault.No) = True Then '... ask the user.
                            IO.File.Copy(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"), True)
                        End If
                    End If
                Else
                    IO.File.Copy(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                End If
            Case "movefile"
                If IO.File.Exists(lineRegEx(0).Result("$4")) = True Then 'If file already exist then ...
                    If Boolean.TryParse(lineRegEx(0).Result("$5"), New Boolean) = True Then '... check if the script has an advice to overwrite it.
                        If Boolean.Parse(lineRegEx(0).Result("$5")) = True Then 'If the script says to overwrite then override (in other case nothing happens)
                            IO.File.Delete(lineRegEx(0).Result("$4"))
                            IO.File.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                        End If
                    Else 'If the script says nothing then ...
                        If AskYesNo("Overwrite """ & lineRegEx(0).Result("$4") & """ with """ & lineRegEx(0).Result("$3") & """?", YesNoQuestionDefault.No) = True Then '... ask the user.
                            IO.File.Delete(lineRegEx(0).Result("$4"))
                            IO.File.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                        End If
                    End If
                Else
                    IO.File.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                End If
            Case "deletefile"
                If IO.File.Exists(lineRegEx(0).Result("$3")) = True Then
                    IO.File.Delete(lineRegEx(0).Result("$3"))
                Else
                    Throw New Exception("Can't delete file. File doesn't exist: " & lineRegEx(0).Result("$3"))
                End If
            Case "movedirectory"
                If IO.Directory.Exists(lineRegEx(0).Result("$4")) = True Then 'If file already exist then ...
                    If Boolean.TryParse(lineRegEx(0).Result("$5"), New Boolean) = True Then '... check if the script has an advice to overwrite it.
                        If Boolean.Parse(lineRegEx(0).Result("$5")) = True Then 'If the script says to overwrite then override (in other case nothing happens)
                            IO.Directory.Delete(lineRegEx(0).Result("$4"))
                            IO.Directory.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                        End If
                    Else 'If the script says nothing then ...
                        If AskYesNo("Overwrite """ & lineRegEx(0).Result("$4") & """ with """ & lineRegEx(0).Result("$3") & """?", YesNoQuestionDefault.No) = True Then '... ask the user.
                            IO.Directory.Delete(lineRegEx(0).Result("$4"))
                            IO.Directory.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                        End If
                    End If
                Else
                    IO.Directory.Move(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                End If
            Case "deletedirectory"
                If IO.Directory.Exists(lineRegEx(0).Result("$3")) = True Then
                    IO.Directory.Delete(lineRegEx(0).Result("$3"), True) 'Deletes all files
                Else
                    Throw New Exception("Can't delete directory. Directory doesn't exist: " & lineRegEx(0).Result("$3"))
                End If
            Case Else
                Throw New Exception(lineRegEx(0).Result("Uknown command: $2"))
        End Select
    End Sub

    ''' <summary>
    ''' Interprets a line in the group "var"
    ''' </summary>
    ''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    Sub InterpretLineGroupVar(ByVal lineRegEx As Text.RegularExpressions.MatchCollection)
        Select Case lineRegEx(0).Result("$2")
            Case "boolean"
                Dim parsed As Boolean = False
                If booleans.ContainsKey(lineRegEx(0).Result("$3")) = False Then
                    If Boolean.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                        booleans.Add(lineRegEx(0).Result("$3"), parsed)
                    Else
                        Throw New Exception(lineRegEx(0).Result("The new value of $3 is not a boolean: $4"))
                    End If
                Else
                    If booleans.ContainsKey(lineRegEx(0).Result("$3")) Then
                        If Boolean.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                            booleans(lineRegEx(0).Result("$3")) = parsed
                        Else
                            Throw New Exception(lineRegEx(0).Result("The new value of $3 is not a boolean: $4"))
                        End If
                    Else
                        Throw New Exception(lineRegEx(0).Result("Variable $3 is already used."))
                    End If
                End If
            Case "integer"
                Dim parsed As Integer = 0
                If integers.ContainsKey(lineRegEx(0).Result("$3")) = False Then
                    If Integer.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                        integers.Add(lineRegEx(0).Result("$3"), parsed)
                    Else
                        Throw New Exception(lineRegEx(0).Result("The new value of $3 is not an integer: $4"))
                    End If
                Else
                    If integers.ContainsKey(lineRegEx(0).Result("$3")) Then
                        If Integer.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                            integers(lineRegEx(0).Result("$3")) = parsed
                        Else
                            Throw New Exception(lineRegEx(0).Result("The new value of $3 is not an integer: $4"))
                        End If
                    Else
                        Throw New Exception(lineRegEx(0).Result("Variable $3 is already used."))
                    End If
                End If
            Case "string"
                If strings.ContainsKey(lineRegEx(0).Result("$3")) = False Then
                    strings.Add(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
                Else
                    strings(lineRegEx(0).Result("$3")) = lineRegEx(0).Result("$4")
                End If
            Case Else
                Throw New Exception(lineRegEx(0).Result("Uknown command: $2"))
        End Select
    End Sub

    '''' <summary>
    '''' Interprets a line in the group "template"
    '''' </summary>
    '''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    'Sub InterpretLineGroupTemplate(ByVal lineRegEx As Text.RegularExpressions.MatchCollection)
    '    Select Case lineRegEx(0).Result("$2")
    '        Case "example"
    '            'doSomething(lineRegEx(0).Result("$3"))
    '        Case Else
    '            Throw New Exception(lineRegEx(0).Result("Uknown command: $2"))
    '    End Select
    'End Sub

End Module
