Module ModInterpreter

    ''' <summary>
    ''' Interprets a line
    ''' </summary>
    ''' <param name="line">Line to interpret</param>
    ''' <returns>Success or an error</returns>
    Function InterpretLine(ByVal line As String) As ScriptEventInfo
        Try
            If line.StartsWith("'") = False And String.IsNullOrWhiteSpace(line) = False Then 'Ignore comments
                Dim pattern As String = "([^\n.:;]+)\.(.+):([^\n;]*);?([^\n;]*);?([^\n;]*);?([^\n;]*);?([^\n;]*)"
                Dim regexMC As Text.RegularExpressions.MatchCollection = Text.RegularExpressions.Regex.Matches(line, pattern, Text.RegularExpressions.RegexOptions.IgnoreCase)
                If regexMC.Count = 0 Then
                    Throw New Exception("Bad line")
                End If
                Select Case regexMC(0).Result("$1")
                    Case "me"
                        InterpretLineGroupMe(regexMC)
                    Case "io"
                        InterpretLineGroupIo(regexMC)
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
            Case ""

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
