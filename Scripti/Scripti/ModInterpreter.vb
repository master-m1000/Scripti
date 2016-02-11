Module ModInterpreter

    ''' <summary>
    ''' Interprets a line
    ''' </summary>
    ''' <param name="line">Line to interpret</param>
    ''' <param name="lineNo">Line number</param>
    ''' <returns>Success or an error</returns>
    Function InterpretLine(ByVal line As String, ByVal lineNo As Integer) As ScriptEventInfo
        Try
            If line.StartsWith("'") = False And String.IsNullOrWhiteSpace(line) = False Then 'Ignore comments
                If line.Contains("%ask") Then 'Ask for userinputs
                    Dim patternAsk As String = "([^%]*)%ask([bis])(t|f|n)(.*)%([^%]*)"
                    Dim regexMACAsk As Text.RegularExpressions.MatchCollection = Text.RegularExpressions.Regex.Matches(line, patternAsk, Text.RegularExpressions.RegexOptions.IgnoreCase)
                    For Each match As Text.RegularExpressions.Match In regexMACAsk
                        Select Case match.Result("$2")
                            Case "b"
                                Dim askOption As YesNoQuestionDefault
                                Select Case match.Result("$3")
                                    Case "t"
                                        askOption = YesNoQuestionDefault.Yes
                                    Case "f"
                                        askOption = YesNoQuestionDefault.No
                                    Case "n"
                                        askOption = YesNoQuestionDefault.Nothing
                                End Select
                                line = match.Result("$1") & AskYesNo(match.Result("$4"), askOption) & match.Result("$5")
                            Case "i"
                                Dim askOption As Boolean = True
                                If match.Result("$3") = "f" Then
                                    askOption = False
                                End If
                                line = match.Result("$1") & AskInteger(match.Result("$4"), askOption) & match.Result("$5")
                            Case "s"
                                Dim askOption As Boolean = True
                                If match.Result("$3") = "f" Then
                                    askOption = False
                                End If
                                line = match.Result("$1") & AskString(match.Result("$4"), askOption) & match.Result("$5")
                        End Select
                    Next
                End If
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
                        InterpretLineGroupVar(regexMC, lineNo)
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
    ''' Interprets a line in the group "io"
    ''' </summary>
    ''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    Sub InterpretLineGroupIo(ByVal lineRegEx As Text.RegularExpressions.MatchCollection)
        Select Case lineRegEx(0).Result("$2")
            Case "copydirectory"
                If IO.Directory.Exists(lineRegEx(0).Result("$4")) = True Then 'If directory already exist then ...
                    Dim dir As New IO.DirectoryInfo(lineRegEx(0).Result("$3"))
                    IO.Directory.CreateDirectory(lineRegEx(0).Result("$4"))
                    For Each subdir In dir.GetDirectories("*", IO.SearchOption.AllDirectories)
                        IO.Directory.CreateDirectory(My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), subdir.FullName.Replace(dir.FullName, "")))
                    Next
                    For Each file In dir.GetFiles("*", IO.SearchOption.AllDirectories)
                        If IO.File.Exists(My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", ""))) = True Then 'if file already exist
                            If Boolean.TryParse(lineRegEx(0).Result("$5"), New Boolean) = True Then '... check if the script has an advice to overwrite it.
                                If Boolean.Parse(lineRegEx(0).Result("$5")) = True Then 'If the script says to overwrite then override (in other case nothing happens)
                                    IO.File.Copy(file.FullName, My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", "")), True)
                                End If
                            Else 'If the script says nothing then ...
                                If AskYesNo("Overwrite """ & My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", "")) & """ with """ & file.FullName & """?", YesNoQuestionDefault.No) = True Then '... ask the user.
                                    IO.File.Copy(file.FullName, My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", "")), True)
                                End If
                            End If
                        Else
                            IO.File.Copy(file.FullName, My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", "")))
                        End If

                        If IO.File.Exists(My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", ""))) Then
                            Console.WriteLine(My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName & "\", "")))
                        End If
                    Next
                Else
                    Dim dir As New IO.DirectoryInfo(lineRegEx(0).Result("$3"))
                    IO.Directory.CreateDirectory(lineRegEx(0).Result("$4"))
                    For Each subdir In dir.GetDirectories("*", IO.SearchOption.AllDirectories)
                        IO.Directory.CreateDirectory(My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), subdir.FullName.Replace(dir.FullName, "")))
                    Next
                    For Each file In dir.GetFiles("*", IO.SearchOption.AllDirectories)
                        IO.File.Copy(file.FullName, My.Computer.FileSystem.CombinePath(lineRegEx(0).Result("$4"), file.FullName.Replace(dir.FullName, "")))
                    Next
                End If
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
            Case "deletedirectory"
                If IO.Directory.Exists(lineRegEx(0).Result("$3")) = True Then
                    IO.Directory.Delete(lineRegEx(0).Result("$3"), True) 'Deletes all files
                Else
                    Throw New Exception("Can't delete directory. Directory doesn't exist: " & lineRegEx(0).Result("$3"))
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
            Case "run"
                Dim newProcess As New Process
                newProcess.StartInfo.FileName = lineRegEx(0).Result("$3")
                newProcess.StartInfo.Arguments = lineRegEx(0).Result("$4")
                Dim runAs As Boolean 'Run as administrator (Windows Vista+/UAC)
                If Boolean.TryParse(lineRegEx(0).Result("$6"), runAs) = True And runAs = True Then
                    newProcess.StartInfo.Verb = "runas"
                End If
                newProcess.Start()
                Dim waitForFinish As Boolean
                If Boolean.TryParse(lineRegEx(0).Result("$5"), waitForFinish) = True And waitForFinish = True Then
                    While newProcess.HasExited = False 'Wait for exit
                    End While
                End If
            Case Else
                Throw New Exception(lineRegEx(0).Result("Uknown command: $2"))
        End Select
    End Sub

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
            Case "exit"
                exitScript = True
            Case "license"
                Console.WriteLine("The MIT License (MIT)")
                Console.WriteLine("")
                Console.WriteLine("Copyright (c) 2016 Mario Wagenknecht")
                Console.WriteLine("")
                Console.WriteLine("Permission is hereby granted, free of charge, to any person obtaining a copy")
                Console.WriteLine("of this software and associated documentation files (the ""Software""), to deal")
                Console.WriteLine("in the Software without restriction, including without limitation the rights")
                Console.WriteLine("to use, copy, modify, merge, publish, distribute, sublicense, and/or sell")
                Console.WriteLine("copies of the Software, and to permit persons to whom the Software is")
                Console.WriteLine("furnished to do so, subject to the following conditions:")
                Console.WriteLine("")
                Console.WriteLine("The above copyright notice and this permission notice shall be included in all")
                Console.WriteLine("copies or substantial portions of the Software.")
                Console.WriteLine("")
                Console.WriteLine("THE SOFTWARE IS PROVIDED ""As Is"", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR")
                Console.WriteLine("IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,")
                Console.WriteLine("FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE")
                Console.WriteLine("AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER")
                Console.WriteLine("LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,")
                Console.WriteLine("OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE")
                Console.WriteLine("SOFTWARE.")
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
    ''' Interprets a line in the group "var"
    ''' </summary>
    ''' <param name="lineRegEx">Line to interpret as Regex MatchCollection</param>
    ''' <param name="lineNo">Line number</param>
    Sub InterpretLineGroupVar(ByVal lineRegEx As Text.RegularExpressions.MatchCollection, ByVal lineNo As Integer)
        Select Case lineRegEx(0).Result("$2")
            Case "boolean"
                Dim parsed As Boolean = False
                If Boolean.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                    If booleans.ContainsKey(lineRegEx(0).Result("$3")) = True Then
                        booleans(lineRegEx(0).Result("$3")) = parsed
                    Else
                        booleans.Add(lineRegEx(0).Result("$3"), parsed)
                    End If
                Else
                    Throw New Exception(lineRegEx(0).Result("The new value of $3 is not a boolean: $4"))
                End If
            Case "integer"
                Dim parsed As Integer
                If Integer.TryParse(lineRegEx(0).Result("$4"), parsed) = True Then
                    If integers.ContainsKey(lineRegEx(0).Result("$3")) = True Then
                        integers(lineRegEx(0).Result("$3")) = parsed
                    Else
                        integers.Add(lineRegEx(0).Result("$3"), parsed)
                    End If
                Else
                    Throw New Exception(lineRegEx(0).Result("The new value of $3 is not a integer: $4"))
                End If
            Case "string"
                If integers.ContainsKey(lineRegEx(0).Result("$3")) = True Then
                    strings(lineRegEx(0).Result("$3")) = lineRegEx(0).Result("$4")
                Else
                    strings.Add(lineRegEx(0).Result("$3"), lineRegEx(0).Result("$4"))
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
