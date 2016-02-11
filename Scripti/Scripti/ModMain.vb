Module ModMain
    Public consoleMode As Boolean = False
    Public debugMode As Boolean = False
    Public exitScript As Boolean = False
    Public scripts As New List(Of IO.FileInfo)
    Public booleans As New Dictionary(Of String, Boolean)
    Public integers As New Dictionary(Of String, Integer)
    Public strings As New Dictionary(Of String, String)
    Private ignoreVersion As Boolean = False
    Private wait As Boolean = False

    Public Enum YesNoQuestionDefault
        [Nothing]
        Yes
        No
    End Enum

    Sub Main()
        With Initialize()
            If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                ShowError(.Error)
            End If
        End With

        If consoleMode = True Then
            HLine()
            While exitScript = False
                Console.WriteLine()
                Console.Write("> ")
                With InterpretLine(Console.ReadLine, -1)
                    If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                        ShowError(.Error, -1)
                    End If
                End With
            End While
        End If

        For Each script In scripts
            HLine()
            Console.WriteLine(" Script: " & script.ToString)
            HLine()
            With InterpretScript(script)
                If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                    ShowError(.Error)
                End If
            End With
            HLine()
            Console.WriteLine()
        Next

        If wait = True Then
            Pause()
        End If
    End Sub

    ''' <summary>
    ''' Initalizes Scripti
    ''' </summary>
    ''' <returns>Success or an error</returns>
    Function Initialize() As ScriptEventInfo
        Try
            Console.Title = "Scripti Version " & My.Application.Info.Version.ToString
            Console.WriteLine("Scripti Version " & My.Application.Info.Version.ToString)
            Console.WriteLine("Copyright (c) 2015 Mario Wagenknecht")
            Console.WriteLine("Source at: https://github.com/master-m1000/Scripti")
            Console.WriteLine()
            For Each argument In Environment.GetCommandLineArgs
                Select Case argument
                    Case Environment.GetCommandLineArgs(0)
                    Case "console"
                        consoleMode = True
                    Case "debug"
                        debugMode = (argument = "debug")
                        Console.WriteLine("DEBUG MODE")
                        Console.WriteLine()
                    Case "wait"
                        wait = True
                    Case Else
                        With AddScript(argument)
                            If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                                ShowError(.Error)
                            End If
                        End With
                End Select
            Next
            If consoleMode = True Then
                Console.WriteLine("CONSOLE MODE")
                scripts.Clear()
            ElseIf consoleMode = False And scripts.Count = 0 Then
                Console.Write("Open a script file: ")
                With AddScript(Console.ReadLine)
                    If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                        ShowError(.Error, -1)
                        wait = True
                    End If
                End With
            Else
                Console.WriteLine("List of script files:")
                For Each script In scripts
                    Console.WriteLine(" " & script.ToString)
                Next
            End If
            Console.WriteLine()

            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Interprets a script
    ''' </summary>
    ''' <param name="script"></param>
    ''' <returns>Success or an error</returns>
    Function InterpretScript(ByVal script As IO.FileInfo) As ScriptEventInfo
        Dim lineNo As Integer = 0
        Try
            exitScript = False
            For Each line In IO.File.ReadAllLines(script.ToString)
                lineNo += 1
                If lineNo = 1 Then
                    If line.StartsWith("|SCRIPTI SCRIPT FILE VERSION ") = True Then
                        If Not line.Replace("|SCRIPTI SCRIPT FILE VERSION ", "") = My.Application.Info.Version.ToString And ignoreVersion = False Then
                            Console.WriteLine("This script is not written for this version of Scripti.")
                            If AskYesNo("Do you want to continue?", YesNoQuestionDefault.No) = False Then
                                Exit For
                            End If
                        End If
                    Else
                        Throw New Exception("Not a valid script file.")
                    End If
                Else
                    With InterpretLine(line, lineNo)
                        If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
                            ShowError(.Error, lineNo)
                        End If
                    End With
                End If
                If exitScript = True Then
                    Exit For
                End If
            Next
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Shows an error
    ''' </summary>
    ''' <param name="ex">Exception</param>
    ''' <returns>Success or an error</returns>
    Function ShowError(ByVal ex As Exception, Optional ByVal lineNo As Integer = -1) As ScriptEventInfo
        Try
            If lineNo = -1 Then
                If debugMode = True Then
                    Console.WriteLine("ERROR: " & ex.ToString)
                Else
                    Console.WriteLine("ERROR: " & ex.Message)
                End If
            Else
                If debugMode = True Then
                    Console.WriteLine("ERROR Line " & lineNo & ": " & ex.ToString)
                Else
                    Console.WriteLine("ERROR Line " & lineNo & ": " & ex.Message)
                End If
            End If
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex2 As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex2}
        End Try
    End Function

    ''' <summary>
    ''' Wait for a pressed key until continue
    ''' </summary>
    ''' <returns>Success or an error</returns>
    Function Pause() As ScriptEventInfo
        Try
            Console.Write("Press any key to continue . . . ")
            Console.ReadKey(True)
            Console.WriteLine()
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Asks a Yes/No-Question
    ''' </summary>
    ''' <param name="question">A Yes/No-Question</param>
    ''' <param name="defaultAnswer">Default Answer if enter pressed</param>
    ''' <returns>Boolean</returns>
    Function AskYesNo(ByVal question As String, Optional ByVal defaultAnswer As YesNoQuestionDefault = YesNoQuestionDefault.Nothing) As Boolean
        Dim answer As Boolean
        While True
            Select Case defaultAnswer
                Case YesNoQuestionDefault.Nothing
                    Console.Write(question & " [y/n] ")
                Case YesNoQuestionDefault.Yes
                    Console.Write(question & " [Y/n] ")
                Case YesNoQuestionDefault.No
                    Console.Write(question & " [y/N] ")
            End Select
            Dim x As ConsoleKeyInfo = Console.ReadKey
            Select Case x.Key
                Case ConsoleKey.Y
                    answer = True
                    Exit While
                Case ConsoleKey.N
                    answer = False
                    Exit While
                Case ConsoleKey.Enter
                    Select Case defaultAnswer
                        Case YesNoQuestionDefault.Yes
                            answer = True
                            Exit While
                        Case YesNoQuestionDefault.No
                            answer = False
                            Exit While
                        Case YesNoQuestionDefault.Nothing
                            answer = Nothing
                    End Select
                Case Else
                    answer = Nothing
            End Select
            Console.WriteLine()
        End While
        Console.WriteLine()
        Return answer
    End Function

    ''' <summary>
    ''' Asks for an Integer
    ''' </summary>
    ''' <param name="question">Text before input</param>
    ''' <param name="acceptNegative">If negative inputs should be accepted</param>
    ''' <returns>Integer</returns>
    Function AskInteger(ByVal question As String, Optional ByVal acceptNegative As Boolean = True) As Integer
        Dim answer As Integer
        While True
            Console.Write(question & " ")
            If Integer.TryParse(Console.ReadLine(), answer) = True Then
                If acceptNegative = True Or answer >= 0 Then
                    Exit While
                End If
            End If
        End While
        Return answer
    End Function

    ''' <summary>
    ''' Asks for a String
    ''' </summary>
    ''' <param name="question">Text before input</param>
    ''' <param name="acceptEmpty">If empty Strings should be accepted</param>
    ''' <returns>String</returns>
    Function AskString(ByVal question As String, Optional ByVal acceptEmpty As Boolean = True) As String
        Dim answer As String = String.Empty
        While True
            Console.Write(question & " ")
            answer = Console.ReadLine()
            If acceptEmpty = True Or String.IsNullOrEmpty(answer) = False Then
                Exit While
            End If
        End While
        Return answer
    End Function

    ''' <summary>
    ''' Draws a horizontal line.
    ''' </summary>
    ''' <returns></returns>
    Function HLine() As ScriptEventInfo
        Try
            For i As Integer = 1 To Console.WindowWidth
                Console.Write("-")
            Next
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Adds a script to the script files list
    ''' </summary>
    ''' <param name="path">Path of the script file</param>
    ''' <returns>Success or an error</returns>
    Function AddScript(ByVal path As String) As ScriptEventInfo
        Try
            Dim script As New IO.FileInfo(path)
            If script.Exists = True Then
                scripts.Add(script)
            Else
                Throw New Exception("Could not find script: " & script.ToString)
            End If
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
        Catch ex As Exception
            Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
        End Try
    End Function

    ''' <summary>
    ''' Replaces variables with value
    ''' </summary>
    ''' <param name="line">Line to edit</param>
    ''' <returns>The line with replaced variables.</returns>
    Function ReplaceVariables(ByRef line As String) As String
        Dim xLine As String = line
        While xLine.Length > 0
            If xLine(0) = ":" Then
                xLine = xLine.Remove(0, 1)
                Exit While
            Else
                xLine = xLine.Remove(0, 1)
            End If
        End While
        Dim pattern2 As String = "([^\n%]*)%([bis])([^%]+)%([^\n%]*)"
        Dim regexMC2 As Text.RegularExpressions.MatchCollection = Text.RegularExpressions.Regex.Matches(xLine, pattern2, Text.RegularExpressions.RegexOptions.IgnoreCase)
        For Each match As Text.RegularExpressions.Match In regexMC2
            Select Case match.Result("$2")
                Case "b"
                    If booleans.Keys.Contains(match.Result("$3")) = True Then
                        line = line.Replace("%b" & match.Result("$3") & "%", booleans(match.Result("$3")).ToString)
                    Else
                        Throw New Exception("The boolean """ & match.Result("$3") & """ does not exist.")
                    End If
                Case "i"
                    If integers.Keys.Contains(match.Result("$3")) = True Then
                        line = line.Replace("%i" & match.Result("$3") & "%", integers(match.Result("$3")).ToString)
                    Else
                        Throw New Exception("The integer """ & match.Result("$3") & """ does not exist.")
                    End If
                Case "s"
                    If strings.Keys.Contains(match.Result("$3")) = True Then
                        line = line.Replace("%s" & match.Result("$3") & "%", strings(match.Result("$3")))
                    Else
                        Throw New Exception("The string """ & match.Result("$3") & """ does not exist.")
                    End If
            End Select
        Next
        Return line
    End Function

    'Templates
    '
    '''' <summary>
    '''' This is a template
    '''' </summary>
    '''' <returns>Success or an error</returns>
    'Function Template() As ScriptEventInfo
    '    Try
    '        'Do something
    '        Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Success}
    '    Catch ex As Exception
    '        Return New ScriptEventInfo With {.Sucess = ScriptEventInfo.ScriptEventState.Error, .Error = ex}
    '    End Try
    'End Function
    '
    'With Template()
    'If .Sucess = ScriptEventInfo.ScriptEventState.Error Then
    '            ShowError(.Error)
    '        End If
    'End With

End Module