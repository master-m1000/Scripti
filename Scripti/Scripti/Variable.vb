Public Class Variable
    Public Enum VarType
        [Boolean]
        [Integer]
        [String]
    End Enum

    Property type As Type = Nothing
    Property Value As Object = Nothing

    Function ToBoolean() As Boolean
        If Value.GetType Is New Boolean.GetType Then
            Return CBool(Value)
        ElseIf Value.GetType Is New Integer.GetType Then
            Select Case CInt(Value)
                Case 0
                    Return False
                Case 1
                    Return True
                Case Else
                    Throw New Exception("Can not convert " & Value.ToString & "into a boolean.")
            End Select
        ElseIf Value.GetType Is New String(Nothing).GetType Then
            Select Case Value.ToString.ToLower
                Case "false"
                    Return False
                Case "true"
                    Return True
                Case Else
                    Throw New Exception("Can not convert " & Value.ToString & "into a boolean.")
                    Return Nothing
            End Select
        Else
            Throw New Exception("Can not convert " & Value.ToString & "into a boolean.")
            Return Nothing
        End If
    End Function

    Function ToInteger() As String
        Return CType(Value, String)
    End Function

    Overrides Function ToString() As String
        Return Value.ToString
    End Function
End Class
