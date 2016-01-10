Public Class ScriptEventInfo

    Public Enum ScriptEventState
        Unknown
        Success
        [Error]
    End Enum

    Property Sucess As New ScriptEventState
    Property [Error] As New Exception

    Sub New()
        [Error] = Nothing
    End Sub

End Class
