Public MustInherit Class TBaseClass
    Protected _MessageError As String
    Protected _Data As DataSet
    Public Sub New()
        _MessageError = ""
        _Data = Nothing
    End Sub
    Public Sub Verifica_Mensaje()
        If _MessageError <> "" Then
            Throw New Exception(_MessageError)
        End If
    End Sub
    Public Sub Verifica_Mensaje(ByVal mensaje As String)
        If mensaje <> "" Then
            Throw New Exception(mensaje)
        End If
    End Sub
    Public MustOverride Function Insert() As String
    Public MustOverride Function Delete() As String
    Public MustOverride Function Update() As String
    Public MustOverride Function SelectTable(ByVal tableName As String) As String
    Public MustOverride Function Load() As String
End Class
