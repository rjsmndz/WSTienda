Imports DataAccess
Public Class TSubCategoria
    Inherits TBaseClass

    Property Emp_Id As Integer
    Property Categoria_Id As Integer
    Property SubCategoria_Id As Integer
    Property SubCategoria_Nombre As String

    Public Sub New(ByVal pEmpId As Integer)
        MyBase.New()
    End Sub

    Public Overrides Function Delete() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Insert() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Load() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SelectTable(tableName As String) As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Update() As String
        Throw New NotImplementedException()
    End Function
End Class
