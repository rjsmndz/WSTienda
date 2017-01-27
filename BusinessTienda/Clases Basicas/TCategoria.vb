Imports DataAccess
Public Class TCategoria
    Inherits TBaseClass

#Region "Propiedades"
    Property EmpId As Integer
    Property CategoriaId As Integer
    Property CategoriaNombre As String
    Property CategoriaPrincipal As String
    Shadows Property Data As DataSet
#End Region
#Region "Constructor"
    Public Sub New()
        MyBase.New()
        EmpId = 0
        CategoriaId = 0
        CategoriaNombre = ""
        CategoriaPrincipal = "no"
        Data = New DataSet
    End Sub
    Public Sub New(ByVal pEmpId As Integer)
        MyBase.New()
        EmpId = pEmpId
        CategoriaId = 0
        CategoriaNombre = ""
        CategoriaPrincipal = "no"
        Data = New DataSet
    End Sub
#End Region
#Region "Funciones Publicas"

#End Region
    Public Overrides Function Delete() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Insert() As String
        Try
            Dim query As String = ""
            query = "select isnull(max(categoria_id),0)+1 categoria_id from categoria where emp_id = " & EmpId
            DA.TransactionBegin()
            _MessageError = DA.GetIntegerValue(query, CategoriaId, True)
            Verifica_Mensaje()

            query = "insert into categoria(emp_id, categoria_id, categoria_nombre,categoria_principal)
                    values(" & EmpId & "," & CategoriaId & ",'" & CategoriaNombre & "','" & CategoriaPrincipal & "')"
            _MessageError = DA.ExecuteQuery(query, True)
            Verifica_Mensaje()

            DA.TransactionCommit()
            Return _MessageError
        Catch ex As Exception
            DA.TransactionRollback()
            Return ex.Message
        End Try
    End Function

    Public Overrides Function Load() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SelectTable(tableName As String) As String
        Try
            Dim query As String = ""
            query = "select Categoria_Id Numero, Categoria_Nombre Nombre from Categoria where emp_id = " & EmpId
            _MessageError = DA.GetDataSet(tableName, Data, query, True)
            Verifica_Mensaje()

            Return _MessageError
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function SelectTablePrincipal(tableName As String) As String
        Try
            Dim query As String = ""
            query = "select Categoria_Id Numero, Categoria_Nombre Nombre from Categoria where emp_id = " & EmpId & " and Categoria_Principal = 'si'"
            _MessageError = DA.GetDataSet(tableName, Data, query, True)
            Verifica_Mensaje()

            Return _MessageError
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Overrides Function Update() As String
        Throw New NotImplementedException()
    End Function
End Class
