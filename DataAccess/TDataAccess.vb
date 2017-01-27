Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports Microsoft.Win32
Imports System.Configuration


Imports System.Xml
Public Class TDataAccess

    Private _Connection As New SqlConnection
    Private _Reader As SqlDataReader
    Private _Transaction As SqlTransaction
    Private _Command As New SqlCommand
    Private _Adapter As New SqlDataAdapter()
    Private _User As String
    Private _Password As String
    Private _Comp_id As Integer 'esta variable se utiliza para la replicacion
    Private _ReplicaLocal As String
    Private _ReplicaServer As String
    Private _PuntoVenta As Boolean
    Private _BaseDatosName As String
    Private _ServerName As String

    Public Property Comp_id() As Integer
        Get
            Return _Comp_id
        End Get
        Set(ByVal value As Integer)
            _Comp_id = value
        End Set
    End Property
    Public ReadOnly Property Reader() As SqlDataReader
        Get
            Return _Reader
        End Get
    End Property

    Public ReadOnly Property Connection() As SqlConnection
        Get
            Return _Connection
        End Get
    End Property
    Public ReadOnly Property Transaction() As SqlTransaction
        Get
            Return _Transaction
        End Get
    End Property
    Public ReadOnly Property ConnectionIsOpen() As Boolean
        Get
            If _Connection.State = ConnectionState.Open Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property User() As String
        Get
            Return _User
        End Get
    End Property
    Public ReadOnly Property Password() As String
        Get
            Return LockPassword(_Password)
        End Get
    End Property

    Public Function UnLockPassword(ByVal key As String) As String
        Dim sResultado As String = Convert.ToString(Text.Encoding.UTF8.GetChars(Convert.FromBase64String(key)))
        Return sResultado.Substring(0, sResultado.Length)
    End Function

    Public Function LockPassword(ByVal key As String) As String
        Dim ValueAndSalt As String = String.Concat(key)
        Dim byteValue() As Byte = Text.Encoding.UTF8.GetBytes(ValueAndSalt)
        Return (Convert.ToBase64String(byteValue))
    End Function

    Private Function GetConecctionString() As String
        Try
            Return ConfigurationManager.ConnectionStrings("Conexion").ConnectionString
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Function OpenConnection() As String
        Try
            If _Connection.ConnectionString = "" Then
                _Connection.ConnectionString = GetConecctionString()
            End If
            If _Connection.State = ConnectionState.Closed Then
                _Connection.Open()
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function CloseConnection() As String
        Try
            If _Connection.State = ConnectionState.Open Then
                _Connection.Close()
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Sub New()
        _Connection.ConnectionString = ""
    End Sub

    Public Function ExecuteQuery(ByVal Query As String, Optional ByVal isInTransaction As Boolean = True, Optional ByVal pReplica_Local As Boolean = False, Optional ByVal pReplica_Local_id As Integer = -1) As String
        Try

            _Command.Connection = _Connection
            _Command.CommandTimeout = 1000000000
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            _Command.CommandType = CommandType.Text
            _Command.Transaction = _Transaction

            'si replica a los locales, inserta en una bitacora todos los queries que se van pasar a los locales
            If _PuntoVenta = False Then
                If (_ReplicaLocal = "si") And (pReplica_Local = True) Then
                    Dim sql As String
                    Dim sQuery As String = Query.Replace("'", "''")
                    'ejecuta el query que inserta en la bitacora
                    sql = " Insert into dbo.Replica_Local(comp_id,Str_Id,Sqlquery) " & _
                          " select " & _Comp_id & ",str_id,'" & sQuery & "'" & _
                          " from  store with (nolock) " & _
                          " where  comp_id = " & _Comp_id

                    'si se esta filtrando por local
                    If pReplica_Local_id <> -1 Then
                        sql = sql & " and str_id = " & pReplica_Local_id
                    End If

                    _Command.CommandText = sql
                    _Command.ExecuteNonQuery()
                End If
            Else 'si es algo del punto de venta inserta en la tabla de replicacion para el server
                If (_ReplicaServer = "si") Then
                    Dim sql As String
                    Dim sQuery As String = Query.Replace("'", "''")
                    'ejecuta el query que inserta en la bitacora
                    sql = " Insert into dbo.Replica_Server(comp_id,Sqlquery) " & _
                          " values ( " & _Comp_id & ",'" & sQuery & "')"
                    _Command.CommandText = sql
                    _Command.ExecuteNonQuery()
                End If
            End If

            'ejecuta el query que esta recibiendo po parametro
            _Command.CommandText = Query
            _Command.ExecuteNonQuery()
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function ExecuteQueryReplicaLocal(ByVal Query As String, Optional ByVal isInTransaction As Boolean = True, Optional ByVal pReplica_Local As Boolean = False, Optional ByVal pReplica_Local_id As Integer = -1) As String
        Try

            _Command.Connection = _Connection
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            _Command.CommandType = CommandType.Text
            _Command.Transaction = _Transaction

            'si replica a los locales, inserta en una bitacora todos los queries que se van pasar a los locales
            Dim sql As String
            Dim sQuery As String = Query.Replace("'", "''")
            'ejecuta el query que inserta en la bitacora
            sql = " Insert into dbo.Replica_Local(comp_id,Str_Id,Sqlquery) " & _
                  " select " & _Comp_id & ",str_id,'" & sQuery & "'" & _
                  " from  store with (nolock) " & _
                  " where  comp_id = " & _Comp_id

            'si se esta filtrando por local
            If pReplica_Local_id <> -1 Then
                sql = sql & " and str_id = " & pReplica_Local_id
            End If

            _Command.CommandText = sql
            _Command.ExecuteNonQuery()

            'ejecuta el query que esta recibiendo po parametro
            _Command.CommandText = Query
            _Command.ExecuteNonQuery()
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function TransactionBegin() As String
        Try
            Dim ErrorMessage = OpenConnection()
            If ErrorMessage <> "" Then
                Throw New Exception(ErrorMessage)
            End If
            _Transaction = _Connection.BeginTransaction
            Return ""
        Catch ex As Exception
            CloseConnection()
            Return ex.Message
        End Try
    End Function
    Public Function TransactionCommit() As String
        Try
            _Transaction.Commit()
            Return ""
        Catch ex As Exception
            Return ex.Message
        Finally
            CloseConnection()
        End Try
    End Function

    Public Function TransactionRollback() As String
        Try
            _Transaction.Rollback()
            Return ""
        Catch ex As Exception
            Return ex.Message
        Finally
            CloseConnection()
        End Try
    End Function
    Public Function GetStringValue(ByVal Query As String, ByRef Value As String, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = ""
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToString(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetIntegerValue(ByVal Query As String, ByRef Value As Integer, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = 0
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToInt64(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetDateValue(ByVal Query As String, ByRef Value As DateTime, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = Date.Now
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToDateTime(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetLongValue(ByVal Query As String, ByRef Value As Long, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = 0
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToInt32(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetBooleanValue(ByVal Query As String, ByRef Value As Boolean, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = 0
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToBoolean(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function GetDoubleValue(ByVal Query As String, ByRef Value As Double, Optional ByVal isInTransaction As Boolean = True) As String
        Try
            Value = 0
            _Command.Connection = _Connection
            _Command.CommandText = Query
            If isInTransaction = True Then
                _Command.Transaction = _Transaction
            End If
            Value = System.Convert.ToDouble(_Command.ExecuteScalar())
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function ReaderClose(Optional ByVal pCloseConection As Boolean = True) As String
        Try
            If _Reader.IsClosed = False Then
                _Reader.Close()
            End If
            Return ""
        Catch ex As Exception
            Return ex.Message
        Finally
            If pCloseConection = True Then
                CloseConnection()
            End If
        End Try
    End Function

    Public Function ReaderOpen(ByVal query As String, Optional ByVal pOpenConection As Boolean = True) As String
        Try
            If pOpenConection = True Then
                OpenConnection()
            Else
                _Command.Transaction = _Transaction
            End If

            _Command.CommandText = query
            _Command.Connection = _Connection
            _Reader = _Command.ExecuteReader
            _Reader.Read()
            Return ""
        Catch ex As Exception
            CloseConnection()
            Return ex.Message
            _Reader.Close()
        End Try
    End Function

    Public Function ReaderOpenTrans(ByVal query As String) As String
        Try
            _Command.CommandText = query
            _Command.Connection = _Connection
            _Command.Transaction = _Transaction
            _Reader = _Command.ExecuteReader
            _Reader.Read()
            Return ""
        Catch ex As Exception
            Return ex.Message
            _Reader.Close()
        End Try
    End Function

    Public Sub DeleteDataTable(ByVal TableName As String, ByRef pData As DataSet)
        Try
            If pData.Tables.Contains(TableName) Then
                pData.Tables.Remove(TableName)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function GetDataSet(ByVal TableName As String, ByRef pData As DataSet, ByVal query As String, Optional ByVal pOpenConection As Boolean = True) As String
        Try
            _Command.CommandText = query
            _Command.Connection = _Connection
            _Command.CommandTimeout = 1000000000
            _Adapter.SelectCommand = _Command

            If pOpenConection = True Then
                OpenConnection()
            Else
                _Command.Transaction = _Transaction
            End If

            DeleteDataTable(TableName, pData)

            _Adapter.Fill(pData, TableName)
            Return ""
        Catch ex As Exception
            Return ex.Message
        Finally
            If pOpenConection = True Then
                CloseConnection()
            End If
        End Try
    End Function

    Public Function GetDataEschema(ByVal TableName As String, ByRef pdata As DataSet, ByVal query As String) As String
        Try
            If _Connection.ConnectionString = "" Then
                _Connection.ConnectionString = GetConecctionString()
            End If
            _Command.CommandText = query
            _Command.Connection = _Connection
            _Adapter.SelectCommand = _Command

            DeleteDataTable(TableName, pdata)
            _Adapter.FillSchema(pdata, SchemaType.Mapped, TableName)
            Return ""
        Catch ex As Exception
            Return ex.Message
        Finally
        End Try
    End Function

    Public Function ConnectionAvailable() As Boolean
        Try
            If _Connection.ConnectionString = "" Then
                _Connection.ConnectionString = GetConecctionString()
            End If
            If _Connection.State = ConnectionState.Closed Then
                _Connection.Open()
            End If
            Return True
        Catch ex As Exception
            Return False
        Finally
            If _Connection.State = ConnectionState.Open Then
                _Connection.Close()
            End If
        End Try
    End Function
End Class

