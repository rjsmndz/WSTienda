Imports System.IO
Imports System.Text
Imports DataAccess
Imports Newtonsoft.Json

Module Common
    Public DA As New TDataAccess()
    Public Sub Verifica_Mensaje(ByVal mensaje As String)
        If mensaje <> "" Then
            Throw New Exception(mensaje)
        End If
    End Sub

    Public Function JsonSerializer(Of T)(ByVal obj As T) As String
        Try
            Dim serialize As New JsonSerializer()
            Dim sb As New StringBuilder
            Dim sw As New StringWriter(sb)


            Dim writer As JsonWriter = New JsonTextWriter(sw)
            writer.Formatting = Formatting.Indented

            serialize.Serialize(writer, obj)

            Return sb.ToString
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function JsonDeserializer(Of T)(ByVal json As String, ByRef obj As T) As String
        Try

            Dim serialize As New JsonSerializer()
            Dim reader As JsonTextReader = New JsonTextReader(New StringReader(json))
            obj = serialize.Deserialize(Of T)(reader)

            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Function booleanToInt(ByVal value As Boolean) As Integer
        If value Then
            Return 1
        Else
            Return 0
        End If
    End Function
    Function intToBoolean(ByVal value As Integer) As Boolean
        If value = 1 Then
            Return True
        Else
            Return False
        End If
    End Function
    Function booleanToString(ByVal value As Boolean) As String
        If value Then
            Return "si"
        Else
            Return "no"
        End If
    End Function
#Region "Encripta y desencripta password del Portal de aplicaciones"


    Public Function UnLockPassword(ByVal key As String) As String
        Dim sResultado As String = Convert.ToString(Text.Encoding.UTF8.GetChars(Convert.FromBase64String(key)))
        Return sResultado.Substring(0, sResultado.Length)
    End Function

    Public Function LockPassword(ByVal key As String) As String
        Dim ValueAndSalt As String = String.Concat(key)
        Dim byteValue() As Byte = Text.Encoding.UTF8.GetBytes(ValueAndSalt)
        Return (Convert.ToBase64String(byteValue))
    End Function
#End Region


#Region "Encripta y desencripta password de punto de venta"

    Public Function Desencriptar(ByVal clave As String) As String
        Dim Cadena As String = ""

        Do While Trim(clave) <> ""
            Cadena = Chr(Val(Mid(clave, 2, (Mid(clave, 1, 1)) / 2))) & Cadena
            clave = Mid(clave, (Val(Mid(clave, 1, 1)) / 2) + 2)
        Loop

        Return Cadena
    End Function

    Public Function Encriptar(ByVal clave As String) As String
        Dim Cadena As String = ""
        Dim Largo As Integer
        Dim i As Integer

        Largo = Len(Trim(clave))

        For i = 1 To Largo
            Cadena = Trim(Str(Len(Trim(Str(Asc(Mid(clave, i, 1))))) * 2)) + Trim(Str(Asc(Mid(clave, i, 1)))) + Cadena
        Next i

        Return Cadena
    End Function
#End Region



#Region "Encripacion y desencriptacion"
    Public Function Encriptacion_unida(ByVal key As String)
        Try
            Return Encriptar(key)
        Catch ex As Exception
            Return LockPassword(key)
        End Try
    End Function

    Public Function Desencriptacion_unida(ByVal key As String)
        Try
            Return Desencriptar(key)
        Catch ex As Exception

            Return UnLockPassword(key)
        End Try
    End Function

#End Region
End Module
