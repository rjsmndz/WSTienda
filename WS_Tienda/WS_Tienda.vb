' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "Service1" en el código y en el archivo de configuración a la vez.
Imports WS_Tienda
Imports BusinessTienda

Public Class WS_Tienda
    Implements IWS_Tienda

    Public Function SelectCategorias(pEmpId As Integer) As WS_Respuesta Implements IWS_Tienda.SelectCategorias
        Dim respuesta As New WS_Respuesta
        Try
            Dim categoria As New TCategoria(pEmpId)
            Dim mensaje As String = ""
            mensaje = categoria.SelectTable("categoria")
            Verifica_Mensaje(mensaje)

            respuesta.MensajeRespuesta = mensaje
            respuesta.Datos = JsonSerializer(Of DataTable)(categoria.Data.Tables("categoria"))
            Return respuesta
        Catch ex As Exception
            respuesta.MensajeRespuesta = ex.Message
            respuesta.Datos = ""
            Return respuesta
        End Try

    End Function

    Public Function SelectCategoriasPrincipales(pEmpId As Integer) As WS_Respuesta Implements IWS_Tienda.SelectCategoriasPrincipales
        Dim respuesta As New WS_Respuesta
        Try
            Dim categoria As New TCategoria(pEmpId)
            Dim mensaje As String = ""
            mensaje = categoria.SelectTablePrincipal("categoria")
            Verifica_Mensaje(mensaje)

            respuesta.MensajeRespuesta = ""
            respuesta.Datos = JsonSerializer(Of DataTable)(categoria.Data.Tables("categoria"))

            Return respuesta

        Catch ex As Exception
            respuesta.MensajeRespuesta = ex.Message
            Return respuesta
        End Try
    End Function
End Class
