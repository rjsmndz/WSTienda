' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
<ServiceContract()>
Public Interface IWS_Tienda

#Region "Categoria"
    <OperationContract()>
    Function SelectCategorias(ByVal pEmpId As Integer) As WS_Respuesta

    <OperationContract()>
    Function SelectCategoriasPrincipales(ByVal pEmpId As Integer) As WS_Respuesta

#End Region

End Interface

' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.
' Puede agregar archivos XSD al proyecto. Después de compilar el proyecto, puede usar directamente los tipos de datos definidos aquí, con el espacio de nombres "WS_Tienda.ContractType".

<DataContract()>
Public Class WS_Respuesta

    <DataMember()>
    Public Property MensajeRespuesta() As String

    <DataMember()>
    Public Property Datos() As String

End Class
