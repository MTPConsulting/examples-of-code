<ServiceContract()>
Public Interface IPersonas

    <OperationContract()>
    Function ObtenerCargosMba(ByVal usuario As String,
                                     ByVal clave As String) As DataSet

    <OperationContract()>
    Function AltaPersona(ByVal datosPersona As DatosPersona,
                                ByVal datosPuntuales As DatosPuntualesPersona) As Integer

    <OperationContract()>
    Function ModiPersona(ByVal datosPersona As DatosPersona,
                                ByVal datosPuntuales As DatosPuntualesPersona,
                                ByVal modiPuntual As Boolean) As Boolean

    <OperationContract()>
    Function BajaPersona(ByVal datosPersona As DatosPersona) As Boolean

    <OperationContract()>
    Function BajaPersonaConMail(ByVal datosPersona As DatosPersona,
                                       ByVal cd_persona_autoriz As Integer) As Boolean

    <OperationContract()>
    Function RecordatorioClaveEnvio(ByVal datosPersona As DatosPersona) As Boolean

    <OperationContract()>
    Function EsAutorizador(ByVal datosPersona As DatosPersona) As Boolean

    <OperationContract()>
    Function ObtenerProductosAutorizador(ByVal datosPersona As DatosPersona,
                                         ByVal cd_pais As String) As DataSet

End Interface

'Contrato de datos básico y general
'usuario y clave se pasan siempre. El resto depende si se necesita o no en el método que lo utiliza.
<DataContract()>
Public Class DatosPersona

    <DataMember()>
    Public Property usuario() As String

    <DataMember()>
    Public Property clave() As String

    <DataMember()>
    Public Property cliente() As Integer

    <DataMember()>
    Public Property persona() As Integer

    <DataMember()>
    Public Property portal() As String

End Class

'Datos de la persona
<DataContract()>
Public Class DatosPuntualesPersona

    <DataMember()>
    Public Property apellido() As String

    <DataMember()>
    Public Property nombres() As String

    <DataMember()>
    Public Property mailPrincipal() As String

    <DataMember()>
    Public Property sector() As String

    <DataMember()>
    Public Property zona() As String

    <DataMember()>
    Public Property sucursal() As String

End Class

