Imports System.IO

Public Class Personas
    Implements IPersonas

#Region "Declaraciones"

    ''' <summary>
    ''' Métodos genéricos para todos los servicios ofrecidos
    ''' </summary>
    Private commonWCF As New CommonWCF

    ''' <summary>
    ''' Constructor público del WCF (para inicializaciones necesarias)
    ''' </summary>
    Public Sub New()
    End Sub

#End Region

#Region "Métodos Públicos"

#Region "ABM Personas"

    ''' <summary>
    ''' Realiza el alta de la persona y retorna su ID asignado.
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <param name="datosPuntuales">Estructura de datos que contiene el nombre y apellido de la persona, el sector, mail, zona y sucursal</param>
    ''' <returns>
    ''' Código de la persona dada de alta.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Public Function AltaPersona(ByVal datosPersona As DatosPersona,
                                ByVal datosPuntuales As DatosPuntualesPersona) As Integer Implements IPersonas.AltaPersona

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Dim cd_persona As Integer = 0

            Try
                cd_persona = _AltaPersona(datosPersona.usuario, datosPersona.cliente, datosPuntuales.apellido, datosPuntuales.nombres,
                                          datosPuntuales.mailPrincipal, datosPuntuales.sector, datosPuntuales.zona, datosPuntuales.sucursal)
            Catch ex As Exception
                Throw ex
            End Try

            'Ok
            Return cd_persona

        Else
            Throw New Exception("Logueo inválido")
        End If

    End Function

    ''' <summary>
    ''' Realiza una modificación de los datos de la persona identificada por su código.
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <param name="datosPuntuales">Estructura de datos que contiene el nombre y apellido de la persona, el sector, mail, zona y sucursal</param>
    ''' <param name="modiPuntual">Indica si se trata de la modi de datos puntual de una persona o bien es una modificación general</param>
    ''' <returns>
    ''' TRUE si realizó la operación con éxito.
    ''' FALSE si no realizó la operación pero no ocurrieron errores.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Public Function ModiPersona(ByVal datosPersona As DatosPersona,
                                ByVal datosPuntuales As DatosPuntualesPersona,
                                ByVal modiPuntual As Boolean) As Boolean Implements IPersonas.ModiPersona

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Dim rta As Boolean

            Try
                rta = _ModiPersona(datosPersona.usuario, datosPersona.cliente, datosPersona.persona, datosPuntuales.apellido,
                                   datosPuntuales.nombres, datosPuntuales.mailPrincipal, datosPuntuales.sector, datosPuntuales.zona,
                                   datosPuntuales.sucursal, modiPuntual)
            Catch ex As Exception
                Throw ex
            End Try

            'Ok
            Return rta

        Else
            Throw New Exception("Logueo inválido")
        End If

    End Function

    ''' <summary>
    ''' Realiza una baja de la persona identificada por su código.
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <returns>
    ''' TRUE si realizó la operación con éxito.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Public Function BajaPersona(ByVal datosPersona As DatosPersona) As Boolean Implements IPersonas.BajaPersona

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Try

                Dim pers As New Mba.Gestion.BR.Maestros.Personas

                Dim dsPersonas As DataSet
                Dim sWhere As String
                Dim sError As String = ""

                'Carga el dataset principal de la clase de negocios
                dsPersonas = pers.getDataSet()

                'Busca los datos de la persona para llenar el dataset
                sWhere = "cd_cliente = " & datosPersona.cliente.ToString
                sWhere &= " and cd_persona = " & datosPersona.persona.ToString
                sWhere &= " and st_estado='A'"
                dsPersonas = pers.Buscar(sWhere)

                '...verifica que exista
                If SystemFunctions.Files.IsEmpty(dsPersonas) Then
                    'sError = String.Format("ERR01#La persona {0} no existe como activa dentro del cliente {1} y por lo tanto no puede ser eliminada.#", cd_persona, cd_cliente)
                    Return True 'asumo que no es error
                End If


                '-----------------------------
                'PROCESO
                '-----------------------------
                If sError <> String.Empty Then
                    'Reporta el error
                    Throw New Exception(sError)
                Else
                    'Elimina la persona
                    dsPersonas.Tables(0).Rows(0).Delete()

                    'Graba utilizando las reglas del negocio
                    pers.ActualizarDatos(dsPersonas)
                End If

                'Ok
                Return True

            Catch ex As Exception
                Throw ex
            End Try

        Else
            Throw New Exception("Logueo inválido")
        End If

    End Function

    ''' <summary>
    ''' Realiza una baja de la persona identificada por su código.
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <param name="cd_persona_autoriz">Código de la persona autorizadora</param>
    ''' <returns>
    ''' TRUE si realizó la operación con éxito.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Public Function BajaPersonaConMail(ByVal datosPersona As DatosPersona,
                                       ByVal cd_persona_autoriz As Integer) As Boolean Implements IPersonas.BajaPersonaConMail

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Try

                Dim sWhere As String

                Dim persBaja As New Mba.Gestion.BR.Maestros.Personas
                Dim dsPersonasBaja As DataSet
                Dim nm_persona_baja As String = ""
                Dim st_aut_fc As String = ""
                Dim nm_cliente As String = ""
                Dim dsCliente As DataSet
                Dim clien As New Mba.Gestion.BR.Maestros.Clientes
                Dim persAut As New Mba.Gestion.BR.Maestros.Personas
                Dim dsPersonasAut As DataSet
                Dim nm_persona_autoriz As String = ""

                'Carga el dataset principal de la clase de negocios
                dsPersonasBaja = persBaja.getDataSet()

                'Busca los datos de la persona para llenar el dataset
                sWhere = "cd_cliente = " & datosPersona.cliente.ToString
                sWhere &= " and cd_persona = " & datosPersona.persona.ToString
                sWhere &= " and st_estado='A'"
                dsPersonasBaja = persBaja.Buscar(sWhere)

                '...verifica que exista
                If SystemFunctions.Files.IsEmpty(dsPersonasBaja) Then
                    Return True 'asumo que no es error
                End If

                nm_persona_baja = dsPersonasBaja.Tables(0).Rows(0).Item("nm_apellido").ToString & " " & dsPersonasBaja.Tables(0).Rows(0).Item("nm_nombres").ToString
                st_aut_fc = dsPersonasBaja.Tables(0).Rows(0).Item("st_autorizador_factura").ToString

                ' Si la persona a dar de baja es autorizadora de facturas entonces averiguo todo como para mandar mail
                If st_aut_fc = "S" Then

                    ' Consigo dato de cliente
                    dsCliente = clien.getDataSet()

                    sWhere = "cd_cliente = " & datosPersona.cliente.ToString
                    dsCliente = clien.Buscar(sWhere)

                    If SystemFunctions.Files.IsEmpty(dsCliente) Then
                        Return True 'asumo que no es error
                    End If

                    nm_cliente = dsCliente.Tables(0).Rows(0).Item("nm_cliente").ToString

                    'Consigo dato autorizador
                    dsPersonasAut = persAut.getDataSet()

                    'Busca los datos de la persona para llenar el dataset
                    sWhere = "cd_cliente = " & datosPersona.cliente.ToString
                    sWhere &= " and cd_persona = " & cd_persona_autoriz.ToString
                    sWhere &= " and st_estado='A'"
                    dsPersonasAut = persAut.Buscar(sWhere)

                    '...verifica que exista
                    If SystemFunctions.Files.IsEmpty(dsPersonasAut) Then
                        Return True 'asumo que no es error
                    End If

                    nm_persona_autoriz = dsPersonasAut.Tables(0).Rows(0).Item("nm_apellido").ToString _
                                            & " " & dsPersonasAut.Tables(0).Rows(0).Item("nm_nombres").ToString

                End If

                If BajaPersona(datosPersona) Then
                    If st_aut_fc = "S" Then
                        commonWCF.EnviaMailAvisoADMI(datosPersona.cliente.ToString, nm_cliente, nm_persona_baja,
                                                     nm_persona_autoriz, datosPersona.persona, datosPersona.portal)
                    End If
                End If


                'Ok
                Return True

            Catch ex As Exception
                Throw ex
            End Try

        Else
            Throw New Exception("Logueo inválido")
        End If

    End Function

#End Region

#Region "Autorizador"

    ''' <summary>
    ''' Retorna TRUE o FALSE indicando si una persona es o no autorizadora de alguna UA del portal.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <returns>
    ''' TRUE si es autorizador
    ''' FALSE caso contrario o por error
    ''' </returns>
    Public Function EsAutorizador(ByVal datosPersona As DatosPersona) As Boolean Implements IPersonas.EsAutorizador

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Dim nExiste As Double = 0
            Try

                Dim pers As New Mba.Gestion.BR.Maestros.Personas

                'Busca si la persona es autorizadora de algúna UA del portal
                Dim sSelect, sFrom, sCond As String
                sSelect = "count(*)"
                sFrom = "uas join Productos pro on pro.cd_producto = uas.cd_producto"
                sFrom &= "   join UaAutorizadores aut on aut.cd_cliente = uas.cd_cliente and aut.cd_ua = uas.cd_ua"
                sCond = "Uas.st_estado='A' and pro.st_estado='A'"
                sCond &= " and pro.cd_portal = " & SF.Strings.StringToSql(datosPersona.portal)
                sCond &= " and aut.cd_persona = " & SF.Number.NumberToSql(datosPersona.persona)
                sCond &= " and uas.cd_cliente = " & SF.Number.NumberToSql(datosPersona.cliente)

                nExiste = SF.Number.IsNumNull(Sql.Search(ConnectionString, sSelect, sFrom, sCond))

            Catch ex As Exception
                nExiste = 0
            End Try

            'True o False, dependiendo de la cantidad
            Return (nExiste > 0)

        Else
            Throw New Exception("Logueo inválido")
        End If
    End Function

    ''' <summary>
    ''' Retorna una lista de los productos que el autorizador puede administrar para un cliente dado.
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <param name="cd_pais">código de 2 letras del pais a filtrar (opcional)</param>
    ''' <returns>
    ''' Dataset con la información solicitada.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Public Function ObtenerProductosAutorizador(ByVal datosPersona As DatosPersona,
                                                     ByVal cd_pais As String) As DataSet Implements IPersonas.ObtenerProductosAutorizador

        If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

            Dim dsRdo As DataSet = Nothing

            Try
                dsRdo = commonWCF.ObtenerProductosAutorizador(datosPersona.portal, datosPersona.cliente, datosPersona.persona, cd_pais)
            Catch ex As Exception
                Throw ex
            End Try

            'Ok
            Return dsRdo

        Else
            Throw New Exception("Logueo inválido")
        End If

    End Function

#End Region

#Region "Auxiliares"

    ''' <summary>
    ''' Envía un mail a la persona recordándole su clave
    ''' </summary>
    ''' <param name="datosPersona">Estructura de datos que contiene el usuario y clave para login, el cliente, persona y portal (opcional)</param>
    ''' <returns>
    ''' TRUE si finalizó OK
    ''' FALSE ante un error
    ''' </returns>
    Public Function RecordatorioClaveEnvio(ByVal datosPersona As DatosPersona) As Boolean Implements IPersonas.RecordatorioClaveEnvio

        Dim PathPlantillaMail As String = String.Empty
        Dim PlantillaMail As String = String.Empty
        Dim ds As DataSet
        Dim sMailTo As String = String.Empty
        Dim sMailFrom As String = String.Empty
        Dim sCdUsuario As String = String.Empty
        Dim sCdPassword As String = String.Empty
        Dim sNmPersona As String = String.Empty
        Dim sSubject As String = String.Empty

        Try
            If commonWCF.Login(datosPersona.usuario, datosPersona.clave) Then

                Dim _param As New StartFrame.BR.Utilitarios.Parametros

                Dim sNombreEmpresa As String = _param.getParametro("NOMBRE_EMPRESA")

                PathPlantillaMail = StartFrame.SystemFunctions.Files.AddBackSlash(_param.getParametro("DIRPLANTILLAMAILS"))

                'Si la persona es de SOLO CURSOS entonces busco su plantilla especifica
                If datosPersona.portal = Portales.MBA Then
                    PlantillaMail = "Mail_RecordatorioClave.txt"
                    sSubject = sNombreEmpresa & " – Recordar mi usuario y contraseña"
                    sMailFrom = "MAILS_FROM_REG"
                Else
                    Dim sSoloCursos As String = String.Empty

                    sSoloCursos = CType(Sql.Search(ConnectionString, "st_solo_cursos", "Personas_SoloCursos", "cd_persona = " & datosPersona.persona.ToString()), String).Trim

                    sSubject = "Prevenciondelavado.com – Recordar mi usuario y contraseña"
                    sMailFrom = "MAILS_FROM_REGPDL"


                    If sSoloCursos = "NO" Then
                        PlantillaMail = "Mail_RecordatorioClave_pdl.txt"
                    Else
                        Dim cd_cliente As Integer

                        cd_cliente = CType(Sql.Search(ConnectionString, "cd_cliente", "Personas", "cd_persona = " & datosPersona.persona.ToString()), String).Trim

                        PlantillaMail = String.Format("Mail_RecordatorioClave_{0}.txt", cd_cliente.ToString())

                        If Not File.Exists(PathPlantillaMail & PlantillaMail) Then
                            PlantillaMail = "Mail_RecordatorioClave_pdl.txt"
                        End If

                    End If

                End If

                ds = Sql.Select(ConnectionString, "cd_usuario, cd_password, de_mail_principal, nm_persona = RTRIM(nm_apellido) + ', ' + nm_nombres", "Personas", "cd_persona_id = " & StartFrame.SystemFunctions.Number.NumberToSql(datosPersona.persona))

                If ds.Tables(0).Rows.Count > 0 Then
                    sMailTo = ds.Tables(0).Rows(0).Item("de_mail_principal")
                    sCdUsuario = ds.Tables(0).Rows(0).Item("cd_usuario")
                    sCdPassword = ds.Tables(0).Rows(0).Item("cd_password")
                    sNmPersona = ds.Tables(0).Rows(0).Item("nm_persona")
                End If

                'Cuerpo del mail
                Dim htmlfile As StreamReader
                htmlfile = New StreamReader(PathPlantillaMail & PlantillaMail, System.Text.Encoding.UTF8)

                '...plantilla
                Dim sMailBody As String = String.Empty

                sMailBody = htmlfile.ReadToEnd

                '...reemplaza variables generales
                sMailBody = sMailBody.Replace("@persona@", sNmPersona)
                sMailBody = sMailBody.Replace("@usuario@", sCdUsuario)
                sMailBody = sMailBody.Replace("@password@", sCdPassword)

                Common.MBA.Mails.Send(sMailTo, sSubject, sMailBody, , sMailFrom, , False)
            End If

        Catch ex As Exception

            Return False

        End Try


        Return True

    End Function

    ''' <summary>
    ''' Devuelve la lista de cargos de las personas
    ''' </summary>
    ''' <param name="usuario">Usuario válido con permisos para publicar</param>
    ''' <param name="clave">Clave del usuario</param>
    ''' <returns>Lista de cargos</returns>
    Public Function ObtenerCargosMba(ByVal usuario As String,
                                 ByVal clave As String) As DataSet Implements IPersonas.ObtenerCargosMba

        Dim dsRdo As DataSet = Nothing

        If commonWCF.Login(usuario, clave) Then

            Try
                dsRdo = _ObtenerCargosMba()
            Catch ex As Exception
                Throw ex
            End Try

            'Ok
            Return dsRdo

        Else
            Throw New Exception("Logueo inválido")
        End If


    End Function

#End Region

#End Region

#Region "Métodos Privados"

    ''' <summary>
    ''' Devuelve la lista de cargos
    ''' </summary>
    ''' <returns></returns>
    Private Function _ObtenerCargosMba() As DataSet

        Dim pers As New Mba.Gestion.BR.Maestros.Personas

        Dim dtAux As DataTable
        Dim dsRdo As New DataSet
        Dim param As New ArrayList

        Try

            'cargos
            param.Clear()
            dtAux = Sql.ExecSPDS(ConnectionString, "CargosMBA_Consul", param).Tables(0)
            dtAux.TableName = "CargosMBA"
            dsRdo.Merge(dtAux)

        Catch ex As Exception
            Throw ex
        End Try

        'Ok
        Return dsRdo

    End Function

    ''' <summary>
    ''' Realiza el alta de la persona y retorna su ID asignado.
    ''' </summary>
    ''' <param name="usuario">Usuario válido con permisos para publicar</param>
    ''' <param name="cd_cliente">Código de cliente</param>
    ''' <param name="nm_apellido">Apellido de la persona</param>
    ''' <param name="nm_nombres">Nombre de la persona</param>
    ''' <param name="de_mail_principal">Mail principal de la persona</param>
    ''' <param name="nm_sector">Sector de la empresa a la que pertenece de la persona</param>
    ''' <param name="nm_zona">Zona donde está ubicada la empresa a la que pertenece de la persona</param>
    ''' <param name="nm_sucursal">Sucursal de la empresa a la que pertenece de la persona</param>
    ''' <returns>
    ''' Código de la persona dada de alta.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Private Function _AltaPersona(ByVal usuario As String,
                                  ByVal cd_cliente As Integer,
                                  ByVal nm_apellido As String,
                                  ByVal nm_nombres As String,
                                  ByVal de_mail_principal As String,
                                  ByVal nm_sector As String,
                                  ByVal nm_zona As String,
                                  ByVal nm_sucursal As String) As Integer

        Try

            Dim pers As New Mba.Gestion.BR.Maestros.Personas

            Dim param As New ArrayList
            Dim cd_persona As Integer = 0
            Dim existe As Integer
            Dim sSelect, sFrom, sWhere As String
            Dim sError As String = ""

            de_mail_principal = de_mail_principal.Trim

            'Validaciones
            nm_apellido = pers.QuitarCaracteresEspeciales(nm_apellido)
            nm_nombres = pers.QuitarCaracteresEspeciales(nm_nombres)
            nm_sector = pers.QuitarCaracteresEspeciales(nm_sector) 'no estaba - CAR 2013-09-10
            nm_zona = pers.QuitarCaracteresEspeciales(nm_zona)
            nm_sucursal = pers.QuitarCaracteresEspeciales(nm_sucursal)

            '...datos vacíos
            If cd_cliente = 0 Or nm_apellido = String.Empty Or nm_nombres = String.Empty Or de_mail_principal = String.Empty Then
                sError = "ERR01#Datos mínimos requeridos de la persona incompletos.#"
            End If

            '...si ingresa zona y/o sucursal obligo a ingresar zona, sucursal y sector
            If sError = String.Empty Then
                If (nm_zona <> String.Empty OrElse nm_sucursal <> String.Empty) AndAlso (nm_zona = String.Empty OrElse nm_sucursal = String.Empty OrElse nm_sector = String.Empty) Then
                    sError = "ERR07#Datos mínimos requeridos de la persona incompletos (zona, sucursal y/o sector).#"
                End If
            End If

            '...longitud inválida de campos
            If sError = String.Empty Then
                If nm_apellido.Length > 100 OrElse nm_nombres.Length > 100 OrElse de_mail_principal.Length > 250 Then
                    sError = "ERR08#Longitud de los campos inválida (nombre, apellido, E-mail).#"
                End If
            End If
            If sError = String.Empty Then
                If nm_zona.Length > 50 OrElse nm_sucursal.Length > 50 OrElse nm_sector.Length > 50 Then
                    sError = "ERR09#Longitud de los campos inválida (zona, sucursal, sector).#"
                End If
            End If

            '...email válido
            If sError = String.Empty Then
                If Not pers.ValidarEmail(de_mail_principal) Then
                    sError = String.Format("ERR02#La dirección de email {0} es inválida.#", de_mail_principal)
                End If
            End If

            '...mail sin duplicados
            If sError = String.Empty Then
                sSelect = "cd_cliente"
                sFrom = "Personas"
                sWhere = "st_estado='A'"
                sWhere &= " and de_mail_principal = " & SF.Strings.StringToSql(de_mail_principal)
                Try
                    existe = CType(Sql.Search(ConnectionString, sSelect, sFrom, sWhere), Integer)
                Catch
                    existe = 0
                End Try
                If existe <> 0 Then
                    'Verifica si existe en el mismo cliente o en otro
                    If existe = cd_cliente Then
                        'existe para el mismo cliente
                        sError = "ERR03#Email duplicado con persona existente dentro del mismo cliente#"
                    Else
                        'existe para otro cliente
                        sError = String.Format("ERR04#Email duplicado con persona existente dentro de otro cliente ({0})#", existe.ToString.Trim)
                    End If
                End If
            End If

            '...dominio
            If sError = String.Empty Then
                '...verifica si la persona está registrada en algún otro producto además de CURSOS
                sSelect = "existe = count(*)"
                sFrom = "UaPersonasNominadas pn"
                sFrom &= " join uas u on u.cd_cliente=pn.cd_cliente and u.cd_ua=pn.cd_ua"
                sFrom &= " join Productos p on p.cd_producto=u.cd_producto"
                sWhere = "pn.st_estado='A'"
                sWhere &= " and p.cd_familia<>'CURSOS'"
                sWhere &= " and pn.cd_cliente=" & cd_cliente.ToString
                sWhere &= " and pn.cd_persona=" & cd_persona.ToString

                Try
                    existe = CType(Sql.Search(ConnectionString, sSelect, sFrom, sWhere), Integer)
                Catch
                    existe = 0
                End Try

                If existe > 0 Then
                    '...valida el dominio si tiene productos de cualq. otra familia
                    Dim sDominios As String = String.Empty
                    Dim sDomMail As String
                    sDomMail = de_mail_principal.Substring(de_mail_principal.IndexOf("@")).Trim

                    Dim dsDominios As DataSet
                    dsDominios = Sql.Select(ConnectionString, "nm_dominio", "ClienteDominios", "cd_cliente = " & SF.Strings.StringToSql(cd_cliente))
                    For Each rowDominio As DataRow In dsDominios.Tables(0).Rows
                        sDominios &= Trim(rowDominio(0)) & ";"
                    Next

                    If sDominios <> String.Empty _
                        AndAlso sDominios.ToLower.IndexOf(sDomMail.ToLower) = -1 Then
                        'El dominio del mail no figura entre los dominios del cliente
                        sError = String.Format("ERR06#El dominio {0} no figura entre los dominios del cliente#", sDomMail)
                    End If
                End If
            End If

            '-----------------------------
            'PROCESO
            '-----------------------------
            If sError <> String.Empty Then
                'Reporta el error
                Throw New Exception(sError)
            Else
                'Da de alta la persona
                param.Clear()
                param.Add(cd_cliente)           'Cliente
                param.Add(nm_apellido)          'Apellido
                param.Add(nm_nombres)           'Nombre
                param.Add(de_mail_principal)    'Mail principal
                param.Add(System.DBNull.Value)  'Sexo
                param.Add(nm_sector)            'Sector
                param.Add(System.DBNull.Value)  'Teléfono
                param.Add(usuario)              'Usuario
                param.Add(System.DBNull.Value)  'Roll
                param.Add(System.DBNull.Value)  'Cargo
                param.Add(System.DBNull.Value)  'Cargo mba
                param.Add(System.DBNull.Value)  'Geren_banco
                param.Add(System.DBNull.Value)  'Area negocio
                param.Add(nm_zona)  'Zona
                param.Add(nm_sucursal)  'Sucursal

                Sql.ExecSP(ConnectionString, "Persona_Alta", param)

                'Busca el cd_persona asignada

                'sWhere = "st_estado='A'"
                'sWhere &= " and cd_cliente = " & SF.Number.NumberToSql(cd_cliente)
                'sWhere &= " and nm_apellido = " & SF.Strings.StringToSql(nm_apellido)
                'sWhere &= " and nm_nombres = " & SF.Strings.StringToSql(nm_nombres)
                'sWhere &= " and de_mail_principal = " & SF.Strings.StringToSql(de_mail_principal)
                'cd_persona = CType(Sql.Search(ConnectionString, "cd_persona", "Personas", sWhere), Integer)

                Dim dtAux As DataTable
                param.Clear()
                param.Add(cd_cliente)           '@cd_cliente
                param.Add(nm_apellido)          '@ape
                param.Add(nm_nombres)           '@nombre
                param.Add(de_mail_principal)     '@de_mail_principal
                dtAux = Sql.ExecSPDS(ConnectionString, "Persona_consul_2", param).Tables(0)
                If dtAux.Rows.Count = 0 Then
                    cd_persona = 0
                Else
                    cd_persona = dtAux.Rows(0)("cd_persona")
                End If

            End If

            'Ok
            Return cd_persona

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' <summary>
    ''' Realiza una modificación de los datos de la persona identificada por su código.
    ''' </summary>
    ''' <param name="usuario">Usuario válido con permisos para publicar</param>
    ''' <param name="clave">Clave del usuario</param>
    ''' <param name="cd_cliente">Código de cliente</param>
    ''' <param name="cd_persona">Código de la persona a modificar</param>
    ''' <param name="nm_apellido">Apellido de la persona</param>
    ''' <param name="nm_nombres">Nombre de la persona</param>
    ''' <param name="de_mail_principal">Mail principal de la persona</param>
    ''' <param name="nm_sector">Sector de la empresa a la que pertenece de la persona</param>
    ''' <param name="nm_zona">Zona donde está ubicada la empresa a la que pertenece de la persona</param>
    ''' <param name="nm_sucursal">Sucursal de la empresa a la que pertenece de la persona</param>
    ''' <param name="modiPuntual">Indica si se trata de la modi de una única persona o es un proceso masivo</param>
    ''' <returns>
    ''' TRUE si realizó la operación con éxito.
    ''' FALSE si no realizó la operación pero no ocurrieron errores.
    ''' EXCEPCION ante un error.
    ''' </returns>
    Private Function _ModiPersona(ByVal usuario As String,
                                  ByVal cd_cliente As Integer,
                                  ByVal cd_persona As Integer,
                                  ByVal nm_apellido As String,
                                  ByVal nm_nombres As String,
                                  ByVal de_mail_principal As String,
                                  ByVal nm_sector As String,
                                  ByVal nm_zona As String,
                                  ByVal nm_sucursal As String,
                                  ByVal modiPuntual As Boolean) As Boolean

        Dim pers As New Mba.Gestion.BR.Maestros.Personas
        Dim cambios As New Mba.Manager.BR.Cambios

        Dim dsPersonas As DataSet
        Dim existe As Integer
        Dim sSelect, sFrom, sWhere, sCampos, sValores As String
        Dim sError As String = ""

        de_mail_principal = de_mail_principal.Trim

        Try
            'Carga el dataset principal de la clase de negocios
            dsPersonas = pers.getDataSet()

            'Busca los datos de la persona para llenar el dataset
            sWhere = "cd_cliente = " & cd_cliente.ToString
            sWhere &= " and cd_persona = " & cd_persona.ToString
            dsPersonas = pers.Buscar(sWhere)

            'Validaciones
            nm_apellido = pers.QuitarCaracteresEspeciales(nm_apellido)
            nm_nombres = pers.QuitarCaracteresEspeciales(nm_nombres)
            nm_sector = pers.QuitarCaracteresEspeciales(nm_sector) 'no estaba - CAR 2013-09-10
            nm_zona = pers.QuitarCaracteresEspeciales(nm_zona)
            nm_sucursal = pers.QuitarCaracteresEspeciales(nm_sucursal)

            '...modi inválida
            If SystemFunctions.Files.IsEmpty(dsPersonas) Then
                sError = String.Format("ERR05#La persona {0} no existe dentro del cliente {1} y por lo tanto no puede ser modificada.#",
                                       cd_persona, cd_cliente)
            End If

            '...datos vacíos
            If sError = String.Empty Then
                If cd_cliente = 0 Or nm_apellido = String.Empty Or nm_nombres = String.Empty Or de_mail_principal = String.Empty Or nm_sector = String.Empty Then
                    sError = "ERR01#Datos mínimos requeridos de la persona incompletos.#"
                End If
            End If

            '...si ingresa zona y/o sucursal obligo a ingresar zona, sucursal y sector
            If sError = String.Empty Then
                If (nm_zona <> String.Empty OrElse nm_sucursal <> String.Empty) AndAlso (nm_zona = String.Empty OrElse nm_sucursal = String.Empty OrElse nm_sector = String.Empty) Then
                    sError = "ERR07#Datos mínimos requeridos de la persona incompletos (zona, sucursal y/o sector).#"
                End If
            End If

            '...longitud inválida de campos
            If sError = String.Empty Then
                If nm_apellido.Length > 100 OrElse nm_nombres.Length > 100 OrElse de_mail_principal.Length > 250 Then
                    sError = "ERR08#Longitud de los campos inválida (nombre, apellido, E-mail).#"
                End If
            End If
            If sError = String.Empty Then
                If nm_zona.Length > 50 OrElse nm_sucursal.Length > 50 OrElse nm_sector.Length > 50 Then
                    sError = "ERR09#Longitud de los campos inválida (zona, sucursal, sector).#"
                End If
            End If

            '...email válido
            If sError = String.Empty Then
                If Not pers.ValidarEmail(de_mail_principal) Then
                    sError = String.Format("ERR02#La dirección de email {0} es inválida.#", de_mail_principal)
                End If
            End If

            '...mail sin duplicados (en el resto de las personas)
            If sError = String.Empty Then
                sSelect = "cd_cliente"
                sFrom = "Personas"
                sWhere = "st_estado='A'"
                sWhere &= " and de_mail_principal = " & SF.Strings.StringToSql(de_mail_principal)
                sWhere &= " and cd_persona <> " & cd_persona.ToString
                Try
                    existe = CType(Sql.Search(ConnectionString, sSelect, sFrom, sWhere), Integer)
                Catch
                    existe = 0
                End Try
                If existe <> 0 Then
                    'Verifica si existe en el mismo cliente o en otro
                    If existe = cd_cliente Then
                        'existe para el mismo cliente
                        sError = "ERR03#Email duplicado con persona existente dentro del mismo cliente#"
                    Else
                        'existe para otro cliente
                        sError = String.Format("ERR04#Email duplicado con persona existente dentro de otro cliente ({0})#", existe.ToString.Trim)
                    End If
                End If
            End If

            '...dominio
            If sError = String.Empty Then
                '...verifica si la persona está registrada en algún otro producto además de CURSOS
                sSelect = "existe = count(*)"
                sFrom = "UaPersonasNominadas pn"
                sFrom &= " join uas u on u.cd_cliente=pn.cd_cliente and u.cd_ua=pn.cd_ua"
                sFrom &= " join Productos p on p.cd_producto=u.cd_producto"
                sWhere = "pn.st_estado='A'"
                sWhere &= " and p.cd_familia<>'CURSOS'"
                sWhere &= " and pn.cd_cliente=" & cd_cliente.ToString
                sWhere &= " and pn.cd_persona=" & cd_persona.ToString

                Try
                    existe = CType(Sql.Search(ConnectionString, sSelect, sFrom, sWhere), Integer)
                Catch
                    existe = 0
                End Try

                If existe > 0 Then
                    '...valida el dominio si tiene productos de cualq. otra familia
                    Dim sDominios As String = String.Empty
                    Dim sDomMail As String
                    sDomMail = de_mail_principal.Substring(de_mail_principal.IndexOf("@")).Trim

                    Dim dsDominios As DataSet
                    dsDominios = Sql.Select(ConnectionString, "nm_dominio", "ClienteDominios", "cd_cliente = " & SF.Strings.StringToSql(cd_cliente))
                    For Each rowDominio As DataRow In dsDominios.Tables(0).Rows
                        sDominios &= Trim(rowDominio(0)) & ";"
                    Next

                    If sDominios <> String.Empty _
                    AndAlso sDominios.ToLower.IndexOf(sDomMail.ToLower) = -1 Then
                        'El dominio del mail no figura entre los dominios del cliente
                        sError = String.Format("ERR06#El dominio {0} no figura entre los dominios del cliente#", sDomMail)
                    End If
                End If
            End If

            '-----------------------------
            'PROCESO
            '-----------------------------
            If sError <> String.Empty Then
                'Reporta el error
                Throw New Exception(sError)
            Else
                'Modifica los datos de la persona localizada

                If modiPuntual Then
                    'modi de una sola persona desde el visor

                    With dsPersonas.Tables(0).Rows(0)
                        .BeginEdit()
                        .Item("nm_apellido") = nm_apellido
                        .Item("nm_nombres") = nm_nombres
                        .Item("de_mail_principal") = de_mail_principal
                        .Item("nm_sector") = nm_sector
                        .Item("nm_zona") = nm_zona
                        .Item("nm_sucursal") = nm_sucursal
                        .EndEdit()
                    End With

                    'Graba utilizando las reglas del negocio
                    pers.ActualizarDatos(dsPersonas)

                Else
                    'modi invocada desde el alta masiva de registraciones

                    With dsPersonas.Tables(0).Rows(0)
                        sFrom = "Personas"

                        sCampos = "nm_apellido, nm_nombres, de_mail_principal, nm_sector, nm_zona, nm_sucursal, fe_modi, cd_usuario_modi"

                        sValores = SF.Strings.StringToSql(nm_apellido)
                        sValores &= "# " & SF.Strings.StringToSql(nm_nombres)
                        sValores &= "# " & SF.Strings.StringToSql(de_mail_principal)
                        sValores &= "# " & SF.Strings.StringToSql(nm_sector)
                        sValores &= "# " & SF.Strings.StringToSql(nm_zona)
                        sValores &= "# " & SF.Strings.StringToSql(nm_sucursal)
                        sValores &= "# " & SF.Dates.DateToSql(Date.Now, True)
                        sValores &= "# " & SF.Strings.StringToSql("WEB")

                        sWhere = "cd_persona = " & cd_persona.ToString
                    End With

                    'Actualiza los datos de la persona
                    DA.Sql.Update(ConnectionString, sFrom, sCampos, sValores, , sWhere)

                    'Cambios
                    cambios.GrabarCambio("M", "Personas", cd_persona, "Clientes", cd_cliente, de_mail_principal)

                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            cambios = Nothing
        End Try

        'Ok
        Return True

    End Function

#End Region

End Class