Imports System.ServiceProcess
Imports System.IO
Imports Common.Env
Imports Common.MBA.Constants
Imports SF = StartFrame.SystemFunctions
Imports StartFrame
Imports System.Data.OleDb

Public Class Service1

#Region "Declaraciones"

    'Timer
    Private WithEvents temporizador As Timers.Timer
    Private intervalo As Integer = 60               'Intervalo en minutos
    Private EnEjecucion As Boolean = False          'Impide ejecución simultánea del temporizador
    Private DetenerElServicio As Boolean = False    'Flag de error grave para detener el servicio
    Private HuboErrores As Boolean = False

    'Componentes de BR
    Private sca As New StartFrame.BR.Sca
    Private Parametros As StartFrame.BR.Utilitarios.Parametros
    Private Talonarios As StartFrame.BR.Utilitarios.Talonarios

    'Log en archivo
    Private swLog As StreamWriter
    Private NombreArchivoLog As String
    Private WriteEventLog As Boolean = False     'DEBUG: Activar si es neceario

    'Parámetros del config
    Private _ServerUsr As String
    Private _DirLogs As String
    Private _MailErrores As String

    'Parametros especiales
    Private _AlertasAltas As Integer
    Private _AlertasBajas As Integer
    Private _AlertasModis As Integer
    Private _CodigoTarea As String
    Private _Entorno As String
    Private _CnnListas As String
    Private _EstadoProyecto As String = "PEN"

    'ID único que identifica la ejecución
    Private idProceso As Integer

    'Varios
    Private cantAltas As Integer = 0
    Private cantBajas As Integer = 0
    Private cantModis As Integer = 0

    'Determina el paso actual del proceso
    Private EtapaProcesamiento As EtapasProceso = EtapasProceso.NoIniciado
    Private Enum EtapasProceso
        NoIniciado = 0
        ProcesandoCierreListas = 1
        VerificandoFinCierreListas = 2
        ControlNovedades = 3
        Fin = 4
    End Enum

    'Tipos de cambio
    Private Enum TiposCambio
        Alta = 0
        Baja = 1
        Modi = 2
    End Enum

    'Constantes
    Const HeaderAppStatus As String = "Listas Service - Informe de estado: "
    Const HeaderAppAudit As String = "Listas Service - Auditoría: "
    Private HeaderAppError As String
    Const SourceApp As String = "MBA Listas"
    Const _Enter As String = Microsoft.VisualBasic.vbCr & Microsoft.VisualBasic.vbLf
    Const _Tab As String = Microsoft.VisualBasic.vbTab
    Const _Linea As String = _Enter & "------------------------------------------------------------------------------" & _Enter
    Const _LineaDoble As String = _Enter & "==============================================================================" & _Enter
    Const _LineaEnBlanco As String = _Enter & _Enter
    Private HeaderMails As String = "LISTAS - Proceso de Actualización <ERROR> - " & Now.Date.ToShortDateString

#End Region

#Region "Main"

    ' The main entry point for the process
    <MTAThread()>
    Shared Sub Main()
        'RELEASE
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New Service1}
        System.ServiceProcess.ServiceBase.Run(ServicesToRun)

        'DEBUG
        'Dim SVC As New Service1
        'Dim args1() As String = Nothing
        'SVC.OnStart(args1)
        'System.Threading.Thread.Sleep(System.Threading.Timeout.Infinite)
    End Sub

#End Region

#Region "Eventos"

    'Al iniciar el servicio
    Protected Overrides Sub OnStart(ByVal args() As String)
        WriteLogEntry(HeaderAppStatus & "Iniciando...", EventLogEntryType.Information)

        Try
            'Levanta el archivo *.config para autoconfigurarse
            Dim path, configFile As String
            path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)
            configFile = path & "\Listas.exe.config"

            If Not System.IO.File.Exists(configFile) Then
                Throw New Exception("Archivo de configuración " & configFile.Trim & " no encontrado.")
            Else
                WriteLogEntry(HeaderAppAudit & "Configurando la aplicación por medio de " & configFile.Trim, EventLogEntryType.Information)

                'Logueo al sistema
                Me._ServerUsr = Common.Env.GetConfigValue("ServerUsr")
                WriteLogEntry(HeaderAppAudit & "Server User: " & Me._ServerUsr, EventLogEntryType.Information)
                If Not Me.EsUnOperadorValido(_ServerUsr) Then
                    WriteLogEntry(HeaderAppError & "Error al loguearse a la aplicación.", EventLogEntryType.Error)
                Else
                    WriteLogEntry(HeaderAppAudit & "Logueo OK del usuario " & _ServerUsr, EventLogEntryType.SuccessAudit)

                    'Levanta los parámetros del *.config
                    Me._DirLogs = Common.Env.GetConfigValue("DirLogs")
                    Me._MailErrores = Common.Env.GetConfigValue("MailErrores")
                    Me.intervalo = CType(Common.Env.GetConfigValue("Intervalo"), Integer)

                    'Determina el nombre correcto del archivo de log para la ejecución actual del servicio
                    NombreArchivoLog = _DirLogs & "LogProceso_" _
                        & Now.Year.ToString.Trim _
                        & Microsoft.VisualBasic.Right("0" & Now.Month.ToString.Trim, 2) _
                        & Microsoft.VisualBasic.Right("0" & Now.Day.ToString.Trim, 2) _
                        & ".txt"

                    WriteLogEntry(HeaderAppAudit & "_DirLogs: " & _DirLogs, EventLogEntryType.SuccessAudit)
                    WriteLogEntry(HeaderAppAudit & "_MailErrores: " & _MailErrores, EventLogEntryType.SuccessAudit)
                    WriteLogEntry(HeaderAppAudit & "Intervalo (minutos): " & intervalo.ToString.Trim, EventLogEntryType.SuccessAudit)

                    'Parametros especiales de la tarea
                    _AlertasAltas = CInt(Common.Env.GetConfigValue("AlertasAltas"))
                    _AlertasBajas = CInt(Common.Env.GetConfigValue("AlertasBajas"))
                    _AlertasModis = CInt(Common.Env.GetConfigValue("AlertasModis"))
                    _CodigoTarea = Common.Env.GetConfigValue("CodigoTarea")
                    _Entorno = Common.Env.GetConfigValue("cnnkey")
                    HeaderAppError = "Listas Service - ERROR(" & _Entorno & ") - Proceso abortado hasta próxima ejecución"
                    _EstadoProyecto = Common.Env.GetConfigValue("EstadoProyecto")
                    _CnnListas = Common.Env.LeerStringConnection("cnnListas")   'cadena de conexión a Listas

                    SetLog("Aplicación configurada y logueada.")
                    SetLog(String.Format("Intervalo fijado en {0} minutos", intervalo.ToString.Trim))
                    SetLog(String.Format("Alarma por más de {0} altas", _AlertasAltas.ToString.Trim))
                    SetLog(String.Format("Alarma por más de {0} bajas", _AlertasBajas.ToString.Trim))
                    SetLog(String.Format("Alarma por más de {0} modificaciones", _AlertasModis.ToString.Trim))
                    SetLog(String.Format("Estado del proyecto a generar: {0}", _EstadoProyecto.ToString.Trim))

                    'Objetos de negocios 
                    Parametros = New StartFrame.BR.Utilitarios.Parametros
                    Talonarios = New StartFrame.BR.Utilitarios.Talonarios

                    'Establece un intervalo pequeño para que se inicie el evento del temporizador,
                    'después se asignará el valor configurado (en minutos)
                    temporizador = New Timers.Timer(1000)
                    temporizador.Start()
                    temporizador.Enabled = True
                End If
            End If

        Catch ex As Exception
            'Error al levantar el archivo de configuración ==> Informa y detiene el servicio
            WriteLogEntry(HeaderAppError & "Error tratando de configurar la aplicación: " & ex.Message, EventLogEntryType.Error)
            'Esto hace que se detenga el servicio
            Throw ex
        End Try
    End Sub

    'Al detener el servicio
    Protected Overrides Sub OnStop()
        SetLog("Finalizando...", , _LineaDoble)
        'Detiene el timer
        temporizador.Stop()
    End Sub

    'Al pausar el servicio
    Protected Overrides Sub OnPause()
        SetLog("En pausa...", _LineaDoble)
        'Detiene el timer
        temporizador.Stop()
    End Sub

    'Al continuar el servicio
    Protected Overrides Sub OnContinue()
        SetLog("Continuando...", _Enter, _LineaEnBlanco)
        'Reinicia el timer para su ejecución inmediata
        temporizador.Interval = 100
        temporizador.Start()
    End Sub

    'Cuando se dispara el temporizador
    Private Sub temporizador_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles temporizador.Elapsed
        'Impide ejecución simultánea de más de un tick
        If EnEjecucion Then
            Exit Sub
        End If

        ' Control horario
        Dim abortar As Boolean = False

        'Inicio
        EnEjecucion = True
        Dim inicio_proceso As New TimeSpan(DateTime.Now.Ticks)
        SetLog(Now.ToShortDateString & ", " & Now.ToLongTimeString & " - En ejecución...", _Linea)

        Try
            'Al iniciar, establece el intervalo predefinido para las siguientes ejecuciones
            temporizador.Interval = intervalo * 60 * 1000

            'Determina el nombre correcto del archivo de log para la ejecución actual del servicio
            NombreArchivoLog = _DirLogs & "LogProceso_" _
                & Now.Year.ToString.Trim _
                & Microsoft.VisualBasic.Right("0" & Now.Month.ToString.Trim, 2) _
                & Microsoft.VisualBasic.Right("0" & Now.Day.ToString.Trim, 2) _
                & ".txt"

            Try
                'Controla que no corra sábados, domingos o feriados
                If Not Common.MBA.MetodosComunes.EsDiaHabil(Now) Then
                    'Aborta la ejecución hasta el día siguiente a la hora de inicio del servicio (para volver a controlar)
                    Dim diaReinicio As Date = DateAdd(DateInterval.Day, 1, Now())
                    Dim fechaHoraReinicio As New DateTime(diaReinicio.Year, diaReinicio.Month, diaReinicio.Day, CInt(Common.Env.GetConfigValue("HoraInicio").Substring(0, 2)), 0, 0)
                    Dim minutosHastaReinicio As Integer = DateDiff(DateInterval.Minute, Now(), fechaHoraReinicio)
                    If minutosHastaReinicio > 6 * 60 Then
                        minutosHastaReinicio = 6 * 60   'Máximo 6 hs para volver a controlar
                    End If
                    temporizador.Interval = minutosHastaReinicio * 60 * 1000

                    abortar = True
                    EnEjecucion = False
                    SetLog(String.Format("{0}No corresponde ejecutar en sábado/domingo/feriado.{4}       Fecha-Hora ejecución actual: {1}. Fecha-Hora estimada para reinicio: {2}. Minutos hasta el reinicio (máx. 6 hs): {3}",
                                             HeaderAppStatus, Now(), fechaHoraReinicio, minutosHastaReinicio, _Enter), EventLogEntryType.Warning, _Linea)
                    Exit Try
                Else
                    ' Control ejecución fuera de horario
                    Dim _hora As String = DateTime.Now.ToString("HH:mm")
                    If _hora < Common.Env.GetConfigValue("HoraInicio") OrElse _hora > Common.Env.GetConfigValue("HoraFinal") Then
                        Dim diaReinicio As Date = Now()
                        Dim fechaHoraReinicio As New DateTime(diaReinicio.Year, diaReinicio.Month, diaReinicio.Day, CInt(Common.Env.GetConfigValue("HoraInicio").Substring(0, 2)), 0, 0)
                        Dim minutosHastaReinicio As Integer = DateDiff(DateInterval.Minute, Now(), fechaHoraReinicio)
                        If minutosHastaReinicio > 6 * 60 Then
                            minutosHastaReinicio = 6 * 60   'Máximo 6 hs para volver a controlar
                        End If
                        If minutosHastaReinicio <= 0 Then
                            minutosHastaReinicio = 5
                        End If
                        temporizador.Interval = minutosHastaReinicio * 60 * 1000

                        abortar = True
                        EnEjecucion = False
                        SetLog(String.Format("{0}No corresponde ejecutar fuera de horario.{4}       Fecha-Hora ejecución actual: {1}. Fecha-Hora estimada para reinicio: {2}. Minutos hasta el reinicio (máx. 6 hs): {3}",
                                             HeaderAppStatus, Now(), fechaHoraReinicio, minutosHastaReinicio, _Enter), EventLogEntryType.Warning, _Linea)
                        Exit Try
                    End If
                End If

                'Ejecuta el proceso principal
                HuboErrores = False
                Me.ProcesoPrincipal()

                If HuboErrores Then
                    Throw New Exception("El log del mailer fue procesado, pero CON ERRORES. Revise el LOG para determinar las acciones a seguir.")
                End If

            Catch ex As Exception
                'Loguea el error
                SetLog("Error procesando la información:" & ex.Message, HeaderAppError, _Linea)
                'Envía el mail informando
                SendMail(Me._MailErrores, HeaderAppError, "Se produjo un error en la ejecución de la aplicación de referencia. " _
                & "Ver detalles en el log " & NombreArchivoLog & ". ERROR: " & ex.Message)
            End Try

            If abortar Then
                Exit Try
            End If

            'Aquí pueden ejecutarse llamados a procesos secundarios...




        Catch ex As Exception
            'Loguea el error
            SetLog("Error procesando la información:" & ex.Message, HeaderAppError, _Linea)
            'Envía el mail informando
            SendMail(Me._MailErrores, HeaderAppError, "Se produjo un error en la ejecución de la aplicación de referencia. " _
                & "Ver detalles en el log " & NombreArchivoLog & ". ERROR: " & ex.Message)

        Finally
            'Permite la ejecución de nuevos ticks
            EnEjecucion = False
            'Fin
            SetLog("Fin ejecución.", _LineaEnBlanco, _LineaDoble)
            If DetenerElServicio Then
                SetLog("SERVICIO DETENIDO POR ERROR GRAVE", _LineaEnBlanco, _LineaDoble)
                Me.Stop()
            End If
            'Cierra el archivo del log temporal
            If Not swLog Is Nothing Then
                swLog.Close()
                swLog = Nothing
            End If
        End Try
    End Sub

#End Region

#Region "Métodos"

#Region "Generales"

    'Verifica el logueo del operador
    Private Function EsUnOperadorValido(ByVal nombreOperador As String) As Boolean

        'Verifica si existe el operador dentro del sistema
        Dim sca As New StartFrame.BR.Sca
        Dim claveAcceso As String
        Try
            If sca.ExisteOperador(nombreOperador) Then
                'Obtiene el password y lo desencripta
                claveAcceso = CType(DA.Sql.Search(ConnectionString, "va_clave_acceso", "wad_operadores", "cd_operador = '" & nombreOperador & "'"), String)
                If claveAcceso <> "" Then
                    claveAcceso = SF.Strings.EncryptString(claveAcceso, KEY, SF.Strings.Accion.DECRYPT)
                End If

                'Registra la terminal si no existe
                sca.SetPrefTerm(Terminal, nombreOperador, 1, "", 0, 0, 0)

                'Verifica si es un usuario válido
                If sca.Login(nombreOperador, claveAcceso) Then
                    'Verifica si debe cambiar su clave de acceso
                    If sca.RequiereCambioClave(nombreOperador) Then
                        Throw New Exception("ERROR: El operador deberá cambiar su clave de acceso.")
                    End If
                    'Propiedades de la Common
                    Operador = nombreOperador
                    Password = claveAcceso
                    'Fin ok
                    Return True
                End If
            End If

            'Fin mal
            Return False

        Catch ex As Exception
            Throw ex

        Finally
            sca = Nothing
        End Try

    End Function

    'Envía un mail vía SMTP
    Public Shared Sub SendMail(ByVal mail_to As String,
                ByVal mail_subject As String,
                ByVal mail_body As String,
                Optional ByVal param_smtp As String = "MAILS_SMTP",
                Optional ByVal param_from As String = "MAILS_FROM",
                Optional ByVal param_pwd As String = "MAILS_PWD",
                Optional ByVal formatoHtml As Boolean = False)

        Try
            'NOTE: Para que este método funcione, se requiere tener registrada MSMAPI32.OCX

            'Busca parámetros
            Dim _param As New StartFrame.BR.Utilitarios.Parametros

            Dim smtpName As String
            Dim mail_from As String
            Dim mail_pwd As String
            smtpName = CType(_param.getParametro(param_smtp), String)
            mail_from = CType(_param.getParametro(param_from), String)
            mail_pwd = CType(_param.getParametro(param_pwd), String)

            'Arma el mail
            Dim Message As System.Net.Mail.MailMessage
            Message = New System.Net.Mail.MailMessage(mail_from, mail_to, mail_subject, mail_body)
            Message.IsBodyHtml = formatoHtml

            'Envía el mail
            Dim smtp As New System.Net.Mail.SmtpClient(smtpName)
            smtp.Send(Message)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    'Escribe en el log de Windows
    Private Sub WriteLogEntry(texto As String, TipoLog As EventLogEntryType)
        If WriteEventLog Then
            WriteLogEntry(texto, TipoLog)
        End If
    End Sub

    'Agrega un texto al log en disco y en el eventlog de aplicaciones
    Private Sub SetLog(ByVal texto As String,
                    Optional ByVal AntesDelTexto As String = "",
                    Optional ByVal DespuesDelTexto As String = _Enter,
                    Optional ByVal tipoLog As EventLogEntryType = EventLogEntryType.SuccessAudit)

        'Informa en el visor de sucesos
        'WriteLogEntry(texto, tipoLog)

        'Crea el log en disco si no existe
        If swLog Is Nothing OrElse Not File.Exists(NombreArchivoLog) Then
            Try
                'Si ya existe, cierra el log anterior
                If Not swLog Is Nothing Then
                    swLog.Close()
                    swLog = Nothing
                End If

                'Abre el archivo de log: si no existe, lo crea. Si existe: agrega texto al final
                swLog = New StreamWriter(NombreArchivoLog, True)

                'Encabezado del archivo
                swLog.WriteLine(HeaderAppStatus)
                swLog.WriteLine("Creación del Log: " & Now.ToShortDateString & " - " & Now.ToShortTimeString)
                swLog.WriteLine("Versión: " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString)

            Catch ex As Exception
                'No genera error de este tipo, simplemente no deja log
                WriteLogEntry(ex.Message, EventLogEntryType.FailureAudit)
            End Try
        End If

        'Agrega la línea en el log en disco
        swLog.Write(AntesDelTexto)
        swLog.Write(texto)
        swLog.Write(DespuesDelTexto)

    End Sub

    'Inicializa una nueva conexión
    Private Function InicializarConexion(Optional sConexion As String = "") As OleDbConnection
        Try
            'Conexión
            If sConexion = "" Then
                sConexion = Common.Env.ConnectionString
            End If
            'Crea una nueva conexión con el motor de datos
            Dim newConnection As OleDbConnection = New OleDbConnection(sConexion)

            Return newConnection

        Catch errorException As Exception
            'Control de errores
            DA.Env.LogError("WRSQL", Terminal, Operador, "DB009", "[NewConnection] " & errorException.Message)
            Throw New Exception(errorException.Message, errorException)
        End Try
    End Function

#End Region

#Region "Específicos"

    'Ejecución del proceso principal cada vez que se dispara el temporizador.
    'No se requiere Try-Catch. Ante un error se recomienda propagar la excepción.
    Private Sub ProcesoPrincipal()
        HuboErrores = False
        DetenerElServicio = False

        'Genera un nuevo ID de Proceso
        idProceso = Talonarios.getNum("LISTASEXP", True)
        SetLog(String.Format("ID del Proceso Actual: {0}", idProceso.ToString.Trim))

        Dim cn As OleDbConnection = Nothing
        Dim tran As OleDbTransaction = Nothing
        Dim param As New ArrayList
        Dim ttr As Integer = 0

        Try
            'Inicializa la conexión e inicia una transacción
            cn = InicializarConexion()
            cn.Open()

            '===================================================================
            'PASO 1: Ejecuta el cierre de Listas
            '===================================================================
            EtapaProcesamiento = EtapasProceso.ProcesandoCierreListas
            SetLog(String.Format("{1} Iniciando paso '{0}':", EtapaProcesamiento.ToString.Trim, Now().ToLocalTime), _Enter)

            'Ejecuta el cierre
            param.Clear()
            param.Add(idProceso)
            DA.Sql.ExecSP(_CnnListas, "Listas_Cierre", param, 3 * 60 * 1000) 'timeout de 10 minutos

            '===================================================================
            'PASO 2: Verifica el fin del cierre de listas
            '===================================================================
            EtapaProcesamiento = EtapasProceso.VerificandoFinCierreListas
            SetLog(String.Format("{1} Iniciando paso '{0}':", EtapaProcesamiento.ToString.Trim, Now().ToLocalTime), _Enter)

            'Verifica el fin del cierre de listas
            Dim rdoProceso As Integer = VerificarFinProceso("CIERRE.LISTAS")
            If rdoProceso = 10 Then
                'Proceso finalizado => Siguiente paso
            ElseIf rdoProceso = 7 Then
                'Finalizó con errores => avisa y aborta
                Throw New Exception("El proceso de Cierre de Listas finalizó informando el paso 7. Proceso abortado.")
            Else
                'No finalizado aún => error
                Throw New Exception("El proceso de Cierre de Listas no finalizó correctamente. Debe informar paso 7 o 10. Proceso abortado.")
            End If

            '===================================================================
            'PASO 3: Novedades a publicar
            '===================================================================
            EtapaProcesamiento = EtapasProceso.ControlNovedades
            SetLog(String.Format("{1} Iniciando paso '{0}':", EtapaProcesamiento.ToString.Trim, Now().ToLocalTime), _Enter)

            'Publica las novedades detectadas
            ControlNovedadesAPublicar()

            '===================================================================
            'PAOS 4: Marca la fecha de la última corrida
            '===================================================================
            EtapaProcesamiento = EtapasProceso.Fin
            SetLog(String.Format("{1} Iniciando paso '{0}':", EtapaProcesamiento.ToString.Trim, Now().ToLocalTime), _Enter)

            'Actualiza TareasStatus
            Dim FechaHoy As String = String.Empty
            FechaHoy = Now.Year.ToString.Trim.PadLeft(4, "0") & "/"
            FechaHoy &= Now.Month.ToString.Trim.PadLeft(2, "0") & "/"
            FechaHoy &= Now.Day.ToString.Trim.PadLeft(2, "0")

            DA.Sql.Update(ConnectionString, "TareasStatus", "fe_ultima_corrida", "'" & FechaHoy & "'", Nothing, "cd_tarea = " & Me._CodigoTarea)
            DA.Sql.Update(ConnectionString, "TareasStatus", "fe_ult_publicacion", "getdate()", Nothing, "cd_tarea = " & Me._CodigoTarea)

            SetLog("Actualizada la fecha de última corrida ( " & FechaHoy & " )")

            'Manda mail informando
            Dim ssubject As String
            If _EstadoProyecto = "PEN" Then
                ssubject = "LISTAS CIERRE (" & _Entorno & ") - Terminó (falta publicar proyectos y cambios)"
                SendMail(_MailErrores, ssubject, "Finalizó OK el procesamiento de LISTAS del " & Now.Date.ToShortDateString.Trim & ". Ver el log (" & _DirLogs & ") para mayores detalles.")
            Else
                ssubject = "LISTAS CIERRE (" & _Entorno & ")- Terminó OK"
            End If

        Catch ex As Exception
            'Graba el error en el log y cancela el proceso
            SetLog(ex.Message, HeaderAppError)
            HuboErrores = True
            'DetenerElServicio = True 'No detiene el servicio por errores en este proceso. Únicamente informa.
            Throw ex

        Finally
            'Cierra la conexión
            tran = Nothing
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
            cn = Nothing
        End Try

    End Sub

#End Region

#Region "Procesamiento de Listas"

    'Publicación de novedades
    Private Sub ControlNovedadesAPublicar()
        Dim hayNovedades As Boolean = False

        'Objetos de negocios
        Dim _Proyectos As New MBA.Manager.BR.Proyectos(Operador)
        Dim _Cambios As New MBA.Manager.BR.Cambios(Operador)

        'Busca las fuentes que deben ser incluidas (excluirá al resto)
        Dim fuentesIncluidas As New ArrayList
        fuentesIncluidas.Clear()
        Dim ds As DataSet
        Dim sCampos, sFrom, sOrder As String

        sCampos = "fuente"
        sFrom = " listas_descripciones"
        sOrder = " fuente"

        ds = DA.Sql.Select(ConnectionString, sCampos, sFrom, , sOrder)

        '... las agrega a un arraylist para luego validar
        For Each row As DataRow In ds.Tables(0).Rows
            fuentesIncluidas.Add(row.Item(0).ToString.Trim)
        Next

        'Procesa las bajas
        SetLog("...Procesando novedades: BAJAS", HeaderAppStatus)
        If GenerarCambios(TiposCambio.Baja, fuentesIncluidas) Then
            hayNovedades = True
        End If

        'Procesa las altas
        SetLog("...Procesando novedades: ALTAS", HeaderAppStatus)
        If GenerarCambios(TiposCambio.Alta, fuentesIncluidas) Then
            hayNovedades = True
        End If

        'Procesa las modificaciones
        SetLog("...Procesando novedades: MODIFICACIONES", HeaderAppStatus)
        If GenerarCambios(TiposCambio.Modi, fuentesIncluidas) Then
            hayNovedades = True
        End If

        'Proyectos y Cambios
        If hayNovedades Then
            SetLog("...Generando proyectos y cambios", HeaderAppStatus)
            _Cambios.GrabarCambio("M", NombresTablas.Listas, 1, , , "LISTAS")
            _Proyectos.AgregarProyectoPadre(NombresTablas.Listas, 1, TiposDeProyectos.MAE,
                                                    "Listas", "Actualización Listas", _EstadoProyecto)
            SetLog("...OK", HeaderAppStatus)
        End If

        'Verifica alarmas
        If cantAltas > _AlertasAltas Or cantBajas > _AlertasBajas Or cantModis > _AlertasModis Then
            Dim msgAlerta As String = ""
            msgAlerta &= "En el día de la fecha se detectaron:"
            If cantAltas > _AlertasAltas Then
                msgAlerta &= String.Format("Altas = {0} cant. de reg. (siendo {1} la cant. alarmada para las altas){2}",
                                           cantAltas.ToString.Trim, _AlertasAltas.ToString.Trim, vbCrLf)
            End If
            If cantBajas > _AlertasBajas Then
                msgAlerta &= String.Format("Bajas = {0} cant. de reg. (siendo {1} la cant. alarmada para las bajas){2}",
                                           cantBajas.ToString.Trim, _AlertasBajas.ToString.Trim, vbCrLf)
            End If
            If cantModis > _AlertasModis Then
                msgAlerta &= String.Format("Modis = {0} cant. de reg. (siendo {1} la cant. alarmada para las modis){2}",
                                           cantModis.ToString.Trim, _AlertasModis.ToString.Trim, vbCrLf)
            End If
            msgAlerta &= vbCrLf & "Gracias" & vbCrLf

            SendMail(_MailErrores, "LISTAS ALARMA(" & _Entorno & ") - Proceso de actualización con datos excesivos", msgAlerta)
        End If

        'Fin
        _Proyectos = Nothing
        _Cambios = Nothing
    End Sub

    'Verifica la fecha del archivo solicitado y retorna TRUE si es del día
    Private Function VerificarUltimaCorrida() As Boolean
        Dim ok As Boolean = False

        'Verifica la fecha del último procesamiento: debe ser la de ayer
        Dim dFechaAyer As Date = Now.Date.AddDays(-1)
        Dim sFechaAyer As String = String.Empty
        Dim rdo As Integer = 0
        Dim sCond As String

        sFechaAyer = dFechaAyer.Year.ToString.Trim.PadLeft(4, "0") & "/"
        sFechaAyer &= dFechaAyer.Month.ToString.Trim.PadLeft(2, "0") & "/"
        sFechaAyer &= dFechaAyer.Day.ToString.Trim.PadLeft(2, "0")

        sCond = "fe_ultima_corrida = '" & sFechaAyer & "' and cd_tarea = " & Me._CodigoTarea

        rdo = SF.Number.IsNumNull(DA.Sql.Search(ConnectionString, "count(*)", "TareasStatus", sCond), 0)

        If rdo = 1 Then
            ok = True 'la fecha de la última corrida exitosa fue la de ayer
        End If

        'Archivo adecuado para procesar o no
        Return ok

    End Function

    'Verifica el fin del proceso
    Private Function VerificarFinProceso(proceso As String) As Integer
        'rdo = 0 (no terminó el proceso)
        'rdo = 7 (terminó con errores)
        'rdo = 10 (terminó ok)
        Dim rdo As Integer = 0
        Dim sSelect, sFrom, sWhere As String
        Dim sError As String

        SetLog(String.Format("...Verificando fin del proceso '{0}", proceso), HeaderAppStatus)

        Select Case proceso
            Case "CIERRE.LISTAS"
                'Busca el log del fin del proceso de cierre de Listas
                sSelect = "nu_paso"
                sFrom = "Listas_Exportacion"
                sWhere = "cd_operacion = " & idProceso.ToString
                sWhere &= " and (nu_paso = 7 or nu_paso = 10)"

                Try
                    rdo = SF.Number.IsNumNull(DA.Sql.Search(_CnnListas, sSelect, sFrom, sWhere))
                Catch ex As Exception
                    rdo = 0
                End Try
        End Select

        If rdo = 0 Then
            'Proceso en curso...
            SetLog("...No terminado aún", HeaderAppStatus)
        ElseIf rdo = 7 Then
            'backup terminado => fin
            sSelect = "de_operacion"
            sFrom = "Listas_Exportacion"
            sWhere = "nu_paso = 7 and cd_operacion = " & idProceso.ToString
            sError = DA.Sql.Search(_CnnListas, sSelect, sFrom, sWhere)

            SetLog("...Finalizado con errores (paso 7) - " & sError.Trim, HeaderAppStatus)
        ElseIf rdo = 10 Then
            SetLog("...Finalizado OK (paso 10)", HeaderAppStatus)
        End If

        Return rdo

    End Function

    'Levanta el archivo de novedades de altas, bajas o modis y genera los cambios necesarios
    Private Function GenerarCambios(tipoCambio As TiposCambio, fuentesIncluidas As ArrayList) As Boolean
        Dim tabla As Integer
        Dim hayNovedades As Boolean = False
        Dim cantRegistros As Integer = 0
        Dim cantIgnorados As Integer = 0

        Try
            'Verifica las novedades a nivel de BD
            Dim dsDiferencias As DataSet
            dsDiferencias = DA.Sql.ExecSPDS(ConnectionString, "Listas_controlWS")
            If Not SF.Files.IsEmpty(dsDiferencias) Then

                'Determina el tipo de novedad a procesar
                Select Case tipoCambio
                    Case TiposCambio.Alta
                        tabla = 0 'LS-noWS
                    Case TiposCambio.Baja
                        tabla = 1 'WS-noLS
                    Case TiposCambio.Modi
                        tabla = 2 'WS-modi
                End Select

            End If

            Dim fuente, sdn, alt As String
            Dim param As New ArrayList

            'Recorre el archivo
            For Each row As DataRow In dsDiferencias.Tables(tabla).Rows
                fuente = "" : sdn = "" : alt = ""

                Try
                    'Busca la clave
                    fuente = row("FUENTE").ToString.Trim
                    sdn = row("SDN_ID")
                    alt = row("ALT_ID")

                    'Verifica si debe publicar esta novedad: 
                    If fuentesIncluidas.IndexOf(fuente) = -1 Then
                        'Ignora la novedad
                        cantIgnorados += 1
                        SetLog(String.Format("Fuente ignorada: '{0}'", fuente), HeaderAppAudit)
                    Else
                        'Actualiza los datos en SQL
                        param.Clear()
                        param.Add(tipoCambio.ToString.Substring(0, 1))
                        param.Add(fuente)
                        param.Add(sdn)
                        param.Add(alt)

                        DA.Sql.ExecSP(ConnectionString, "Listas_ABM", param)
                        hayNovedades = True
                        cantRegistros += 1
                    End If

                Catch ex As Exception
                    'Ignora el registro y procesa el siguiente
                    SetLog(String.Format("Error procesando '{0}' con el siguiente registro: {1} {2} {3}",
                                         tipoCambio.ToString.Trim, fuente, sdn, alt), HeaderAppError)
                End Try

            Next

            'Controles de cant. de registros
            Select Case tipoCambio
                Case TiposCambio.Alta
                    cantAltas = cantRegistros
                Case TiposCambio.Baja
                    cantBajas = cantRegistros
                Case TiposCambio.Modi
                    cantModis = cantRegistros
            End Select

            'Informa
            If cantRegistros > 0 Then
                SetLog("......" & cantRegistros.ToString.Trim & " registros procesados.", HeaderAppStatus)
            Else
                SetLog("......Sin novedad", HeaderAppStatus)
            End If
            If cantIgnorados > 0 Then
                SetLog("......" & cantIgnorados.ToString.Trim & " registros ignorados (fuente no aplicable).", HeaderAppStatus)
            End If

        Catch ex As Exception
            'Propaga el error, lo que detendrá el proceso
            Throw ex
        End Try

        'Indica si proceso algún cambio
        Return hayNovedades

    End Function

#End Region

#End Region

End Class
