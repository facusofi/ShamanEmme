Imports ShamanExpressDLL
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports NLog

Public Class frmLog

    Private Const cnnDynamic As String = "Dynamics"
    Private flgGuadalajara As Boolean = False
    Private vFileLog As String

    Enum modProcess
        All = 0
        NewClients = 1
        AllActives = 2
        NewActives = 3
        AllInactives = 4
        NewInactives = 5
        LastUpdates = 6
        Coded = 9
        Others = 10
    End Enum

    Private Sub frmLog_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub

    Private Sub frmLog_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.tmrInit.Enabled = True
    End Sub

    Private Sub tmrInit_Tick(sender As Object, e As EventArgs) Handles tmrInit.Tick
        Me.tmrInit.Enabled = False
        Me.Syncronize()
    End Sub

    Private Sub AddLog(pVal As String, Optional pCls As Boolean = False)
        Try
            If pCls Then
                Me.txtLog.Text = pVal
            Else
                Me.txtLog.Text = Me.txtLog.Text & vbCrLf & pVal
            End If

            If Not File.Exists(vFileLog) Then

                Using sw As StreamWriter = File.CreateText(vFileLog)
                    sw.WriteLine(pVal)
                End Using

            Else

                Using sw As StreamWriter = File.AppendText(vFileLog)
                    sw.WriteLine(pVal)
                End Using

            End If

        Catch ex As Exception
            HandleError(Me.Name, "AddLog", ex)
        End Try
    End Sub

    Private Sub Syncronize()
        Try

            '--------> Log
            vFileLog = My.Application.Info.DirectoryPath & "\Log " & DateToAnsi(Now.Date) & " " & Format(Now.Hour, "00") & ".log"

            If File.Exists(vFileLog) Then
                File.Delete(vFileLog)
            End If

            If cnnsNET.Count = 0 Then
                AddLog("Licencia vencida o ilegible " & Now, True)
                Me.Close()
            End If

            AddLog("Iniciando Sincronización " & Now, True)

            '-------> Reconecto al server en otra DB
            If Not shamanStartUp.AbrirConexion(cnnDynamic) Then
                AddLog("No se pudo instanciar una nueva conexión")
                Exit Sub
            End If

            '-------> Cambio el database
            Dim vCnnStr As String = ConfigurationManager.AppSettings.Get("dbDynamics")
            AddLog("Conectando a " & vCnnStr)
            cnnsNET(cnnDynamic).Close()
            cnnsNET(cnnDynamic).ConnectionString = vCnnStr
            cnnsNET(cnnDynamic).Open()
            AddLog("Conexión exitosa " & cnnsNET(cnnDynamic).Database)

            Dim myProcess As modProcess = CType(ConfigurationManager.AppSettings.Get("modProcess"), modProcess)
            flgGuadalajara = setIntToBool(Val(ConfigurationManager.AppSettings.Get("modGuadalajara")))

            Dim selCliente As String = ConfigurationManager.AppSettings.Get("modCliente")

            '-------> Inicio proceso...
            If selCliente = "" And myProcess <> modProcess.Others Then
                Me.SyncConvenios()
            End If

            If Not flgGuadalajara Then

                If myProcess = modProcess.Coded Then
                    Me.SyncClientes(modProcess.NewActives, selCliente)
                    Me.SyncClientes(modProcess.NewInactives, selCliente)
                    Me.SyncClientes(modProcess.LastUpdates, selCliente)
                    'Me.SyncClientes(modProcess.AllInactives, selCliente)
                Else
                    Me.SyncClientes(myProcess, selCliente)
                End If

            Else
                If myProcess = modProcess.Coded Then
                    'Me.SyncClientes(modProcess.All, selCliente)
                    Me.SyncClientes(modProcess.LastUpdates, selCliente)
                Else
                    Me.SyncClientes(myProcess, selCliente)
                End If
            End If

            '-------> Final
            Me.Close()

        Catch ex As Exception

            AddLog(ex.Message)
            Me.Close()

        End Try
    End Sub

    Private Sub SyncConvenios()

        Dim vStep As String = "Inicio"

        Try

            Dim SQL As String
            Dim objClientes As New conClientes
            Dim objSituacionesIva As New conSituacionesIva
            Dim objRubros As New conRubrosClientes
            Dim objEmpresas As New conEmpresasLegales


            Dim vEmpId As Int64 = objEmpresas.GetDefault()

            AddLog("Sincronizando convenios")

            SQL = "SELECT cnv.*, "
            SQL = SQL & "prv.STATE AS ProvinciaCodigo, prv.DSCRIPTN AS ProvinciaDescripcion, "
            SQL = SQL & "loc.Codigo_Ciudad AS LocalidadCodigo, loc.Descripcion_Ciudad AS LocalidadDescripcion "

            If Not flgGuadalajara Then
                SQL = SQL & "FROM EMP01004 (nolock) cnv "
                SQL = SQL & "LEFT JOIN EMP01007 (nolock) loc ON cnv.Codigo_Ciudad = loc.Codigo_Ciudad "
                SQL = SQL & "LEFT JOIN EMP41004 (nolock) prv ON cnv.STATE = prv.STATE "
                SQL = SQL & "LEFT JOIN [ShamanDev].[dbo].[Clientes] shm ON cnv.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS AS varbinary(12)) "
                SQL = SQL & "WHERE shm.AbreviaturaId IS NULL "
                SQL = SQL & "ORDER BY cnv.CUSTNMBR"
            Else
                SQL = SQL & "FROM EMP00103 cnv "
                SQL = SQL & "LEFT JOIN EMP00306 loc ON cnv.Codigo_Ciudad = loc.Codigo_Ciudad "
                SQL = SQL & "LEFT JOIN EMP00302 prv ON cnv.STATE = prv.STATE "
                SQL = SQL & "ORDER BY cnv.CUSTNMBR"
            End If

            Dim cmdDyn As New SqlCommand(SQL, cnnsNET(cnnDynamic), cnnsTransNET(cnnDynamic))
            cmdDyn.CommandTimeout = 600
            Dim dt As New DataTable
            dt.Load(cmdDyn.ExecuteReader)
            Dim vIdx As Integer

            For vIdx = 0 To dt.Rows.Count - 1

                Dim vAddNew As Boolean = False

                If Not objClientes.Abrir(objClientes.GetIDByAbreviaturaId(dt(vIdx)("CUSTNMBR").ToString.Trim())) Then

                    objClientes.CleanProperties(objClientes)
                    vAddNew = True

                End If

                vStep = "Cabecera"

                objClientes.AbreviaturaId = dt(vIdx)("CUSTNMBR").ToString.Trim()

                vStep = "Razón Social"

                objClientes.RazonSocial = Microsoft.VisualBasic.Left(dt(vIdx)("CUSTNAME").ToString.Trim(), 100)

                vStep = "Rubro"

                objClientes.RubroClienteId.SetObjectId(Me.getRubroId("CNV", "CABECERA DE CONVENIO"))

                vStep = "Domicilio"

                If Not IsDBNull(dt(vIdx)("Direccion")) Then objClientes.Domicilio.dmCalle = Microsoft.VisualBasic.Left(dt(vIdx)("Direccion").ToString.Trim(), 200)

                vStep = "Localidad"

                If Not IsDBNull(dt(vIdx)("LocalidadCodigo")) AndAlso Not IsDBNull(dt(vIdx)("ProvinciaCodigo")) Then
                    objClientes.LocalidadId.SetObjectId(Me.getLocalidadId(dt(vIdx)("LocalidadCodigo"), dt(vIdx)("LocalidadDescripcion"), dt(vIdx)("ProvinciaCodigo"), dt(vIdx)("ProvinciaDescripcion")))
                End If

                vStep = "CPO"

                If Not IsDBNull(dt(vIdx)("Codigo_Postal")) Then objClientes.CodigoPostal = Microsoft.VisualBasic.Left(dt(vIdx)("Codigo_Postal").ToString.Trim(), 10)

                vStep = "Fechas de Alta y Baja"

                objClientes.FecEgreso = Now.Date
                objClientes.FecIngreso = Now.Date.AddDays(-1)

                vStep = "Tipo de Convenio"

                objClientes.TipoConvenio = 2

                vStep = "Otros"

                objClientes.EmpresaLegalId.SetObjectId(vEmpId)

                vStep = "Salvando Cliente"

                If objClientes.Salvar(objClientes, , False) Then

                    If vAddNew Then
                        AddLog("Se creó el convenio " & dt(vIdx)("CUSTNMBR"))
                    Else
                        AddLog("Se actualizó el convenio " & dt(vIdx)("CUSTNMBR"))
                    End If

                Else

                    AddLog("Error al salvar convenio " & dt(vIdx)("CUSTNMBR") & ": " & objClientes.MyLastExec.ErrorDescription)

                End If

            Next vIdx

            objClientes = Nothing
            objSituacionesIva = Nothing
            objRubros = Nothing
            objEmpresas = Nothing

        Catch ex As Exception
            AddLog(vStep & ": " & ex.Message)
        End Try
    End Sub

    Private Sub SyncClientes(Optional pMod As modProcess = modProcess.All, Optional pSelCli As String = "")

        Dim vStep As String = "Inicio"

        Try

            Dim SQL As String
            Dim objClientes As New conClientes
            Dim objSituacionesIva As New conSituacionesIva
            Dim objRubros As New conRubrosClientes
            Dim objEmpresas As New conEmpresasLegales
            Dim objCobradores As New conCobradores
            Dim objFormasPago As New conFormasPago
            Dim objClasif As New conIntegrantesClasificaciones

            Dim vEmpId As Int64 = objEmpresas.GetDefault()
            Dim dtCob As DataTable = objCobradores.GetAll()
            Dim vCobId As Int64 = 0
            Dim vFpg As Int64 = objFormasPago.GetDefaultCobranza()
            Dim vIvaCF As Int64 = objSituacionesIva.GetIDByAbreviaturaId("CF")
            Dim vIvaEX As Int64 = objSituacionesIva.GetIDByAbreviaturaId("EX")
            Dim vClfId As Int64 = objClasif.GetIDByAbreviaturaId("AF")

            If dtCob.Rows.Count > 0 Then vCobId = dtCob(0)("ID")

            AddLog("")
            AddLog("Sincronizando clientes (Proceso " & pMod.ToString & ")")

            SQL = "Select cli.CUSTNMBR, cli.STMTNAME, cli.ADDRESS1, cli.ADDRESS2, cli.ADDRESS3, cli.ZIP, cli.TAXSCHID, adi.FECHSUSP, adi.FECEMCNT, cli.HOLD, adi.CONVENIO, "
            SQL = SQL & "cli.PHONE1, cli.PHONE2, cnv.CUSTNAME As NombreConvenio, "
            SQL = SQL & "prv.STATE As ProvinciaCodigo, prv.DSCRIPTN As ProvinciaDescripcion, "
            SQL = SQL & "loc.Codigo_Ciudad As LocalidadCodigo, loc.Descripcion_Ciudad As LocalidadDescripcion, "
            SQL = SQL & "pla.CodigoConvenio As PlanCodigo, pla.DescripcionConvenio As PlanDescripcion, "
            SQL = SQL & "rub.CLASSID As RubroCodigo, rub.CLASDSCR As RubroDescripcion "

            If Not flgGuadalajara Then

                SQL = SQL & "FROM RM00101 (nolock) cli "
                SQL = SQL & "LEFT JOIN EMP01002 (nolock) adi On cli.CUSTNMBR = adi.CUSTNMBR "
                SQL = SQL & "LEFT JOIN EMP01004 (nolock) cnv On adi.CONVENIO = cnv.CUSTNMBR "
                SQL = SQL & "LEFT JOIN EMP01007 (nolock) loc On cli.CITY = loc.Codigo_Ciudad "
                SQL = SQL & "LEFT JOIN EMP41001 (nolock) pla On cnv.CodigoConvenio = pla.CodigoConvenio "
                SQL = SQL & "LEFT JOIN EMP41004 (nolock) prv On cli.STATE = prv.STATE "
                SQL = SQL & "LEFT JOIN RM00201 (nolock) rub On cli.CUSTCLAS = rub.CLASSID "

                Select Case pMod
                    Case modProcess.AllActives
                        SQL = SQL & "LEFT JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) "
                        SQL = SQL & "WHERE (cli.HOLD = 0) "
                    Case modProcess.NewActives
                        SQL = SQL & "LEFT JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) And shm.virActivo = 1 "
                        SQL = SQL & "WHERE (cli.HOLD = 0) And (shm.AbreviaturaId Is NULL) "
                    Case modProcess.AllInactives
                        SQL = SQL & "LEFT JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) "
                        SQL = SQL & "WHERE (cli.HOLD = 1) And (shm.AbreviaturaId Is NULL) "
                    Case modProcess.NewInactives
                        SQL = SQL & "LEFT JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) And shm.virActivo = 1 "
                        SQL = SQL & "WHERE (cli.HOLD = 1) And (shm.AbreviaturaId Is Not NULL) "
                    Case modProcess.LastUpdates
                        SQL = SQL & "INNER JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) "
                        SQL = SQL & "WHERE ((cli.MODIFDT >= '" & DateToSql(Now.Date,, True, True) & "') "
                        SQL = SQL & "OR (cli.CUSTNMBR IN(SELECT IDGRUPO FROM EMP01001 WHERE (MODIFDT >= '" & DateToSql(Now.Date,, True, True) & "') OR (Fecha_Alta >= '" & DateToSql(Now.Date,, True, True) & "')))) "
                    Case modProcess.Others
                        SQL = SQL & "INNER JOIN [ShamanDev].[dbo].[Clientes] (nolock) shm On cli.CUSTNMBR = CAST(shm.AbreviaturaId COLLATE SQL_Latin1_General_CP1_CI_AS As varbinary(12)) "
                End Select

                SQL = SQL & "AND (cli.CUSTNMBR IN(SELECT IDGRUPO FROM EMP01001)) "

            Else

                SQL = SQL & "FROM RM00101 cli "
                SQL = SQL & "LEFT JOIN EMP00102 adi ON cli.CUSTNMBR = adi.CUSTNMBR "
                SQL = SQL & "LEFT JOIN EMP00103 cnv ON adi.CONVENIO = cnv.CUSTNMBR "
                SQL = SQL & "LEFT JOIN EMP00306 loc ON cli.CITY = loc.Codigo_Ciudad "
                SQL = SQL & "LEFT JOIN EMP00311 pla ON cnv.CodigoConvenio = pla.CodigoConvenio "
                SQL = SQL & "LEFT JOIN EMP00302 prv ON cli.STATE = prv.STATE "
                SQL = SQL & "LEFT JOIN RM00201 rub ON cli.CUSTCLAS = rub.CLASSID "

                Select Case pMod
                    Case modProcess.LastUpdates
                        SQL = SQL & "WHERE ((cli.MODIFDT >= '" & DateToSql(Now.Date.AddDays(-25), , True, True) & "') "
                        SQL = SQL & "OR (cli.CUSTNMBR IN(SELECT IDGRUPO FROM EMP00100 WHERE (MODIFDT >= '" & DateToSql(Now.Date,, True, True) & "') OR (Fecha_Alta >= '" & DateToSql(Now.Date,, True, True) & "')))) "
                End Select

            End If

            If pSelCli <> "" Then
                If Not SQL.Contains("WHERE") Then
                    SQL = SQL & "WHERE (cli.CUSTNMBR = '" & pSelCli & "') "
                Else
                    SQL = SQL & "AND (cli.CUSTNMBR = '" & pSelCli & "') "
                End If
            End If

            If Val(ConfigurationManager.AppSettings.Get("modSorted")) = 0 Then
                SQL = SQL & "ORDER BY cli.CUSTNMBR"
            Else
                SQL = SQL & "ORDER BY cli.CUSTNMBR DESC"
            End If

            Dim cmdDyn As New SqlCommand(SQL, cnnsNET(cnnDynamic), cnnsTransNET(cnnDynamic))
            cmdDyn.CommandTimeout = 600
            Dim dt As New DataTable
            dt.Load(cmdDyn.ExecuteReader)
            Dim vIdx As Integer

            AddLog("")
            AddLog(SQL)
            AddLog("")
            AddLog("Se procesarán " & dt.Rows.Count & " clientes")
            AddLog("")

            'Me.pgbProceso.Value = 0
            'Me.pgbProceso.Maximum = dt.Rows.Count

            For vIdx = 0 To dt.Rows.Count - 1

                Dim vAddNew As Boolean = False

                If Not objClientes.Abrir(objClientes.GetIDByAbreviaturaId(dt(vIdx)("CUSTNMBR").ToString.Trim())) Then

                    objClientes.CleanProperties(objClientes)
                    vAddNew = True

                End If

                Dim vEvl As Boolean = True

                If pMod = modProcess.NewInactives Then

                    If Math.Abs(DateDiff(DateInterval.Day, objClientes.regFechaHora, Now)) <= 0 Then

                        vEvl = False

                    End If

                End If

                If vEvl Then

                    vStep = "Cabecera"

                    objClientes.AbreviaturaId = dt(vIdx)("CUSTNMBR").ToString.Trim()

                    vStep = "Razón Social"

                    objClientes.RazonSocial = Microsoft.VisualBasic.Left(dt(vIdx)("STMTNAME").ToString.Trim(), 100)

                    vStep = "Rubro"

                    If Not IsDBNull(dt(vIdx)("RubroCodigo")) Then
                        objClientes.RubroClienteId.SetObjectId(Me.getRubroId(dt(vIdx)("RubroCodigo"), dt(vIdx)("RubroDescripcion")))
                    End If

                    vStep = "Domicilio"

                    If Not IsDBNull(dt(vIdx)("ADDRESS1")) Then objClientes.Domicilio.dmCalle = Microsoft.VisualBasic.Left(dt(vIdx)("ADDRESS1").ToString.Trim(), 200)
                    If Not IsDBNull(dt(vIdx)("ADDRESS2")) Then objClientes.Domicilio.dmReferencia = dt(vIdx)("ADDRESS2").ToString.Trim()
                    If Not IsDBNull(dt(vIdx)("ADDRESS3")) Then objClientes.Domicilio.dmEntreCalle1 = dt(vIdx)("ADDRESS3").ToString.Trim()

                    vStep = "Localidad"

                    If Not IsDBNull(dt(vIdx)("LocalidadCodigo")) AndAlso Not IsDBNull(dt(vIdx)("ProvinciaCodigo")) Then
                        objClientes.LocalidadId.SetObjectId(Me.getLocalidadId(dt(vIdx)("LocalidadCodigo"), dt(vIdx)("LocalidadDescripcion"), dt(vIdx)("ProvinciaCodigo"), dt(vIdx)("ProvinciaDescripcion")))
                    End If

                    vStep = "CPO"

                    If Not IsDBNull(dt(vIdx)("ZIP")) Then objClientes.CodigoPostal = dt(vIdx)("ZIP").ToString.Trim()

                    vStep = "Situacion IVA"

                    If Not IsDBNull(dt(vIdx)("TAXSCHID")) Then
                        If dt(vIdx)("TAXSCHID") = "EXCENTO VTAS" Then
                            objClientes.SituacionIvaId.SetObjectId(vIvaEX)
                        Else
                            objClientes.SituacionIvaId.SetObjectId(vIvaCF)
                        End If
                    End If

                    '-----> Fechas de Alta y Baja
                    vStep = "Fechas de Alta y Baja"

                    Try
                        objClientes.FecEgreso = dt(vIdx)("FECHSUSP")
                    Catch ex As Exception
                        objClientes.FecEgreso = Now.Date.AddDays(-1)
                    End Try
                    Try
                        objClientes.FecIngreso = dt(vIdx)("FECEMCNT")
                    Catch ex As Exception
                        objClientes.FecIngreso = Now.Date
                    End Try

                    '-----> Convenio
                    vStep = "Covenio"

                    If Not IsDBNull(dt(vIdx)("CONVENIO")) Then
                        objClientes.ConvenioId.SetObjectId(objClientes.GetIDByAbreviaturaId(dt(vIdx)("CONVENIO").ToString.Trim()))
                        If objClientes.ConvenioId.GetObjectId = 0 Then
                            If dt(vIdx)("CONVENIO") <> "" Then
                                AddLog("No existía el convenio " & dt(vIdx)("CONVENIO").ToString.Trim())
                            End If
                            objClientes.TipoConvenio = 0
                        Else
                            objClientes.TipoConvenio = 1
                        End If
                    Else
                        objClientes.TipoConvenio = 1
                    End If

                    '-----> Plan
                    If Not IsDBNull(dt(vIdx)("PlanCodigo")) AndAlso Not IsDBNull(dt(vIdx)("PlanDescripcion")) Then
                        vStep = "Plan"
                        objClientes.PlanId.SetObjectId(Me.getPlanId(dt(vIdx)("PlanCodigo"), dt(vIdx)("PlanDescripcion")))
                    End If

                    '-----> Otros
                    vStep = "Otros"

                    objClientes.EmpresaLegalId.SetObjectId(vEmpId)
                    objClientes.CobradorId.SetObjectId(vCobId)
                    objClientes.FormaPagoId.SetObjectId(vFpg)

                    vStep = "Salvando Cliente"

                    If objClientes.Salvar(objClientes, , False) Then

                        '------> Contactos
                        Dim vTel1 As String = ""
                        Dim vTel2 As String = ""

                        If Not IsDBNull(dt(vIdx)("PHONE1")) Then vTel1 = dt(vIdx)("PHONE1").ToString.Trim()
                        If Not IsDBNull(dt(vIdx)("PHONE2")) Then vTel2 = dt(vIdx)("PHONE2").ToString.Trim()

                        vStep = "Contacto 1"
                        SetContacto(objClientes.ID, vAddNew, vTel1, "PRINCIPAL", "")

                        vStep = "Contacto 2"
                        SetContacto(objClientes.ID, vAddNew, vTel2, "CELULAR", "")

                        vStep = "Plan Principal"
                        Dim objClientesPlanes As New conClientesPlanes
                        objClientesPlanes.SetPlanPrincipal(objClientes.ID, objClientes.PlanId.ID)
                        objClientesPlanes = Nothing

                        '------> Integrantes

                        Dim objIntegrante As New conClientesIntegrantes()

                        SQL = "Select ite.*, "
                        SQL = SQL & "prv.STATE As ProvinciaCodigo, prv.DSCRIPTN As ProvinciaDescripcion, "
                        SQL = SQL & "loc.Codigo_Ciudad As LocalidadCodigo, loc.Descripcion_Ciudad As LocalidadDescripcion "

                        If Not flgGuadalajara Then
                            SQL = SQL & "FROM RM00101 (nolock) cli "
                            SQL = SQL & "LEFT JOIN EMP01002 (nolock) adi On cli.CUSTNMBR = adi.CUSTNMBR "
                            SQL = SQL & "LEFT JOIN EMP01001 (nolock) ite On adi.CUSTNMBR = ite.IDGRUPO "
                            SQL = SQL & "LEFT JOIN EMP01007 (nolock) loc On ite.Codigo_Ciudad = loc.Codigo_Ciudad "
                            SQL = SQL & "LEFT JOIN EMP41004 (nolock) prv On ite.STATE = prv.STATE "
                            SQL = SQL & "WHERE (cli.CUSTNMBR = '" & objClientes.AbreviaturaId & "') "
                            SQL = SQL & "ORDER BY ite.IDSOCIO"
                        Else
                            SQL = SQL & "FROM RM00101 cli "
                            SQL = SQL & "LEFT JOIN EMP00102 adi ON cli.CUSTNMBR = adi.CUSTNMBR "
                            SQL = SQL & "LEFT JOIN EMP00100 ite ON adi.CUSTNMBR = ite.IDGRUPO "
                            SQL = SQL & "LEFT JOIN EMP00306 loc ON ite.Codigo_Ciudad = loc.Codigo_Ciudad "
                            SQL = SQL & "LEFT JOIN EMP00302 prv ON ite.STATE = prv.STATE "
                            SQL = SQL & "WHERE (cli.CUSTNMBR = '" & objClientes.AbreviaturaId & "') "
                            SQL = SQL & "ORDER BY ite.IDSOCIO"
                        End If

                        cmdDyn.CommandText = SQL
                        Dim dtIte As New DataTable
                        Dim vIdxIte As Integer
                        Dim vNoIte As Boolean = True

                        Dim vApe As String = "", vNom As String = ""

                        dtIte.Load(cmdDyn.ExecuteReader)

                        If dtIte.Rows.Count > 0 AndAlso Not IsDBNull(dtIte(0)("IDSOCIO")) Then vNoIte = False

                        If vNoIte Then

                            vStep = "Integrante Titular"

                            '----> Aplico sobre titular!
                            SQL = "SELECT TOP 1 ID FROM ClientesIntegrantes WHERE ClienteId = " & objClientes.ID

                            Dim cmdTit As New SqlCommand(SQL, cnnsNET(cnnDefault), cnnsTransNET(cnnDefault))
                            Dim vTitId As String = CType(cmdTit.ExecuteScalar, String)
                            If vTitId Is Nothing Then

                                '------> Creo el titular

                                objIntegrante.CleanProperties(objIntegrante)
                                objIntegrante.ClienteId.SetObjectId(objClientes.ID)
                                objIntegrante.NroAfiliado = "0000000"
                                objIntegrante.TipoIntegrante = "PER"

                                SetApellidoNombre(objClientes.RazonSocial, vApe, vNom)

                                objIntegrante.Apellido = vApe
                                objIntegrante.Nombre = vNom

                                objIntegrante.IntegranteClasificacionId.SetObjectId(vClfId)

                                objIntegrante.LocalidadId.SetObjectId(objClientes.LocalidadId.ID)
                                objIntegrante.Domicilio = objClientes.Domicilio
                                objIntegrante.CodigoPostal = objClientes.CodigoPostal

                                objIntegrante.Telefono01 = vTel1
                                objIntegrante.Telefono01Fix = GetNumericValue(vTel1)
                                objIntegrante.Telefono02 = vTel2
                                objIntegrante.Telefono02Fix = GetNumericValue(vTel2)

                                If Not objIntegrante.Salvar(objIntegrante) Then
                                    AddLog("Error al salvar el integrante titular de " & objClientes.AbreviaturaId)
                                End If

                            End If

                        Else

                            vStep = "Integrantes DB"

                            For vIdxIte = 0 To dtIte.Rows.Count - 1

                                If Not objIntegrante.Abrir(objIntegrante.GetIDByNroAfiliado(objClientes.ID, dtIte(vIdxIte)("IDSOCIO").ToString().Trim())) Then

                                    objIntegrante.ClienteId.SetObjectId(objClientes.ID)
                                    objIntegrante.NroAfiliado = dtIte(vIdxIte)("IDSOCIO").ToString().Trim()
                                    objIntegrante.TipoIntegrante = "PER"

                                End If

                                SetApellidoNombre(dtIte(vIdxIte)("NOMBRE").ToString().Trim(), vApe, vNom)

                                objIntegrante.Apellido = vApe
                                objIntegrante.Nombre = vNom

                                If vIdxIte = 0 Then
                                    objIntegrante.IntegranteClasificacionId.SetObjectId(vClfId)
                                Else
                                    objIntegrante.IntegranteClasificacionId.SetObjectId(vClfId)
                                End If

                                If Not IsDBNull(dtIte(vIdxIte)("LocalidadCodigo")) AndAlso Not IsDBNull(dtIte(vIdxIte)("ProvinciaCodigo")) Then
                                    objIntegrante.LocalidadId.SetObjectId(Me.getLocalidadId(dtIte(vIdxIte)("LocalidadCodigo"), dtIte(vIdxIte)("LocalidadDescripcion"), dtIte(vIdxIte)("ProvinciaCodigo"), dtIte(vIdxIte)("ProvinciaDescripcion")))
                                End If

                                If Not IsDBNull(dtIte(vIdxIte)("ADDRESS1")) Then objIntegrante.Domicilio.dmCalle = dtIte(vIdxIte)("ADDRESS1").ToString().Trim()
                                If Not IsDBNull(dtIte(vIdxIte)("ADDRESS2")) Then objIntegrante.Domicilio.dmReferencia = dtIte(vIdxIte)("ADDRESS2").ToString().Trim()
                                If Not IsDBNull(dtIte(vIdxIte)("ADDRESS3")) Then objIntegrante.Domicilio.dmEntreCalle1 = dtIte(vIdxIte)("ADDRESS3").ToString().Trim()

                                Try
                                    objIntegrante.FecNacimiento = dtIte(vIdxIte)("FECNAC")
                                Catch ex As Exception

                                End Try

                                objIntegrante.Sexo = "F"
                                If Not IsDBNull(dtIte(vIdxIte)("SEXO")) AndAlso dtIte(vIdxIte)("SEXO") = 2 Then
                                    objIntegrante.Sexo = "M"
                                End If

                                If Not IsDBNull(dtIte(vIdxIte)("ZIP")) Then
                                    objIntegrante.CodigoPostal = dtIte(vIdxIte)("ZIP")
                                End If

                                Try
                                    objIntegrante.FecEgreso = dtIte(vIdxIte)("FECHSUSP")
                                Catch ex As Exception
                                    objIntegrante.FecEgreso = objIntegrante.FecEgreso.AddDays(-1)
                                End Try

                                Try
                                    objIntegrante.FecIngreso = dtIte(vIdxIte)("Fecha_Alta")
                                Catch ex As Exception
                                    objIntegrante.FecIngreso = Now.Date
                                End Try

                                If dtIte(vIdxIte)("IDEFUSR1") = 1 Then
                                    objIntegrante.Observaciones = "ALTO USUARIO"
                                Else
                                    objIntegrante.Observaciones = ""
                                End If

                                If dtIte(vIdxIte)("HOLD") = 1 AndAlso objIntegrante.FecEgreso < objIntegrante.FecIngreso Then
                                    objIntegrante.FecEgreso = Now.Date.AddDays(-1)
                                End If

                                If Not IsDBNull(dtIte(vIdxIte)("PHONE1")) AndAlso dtIte(vIdxIte)("PHONE1").ToString().Trim() <> "" Then
                                    objIntegrante.Telefono01 = dtIte(vIdxIte)("PHONE1").ToString().Trim()
                                    objIntegrante.Telefono01Fix = GetNumericValue(dtIte(vIdxIte)("PHONE1").ToString().Trim())
                                ElseIf Not IsDBNull(dtIte(vIdxIte)("PHONE2")) AndAlso dtIte(vIdxIte)("PHONE2").ToString().Trim() <> "" Then
                                    objIntegrante.Telefono01 = dtIte(vIdxIte)("PHONE2").ToString().Trim()
                                    objIntegrante.Telefono01Fix = GetNumericValue(dtIte(vIdxIte)("PHONE2").ToString().Trim())
                                ElseIf Not IsDBNull(dtIte(vIdxIte)("PHONE3")) AndAlso dtIte(vIdxIte)("PHONE3").ToString().Trim() <> "" Then
                                    objIntegrante.Telefono01 = dtIte(vIdxIte)("PHONE3").ToString().Trim()
                                    objIntegrante.Telefono01Fix = GetNumericValue(dtIte(vIdxIte)("PHONE3").ToString().Trim())
                                End If

                                If Not IsDBNull(dtIte(vIdxIte)("PHONE2")) AndAlso dtIte(vIdxIte)("PHONE2").ToString().Trim() <> "" Then
                                    objIntegrante.Telefono02 = dtIte(vIdxIte)("PHONE2").ToString().Trim()
                                    objIntegrante.Telefono02Fix = GetNumericValue(dtIte(vIdxIte)("PHONE2").ToString().Trim())
                                ElseIf Not IsDBNull(dtIte(vIdxIte)("PHONE3")) AndAlso dtIte(vIdxIte)("PHONE3").ToString().Trim() <> "" Then
                                    objIntegrante.Telefono02 = dtIte(vIdxIte)("PHONE3").ToString().Trim()
                                    objIntegrante.Telefono02Fix = GetNumericValue(dtIte(vIdxIte)("PHONE3").ToString().Trim())
                                End If

                                If Not objIntegrante.Salvar(objIntegrante) Then
                                    AddLog("Error al salvar el integrante de " & objClientes.AbreviaturaId & " - " & dtIte(vIdxIte)("IDSOCIO"))
                                End If

                            Next vIdxIte

                        End If

                        If vAddNew Then
                            AddLog(dt(vIdx)("CUSTNMBR").ToString.Trim & " - Creado")
                        Else
                            AddLog(dt(vIdx)("CUSTNMBR").ToString.Trim & " - Actualizado")
                        End If

                    Else

                        AddLog(dt(vIdx)("CUSTNMBR").ToString.Trim & " - Error: " & objClientes.MyLastExec.ErrorDescription)

                    End If

                Else

                    AddLog(dt(vIdx)("CUSTNMBR").ToString.Trim & " - Omitido")

                End If

                'Me.pgbProceso.Value = Me.pgbProceso.Value + 1
                'Me.pgbProceso.PerformStep()
                'Me.pgbProceso.Refresh()

            Next vIdx

            objClientes = Nothing
            objSituacionesIva = Nothing
            objRubros = Nothing
            objEmpresas = Nothing

        Catch ex As Exception
            AddLog(vStep & ": " & ex.Message)
        End Try
    End Sub
    Private Function getRubroId(ByVal pCod As String, ByVal pDes As String) As Int64

        getRubroId = 0

        Try

            pCod = pCod.Trim()
            pDes = pDes.Trim()

            Dim objRubros As New conRubrosClientes

            Dim vRubId As Int64 = objRubros.GetIDByAbreviaturaId(pCod)

            If vRubId = 0 Then

                objRubros.CleanProperties(objRubros)
                objRubros.AbreviaturaId = pCod
                objRubros.Descripcion = pDes

                If objRubros.Salvar(objRubros) Then
                    vRubId = objRubros.ID
                End If

            End If

            getRubroId = vRubId

        Catch ex As Exception
            AddLog("getRubroId: " & ex.Message)
        End Try
    End Function

    Private Function getLocalidadId(ByVal pCod As String, ByVal pDes As String, ByVal pCodPrv As String, ByVal pDesPrv As String) As Int64

        getLocalidadId = 0
        Dim vStep As String = ""

        Try

            pCod = pCod.Trim()
            pDes = pDes.Trim()
            pCodPrv = pCodPrv.Trim()
            pDesPrv = pDesPrv.Trim()

            Dim objLocalidades As New conLocalidades
            Dim objProvincias As New conProvincias
            Dim objPaises As New conPaises
            Dim objZonas As New conZonasGeograficas

            Dim vLocId As Int64 = objLocalidades.GetIDByAbreviaturaId(pCod)

            If vLocId = 0 Then

                vStep = "LocalidadId = 0"

                Dim vPaiId As Int64 = objPaises.GetIDByDescripcion("MEXICO")
                Dim vPrvId As Int64 = 0

                If pDesPrv <> "" Then

                    vPrvId = objProvincias.GetIDByDescripcion(vPaiId, pDesPrv)

                    If vPrvId = 0 Then

                        objProvincias.CleanProperties(objProvincias)
                        objProvincias.AbreviaturaId = pCodPrv
                        objProvincias.Descripcion = pDesPrv
                        objProvincias.PaisId.SetObjectId(vPaiId)

                        If objProvincias.Salvar(objProvincias) Then
                            vPrvId = objProvincias.ID
                        End If

                    End If

                    vStep = "Provincia OK"

                End If

                Dim dtZon As DataTable = objZonas.GetAll()

                objLocalidades.CleanProperties(objLocalidades)
                objLocalidades.AbreviaturaId = pCod
                objLocalidades.Descripcion = pDes
                objLocalidades.ProvinciaId.SetObjectId(vPrvId)

                If dtZon.Rows.Count > 0 Then
                    objLocalidades.ZonaGeograficaId.SetObjectId(dtZon(0)("ID"))
                End If

                vStep = "Saving Localidad"

                If objLocalidades.Salvar(objLocalidades) Then

                    vLocId = objLocalidades.ID

                End If

            End If

            getLocalidadId = vLocId

            objLocalidades = Nothing
            objProvincias = Nothing
            objPaises = Nothing
            objZonas = Nothing

        Catch ex As Exception
            AddLog("getLocalidadId: " & ex.Message & " Step:" & vStep)
        End Try
    End Function

    Private Function getPlanId(ByVal pPla As String, ByVal pNom As String) As Int64

        getPlanId = 0

        Try

            pPla = pPla.Trim()
            pNom = pNom.Trim()

            Dim objPlanes As New conPlanes

            Dim vPlaId As Int64 = objPlanes.GetIDByAbreviaturaId(pPla)

            If vPlaId = 0 Then

                objPlanes.CleanProperties(objPlanes)
                objPlanes.AbreviaturaId = pPla
                objPlanes.Descripcion = pNom

                If objPlanes.Salvar(objPlanes) Then

                    vPlaId = objPlanes.ID

                    Dim objPlanesGrados As New conPlanesGradosOperativos
                    Dim objGrados As New conGradosOperativos

                    Dim dtGdo As DataTable = objGrados.GetAll(gdoClasificacion.gdoIncidente)
                    Dim vIdx As Integer

                    For vIdx = 0 To dtGdo.Rows.Count - 1

                        If Not objPlanesGrados.Abrir(objPlanesGrados.GetIDByIndex(vPlaId, dtGdo(vIdx)("ID"))) Then

                            objPlanesGrados.PlanId.SetObjectId(vPlaId)
                            objPlanesGrados.GradoOperativoId.SetObjectId(dtGdo(vIdx)("ID"))
                            objPlanesGrados.flgCubierto = 1
                            objPlanesGrados.Salvar(objPlanesGrados)

                        End If

                    Next vIdx

                    objGrados = Nothing
                    objPlanesGrados = Nothing

                End If

            End If

            getPlanId = vPlaId

        Catch ex As Exception
            AddLog("getPlanId: " & ex.Message)
        End Try
    End Function

    Private Sub SetApellidoNombre(ByVal pRaz As String, ByRef pApe As String, ByRef pNom As String)
        Try

            'Dim vNom() As String = pRaz.Split(" ")
            'Dim vNomIdx As Integer

            'pNom = vNom(0)

            'If vNom.Length - 2 > 0 Then
            '    For vNomIdx = 1 To vNom.Length - 2
            '        If vNomIdx = 1 Then
            '            pNom = vNom(vNomIdx)
            '        Else
            '            pNom = pNom & " " & vNom(vNomIdx)
            '        End If
            '    Next
            'End If

            'pNom = Microsoft.VisualBasic.Left(pNom, 70)
            'pApe = Microsoft.VisualBasic.Left(vNom(vNom.Length - 1), 50)

            pApe = "/"
            pNom = Microsoft.VisualBasic.Left(pRaz, 100)

        Catch ex As Exception
            AddLog("SetApellidoNombre: " & ex.Message)
        End Try
    End Sub

    Private Sub SetContacto(pCli As Int64, pAddNew As Boolean, pTel As String, pNom As String, pMai As String)

        Try

            Dim objContacto As New conClientesContactos

            If pAddNew Then

                If pTel <> "" Then

                    objContacto.CleanProperties(objContacto)
                    objContacto.ClienteId.SetObjectId(pCli)
                    objContacto.Contacto.Nombre = "PRINCIPAL"
                    If Not IsDBNull(pMai) Then objContacto.Contacto.Email = pMai
                    objContacto.Contacto.Telefono = pTel

                    If objContacto.Salvar(objContacto) Then
                    End If

                End If

            Else

                Dim vConId As Int64 = objContacto.GetIDByNombre(pCli, pNom)

                If pTel <> "" Then

                    If Not objContacto.Abrir(vConId) Then
                        objContacto.ClienteId.SetObjectId(pCli)
                        objContacto.Contacto.Nombre = "PRINCIPAL"
                    End If
                    If Not IsDBNull(pMai) Then objContacto.Contacto.Email = pMai
                    objContacto.Contacto.Telefono = Microsoft.VisualBasic.Left(pTel, 20)

                    If objContacto.Salvar(objContacto) Then
                    End If

                Else

                    objContacto.Eliminar(vConId)

                End If

            End If


        Catch ex As Exception

            AddLog("Salvando contacto: " & ex.Message)

        End Try

    End Sub

End Class
