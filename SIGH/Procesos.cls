VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para varios procesos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ProgressRpt1 As XP_ProgressBar
Dim mo_ProgressRpt2 As XP_ProgressBar
Dim mo_ProgressRpt3 As XP_ProgressBar
Dim mo_ProgressRpt4 As XP_ProgressBar
Dim mo_ProgressRpt5 As XP_ProgressBar
Dim mo_ProgressRpt6 As XP_ProgressBar
Dim mo_ProgressRpt7 As XP_ProgressBar
Dim mo_ProgressRpt8 As XP_ProgressBar
Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ms_MensajeError As String
Dim lcSql As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim lcParaMsgbox As String
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim oRsDxCpt As New Recordset
Dim ml_ValorActual As String

Property Get ValorActual() As String
   ValorActual = ml_ValorActual
End Property

Property Let ValorActual(lValue As String)
   ml_ValorActual = lValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let MensajeError(lValue As String)
    ms_MensajeError = lValue
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Set progressRpt1(oValue As XP_ProgressBar)
    Set mo_ProgressRpt1 = oValue
End Property

Property Set progressRpt2(oValue As XP_ProgressBar)
    Set mo_ProgressRpt2 = oValue
End Property
Property Set progressRpt3(oValue As XP_ProgressBar)
    Set mo_ProgressRpt3 = oValue
End Property
Property Set progressRpt4(oValue As XP_ProgressBar)
    Set mo_ProgressRpt4 = oValue
End Property
Property Set progressRpt5(oValue As XP_ProgressBar)
    Set mo_ProgressRpt5 = oValue
End Property
Property Set progressRpt6(oValue As XP_ProgressBar)
    Set mo_ProgressRpt6 = oValue
End Property
Property Set progressRpt7(oValue As XP_ProgressBar)
    Set mo_ProgressRpt7 = oValue
End Property

Property Set progressRpt8(oValue As XP_ProgressBar) 'Agregado Yamill
    Set mo_ProgressRpt8 = oValue
End Property




'*****************************HIS***************************
Sub ExportaDatosAlHis()
'    If cmbAnio.Text = "" Then
'       MsgBox "Por favor elija el AÑO", vbInformation, "Mensaje"
'       Exit Sub
'    End If
'    If cmbMes.Text = "" Then
'       MsgBox "Por favor elija el MES", vbInformation, "Mensaje"
'       Exit Sub
'    End If
'    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
'        Me.MousePointer = 11
'        On Error GoTo ErrorProceso
'        Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'        Dim mo_AdminAdmision As New ReglasAdmision
'        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
'        Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
'        Dim lcBuscaParametro As New SIGHDatos.Parametros
'        Dim rsReporte As New Recordset
'        Dim oRsTmp As New Recordset
'        Dim oRsFox1 As New Recordset
'        Dim oRsFox2 As New Recordset
'        Dim oRsFox3 As New Recordset
'        Dim oRsEdadHis As New Recordset
'        Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
'        Dim lcTDx1 As String, lcTDx2 As String, lcTDx3 As String, lcTDx4 As String, lcTDx5 As String, lcTDx6 As String
'        Dim lcCDx1 As String, lcCDx2 As String, lcCDx3 As String, lcCDx4 As String, lcCDx5 As String, lcCDx6 As String
'        Dim lcLDx1 As String, lcLDx2 As String, lcLDx3 As String, lcLDx4 As String, lcLDx5 As String, lcLDx6 As String
'        Dim lnTot_reg As Long, lnEstancia As Integer
'        Dim lcFec1 As String, lcFec2 As String
'        Dim lcMes As String, lcAnio As String
'        Dim lnUltimoDiaMes As Integer
'        Dim lnIdEdad As Long, lnDesde As Integer, lnHasta As Integer
'        Dim lcDistritoDomicilio As String, lcCodigoServicioHIS As String
'        Dim lnTotal As Long, lnIdEmpleado As Long, lnNumPag As Long, lnReg As Long
'        Dim lnTTot_pag As Long, lnTTot_reg As Long, lnMaxPacientesXhoja As Long, lnMaxHISxLOTE As Long
'        Dim lcCod_dig As String, lcCod_2000 As String, lcMt As String, lcNomLote As String
'        Dim ldFecha As Date, lbAumentoPagina As Boolean
'        Dim lcCodif As String, lcPlaza As String, lcCod_servsa As String, lcMensajes As String
'        Dim lcFichaOhistoria As String
'        Const lcCharInicio As String = "A"
'        '
'        lcMes = Right("0" & Trim(Str(cmbMes.ListIndex + 1)), 2)
'        lcAnio = cmbAnio.Text
'        lcFec1 = ("01/" & lcMes & "/" & lcAnio)
'        lnUltimoDiaMes = DevuelveUltimoDiaDelMes(Val(lcMes), Val(cmbAnio.Text))
'        lcFec2 = Right("0" & Trim(Str(lnUltimoDiaMes)), 2) & "/" & lcMes & "/" & lcAnio
'        lnMaxPacientesXhoja = Val(lcBuscaParametro.SeleccionaFilaParametro(272))
'        lnMaxHISxLOTE = Val(lcBuscaParametro.SeleccionaFilaParametro(271))
'        lcCod_dig = Val(lcBuscaParametro.SeleccionaFilaParametro(270))
'        '
'        oConexion.CommandTimeout = 300
'        oConexion.Open SIGHEntidades.CadenaConexion
'        '
'        oConexionFox.CommandTimeout = 300
'        oConexionFox.Open "DSN=his"
'        '
'        lcSql = "delete from his_lote"
'        oRsFox3.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'        oRsFox3.Open "select * from his_lote", oConexionFox, adOpenKeyset, adLockOptimistic
'        lcSql = "delete from his_his1"
'        oRsFox2.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'        oRsFox2.Open "select * from his_his1", oConexionFox, adOpenKeyset, adLockOptimistic
'        lcSql = "delete from his_hisA"
'        oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'        oRsFox1.Open "select * from his_hisA", oConexionFox, adOpenKeyset, adLockOptimistic
'        '
'        Set rsReporte = AtencionesSeleccionarPorFechasDeIngresoYtipoServicio(lcMes, lcAnio)
'        lnTotal = rsReporte.RecordCount
'        If lnTotal = 0 Then
'           MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
'        Else
'           ProgressBar1.Min = 0
'           ProgressBar1.Max = lnTotal
'           lnTotal = 0
'           lcCod_2000 = lcBuscaParametro.SeleccionaFilaParametro(208)
'           lcMensajes = ""
'           rsReporte.MoveFirst
'           Do While Not rsReporte.EOF
'                If IsNull(rsReporte.Fields!LoteHis) Then
'                   lcMensajes = lcMensajes & "Tiene que ingresar el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
'                End If
'                lcMt = IIf(rsReporte.Fields!HoraIngreso >= Me.txtTarde.Text, "T", "M")
'                lcNomLote = rsReporte.Fields!LoteHis & lcMt
'                lnIdEmpleado = rsReporte.Fields!IdEmpleado
'                lnTTot_pag = 0: lnTTot_reg = 0
'                If IsNull(rsReporte.Fields!codigoServicioHIS) Then
'                   MsgBox "No se pudo procesar" & Chr(13) & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio
'                   Exit Do
'                End If
'                lbAumentoPagina = False
'                Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado
'                    If IsNull(rsReporte.Fields!codigoServicioHIS) Then
'                        lcMensajes = lcMensajes & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio & Chr(13)
'                        MsgBox lcMensajes, vbInformation, Me.Caption
'                        Exit Sub
'                    End If
'                    lcCodigoServicioHIS = IIf(IsNull(rsReporte.Fields!codigoServicioHIS), "", rsReporte.Fields!codigoServicioHIS)
'                    ldFecha = rsReporte.Fields!FechaIngreso
'                    lcCodif = IIf(IsNull(rsReporte.Fields!CodigoPlanilla), "", rsReporte.Fields!CodigoPlanilla)
'                    lcPlaza = lcCodif
'                    lcCod_servsa = lcCodigoServicioHIS
'                    If lbAumentoPagina = True Then
'                       lnNumPag = lnNumPag + 1
'                    Else
'                       lnNumPag = 1
'                    End If
'                    lnTot_reg = 0: lnReg = 0
'                    Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso
'                        If IsNull(rsReporte.Fields!NroHistoriaClinica) And IsNull(rsReporte.Fields!FichaFamiliar) Then
'                           MsgBox "El Paciente No existe para la CUENTA N° " & rsReporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, Me.Caption
'                           Me.MousePointer = 1
'                           Exit Sub
'                        End If
'                        lcFichaOhistoria = Trim(Str(rsReporte.Fields!NroHistoriaClinica))
'                        If Not IsNull(rsReporte.Fields!FichaFamiliar) Then
'                           If Len(Trim(rsReporte.Fields!FichaFamiliar)) > 2 Then
'                              lcFichaOhistoria = Trim(rsReporte.Fields!FichaFamiliar)
'                           End If
'                        End If
'                        '
''If Day(rsReporte.Fields!FechaIngreso) = 31 And rsReporte.Fields!IdEmpleado = 1202 And rsReporte.Fields!idServicioIngreso = 75 And Trim(Str(rsReporte.Fields!NroHistoriaClinica)) = "372704" Then
''  lcDx1 = ""
''End If
'                        lcDx1 = "": lcDx2 = "": lcDx3 = "": lcDx4 = "": lcDx5 = "": lcDx6 = ""
'                        lcTDx1 = "": lcTDx2 = "": lcTDx3 = "": lcTDx4 = "": lcTDx5 = "": lcTDx6 = ""
'                        lcCDx1 = "": lcCDx2 = "": lcCDx3 = "": lcCDx4 = "": lcCDx5 = "": lcCDx6 = ""
'                        lcLDx1 = "": lcLDx2 = "": lcLDx3 = "": lcLDx4 = "": lcLDx5 = "": lcLDx6 = ""
'                        lcSql = "select " & _
'                                "   AtencionesDiagnosticos.IdDiagnostico," & _
'                                "   Diagnosticos.CodigoCIE2004," & _
'                                "   Diagnosticos.Descripcion," & _
'                                "   Diagnosticos.ClaseDxHIS," & _
'                                "   AtencionesDiagnosticos.labConfHIS," & _
'                                "   AtencionesDiagnosticos.IdSubClasificacionDx as IdTipoDiagnostico," & _
'                                "   convert(varchar(5), SubClasificacionDiagnosticos.Codigo) + '  =  ' + SubClasificacionDiagnosticos.Descripcion as DescripcionTipoDx" & _
'                                " from (AtencionesDiagnosticos left join SubClasificacionDiagnosticos" & _
'                                "   on AtencionesDiagnosticos.IdSubClasificacionDx = SubClasificacionDiagnosticos.IdSubClasificacionDx)" & _
'                                "   left join Diagnosticos on AtencionesDiagnosticos.IdDiagnostico = Diagnosticos.IdDiagnostico" & _
'                                " where AtencionesDiagnosticos.IdAtencion = " & rsReporte.Fields!idAtencion & _
'                                "   and AtencionesDiagnosticos.IdClasificacionDx = 1"
'                        oRsTmp.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
'                        If oRsTmp.RecordCount > 0 Then
'                           lnUltimoDiaMes = 1
'                           oRsTmp.MoveFirst
'                           Do While Not oRsTmp.EOF
'                              Select Case lnUltimoDiaMes
'                              Case 1
'                                  lcDx1 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx1 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx1 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx1 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              Case 2
'                                  lcDx2 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx2 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx2 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx2 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              Case 3
'                                  lcDx3 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx3 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx3 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx3 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              Case 4
'                                  lcDx4 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx4 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx4 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx4 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              Case 5
'                                  lcDx5 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx5 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx5 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx5 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              Case 6
'                                  lcDx6 = DevuelveCodigoEnTablaHIS_TabDiag(SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6)))
'                                  lcTDx6 = Left(oRsTmp.Fields!DescripcionTipoDx, 1)
'                                  lcCDx6 = Left(oRsTmp.Fields!ClaseDxHIS, 1)
'                                  lcLDx6 = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'                              End Select
'                              lnUltimoDiaMes = lnUltimoDiaMes + 1
'                              oRsTmp.MoveNext
'                           Loop
'                        End If
'                        oRsTmp.Close
'                        '
'                        lcDistritoDomicilio = Space(6)
'                        If Not IsNull(rsReporte.Fields!idDistrito) Then
'                           lcDistritoDomicilio = Right("0" & Trim(Str(rsReporte.Fields!idDistrito)), 6)
'                        End If
'                        '
'                        lnTot_reg = lnTot_reg + 1
'                        lnReg = lnReg + 1
'                        oRsFox1.AddNew
'                        oRsFox1.Fields!Cod_2000 = Val(lcCod_2000)
'                        oRsFox1.Fields!Ano = Val(lcAnio)
'                        oRsFox1.Fields!Mes = Val(lcMes)
'                        oRsFox1.Fields!Nom_lote = lcNomLote
'                        oRsFox1.Fields!Num_pag = lnNumPag
'                        oRsFox1.Fields!num_reg = lnReg
'                        oRsFox1.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
'                        oRsFox1.Fields!fichafam = lcFichaOhistoria
'                        oRsFox1.Fields!Cod_Dpto = Left(lcDistritoDomicilio, 2)
'                        oRsFox1.Fields!Cod_Prov = Mid(lcDistritoDomicilio, 3, 2)
'                        oRsFox1.Fields!Cod_Dist = Right(lcDistritoDomicilio, 2)
'                        oRsFox1.Fields!Edad = rsReporte.Fields!Edad
'                        oRsFox1.Fields!tip_edad = rsReporte.Fields!EdadCodigo
'                        oRsFox1.Fields!Sexo = IIf(rsReporte.Fields!idTipoSexo = 1, "M", "F")
'                        oRsFox1.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
'                        oRsFox1.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
'                        oRsFox1.Fields!diagnost1 = lcTDx1
'                        oRsFox1.Fields!labconf1 = lcLDx1
'                        oRsFox1.Fields!codigo1 = Val(lcDx1)
'                        oRsFox1.Fields!diagnost2 = lcTDx2
'                        oRsFox1.Fields!labconf2 = lcLDx2
'                        oRsFox1.Fields!codigo2 = Val(lcDx2)
'                        oRsFox1.Fields!diagnost3 = lcTDx3
'                        oRsFox1.Fields!labconf3 = lcLDx3
'                        oRsFox1.Fields!codigo3 = Val(lcDx3)
'                        oRsFox1.Fields!diagnost4 = lcTDx4
'                        oRsFox1.Fields!labconf4 = lcLDx4
'                        oRsFox1.Fields!codigo4 = Val(lcDx4)
'                        oRsFox1.Fields!diagnost5 = lcTDx5
'                        oRsFox1.Fields!labconf5 = lcLDx5
'                        oRsFox1.Fields!codigo5 = Val(lcDx5)
'                        oRsFox1.Fields!diagnost6 = lcTDx6
'                        oRsFox1.Fields!labconf6 = lcLDx6
'                        oRsFox1.Fields!codigo6 = Val(lcDx6)
'                        oRsFox1.Fields!esta_reg = "2"
'                        oRsFox1.Fields!mt = lcMt
'                        oRsFox1.Update
'                        lnTotal = lnTotal + 1
'                        DoEvents
'                        ProgressBar1.Value = lnTotal
'                        Me.Refresh
'                        rsReporte.MoveNext
'                        If rsReporte.EOF Then
'                           Exit Do
'                        End If
'                        If lnReg >= lnMaxPacientesXhoja And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso Then
'                           lbAumentoPagina = True
'                           lnReg = 0
'                           Exit Do
'                        Else
'                           lbAumentoPagina = False
'                        End If
'                    Loop
'                    oRsFox2.AddNew
'                    oRsFox2.Fields!Cod_2000 = Val(lcCod_2000)
'                    oRsFox2.Fields!Ano = Val(lcAnio)
'                    oRsFox2.Fields!Mes = Val(lcMes)
'                    oRsFox2.Fields!Nom_lote = lcNomLote
'                    oRsFox2.Fields!Num_pag = lnNumPag
'                    oRsFox2.Fields!Codif = Val(lcCodif)
'                    oRsFox2.Fields!cod_ServSa = Val(Left(lcCod_servsa, 3))
'                    oRsFox2.Fields!plaza = Val(lcCodif)
'                    oRsFox2.Fields!Esta_pag = "2"
'                    oRsFox2.Fields!Tot_reg = lnTot_reg
'                    oRsFox2.Fields!FlagEnvio = Space(1)
'                    oRsFox2.Fields!mt = lcMt
'                    oRsFox2.Update
'                    lnTTot_pag = lnTTot_pag + lnNumPag
'                    lnTTot_reg = lnTTot_reg + lnTot_reg
'                    If rsReporte.EOF Then
'                       Exit Do
'                    End If
'                Loop
'                oRsFox3.AddNew
'                oRsFox3.Fields!Cod_2000 = Val(lcCod_2000)
'                oRsFox3.Fields!Ano = Val(lcAnio)
'                oRsFox3.Fields!Mes = Val(lcMes)
'                oRsFox3.Fields!Nom_lote = lcNomLote
'                oRsFox3.Fields!cod_dig = lcCod_dig
'                oRsFox3.Fields!esta_lote = "LC"
'                oRsFox3.Fields!tot_pag = lnTTot_pag
'                oRsFox3.Fields!Tot_reg = lnTTot_reg
'                oRsFox3.Fields!Tot_rmalos = 0
'                oRsFox3.Fields!tot_pgcarg = 0
'                oRsFox3.Fields!esta_local = Space(1)
'                oRsFox3.Fields!nomfile = Space(14)
'                oRsFox3.Fields!flagReport = "3"
'                oRsFox3.Fields!mt = lcMt
'                oRsFox3.Update
'           Loop
'           Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al HIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
'        End If
'        oRsFox1.Close
'        oRsFox2.Close
'        oRsFox3.Close
'        oConexionFox.Close
'        oConexion.Close
'        If lcMensajes <> "" Then
'           MsgBox lcMensajes, vbInformation, Me.Caption
'        End If
'        Me.MousePointer = 1
'        Me.Visible = False
'        Exit Sub
'    End If
'ErrorProceso:
'    MsgBox Err.Description
''    Resume

End Sub

Public Sub CrearTablaTemporalDxCPT() 'Frank 12092014
    If oRsDxCpt.State = 1 Then
       Set oRsDxCpt = Nothing
    End If
    With oRsDxCpt
        .Fields.Append "IdDxCpt", adInteger
        .Fields.Append "Dx", adVarChar, 255, adFldIsNullable
        .Fields.Append "TDx", adVarChar, 255, adFldIsNullable
        .Fields.Append "CDx", adVarChar, 255, adFldIsNullable
        .Fields.Append "LDx", adVarChar, 255, adFldIsNullable
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub LimpiarTablaDxCpt() 'Frank 12092014
    With oRsDxCpt
        If .RecordCount > 0 Then
           .MoveFirst
           Do While Not .EOF
              .Delete
              .Update
              .MoveNext
           Loop
        End If
    End With
End Sub



'debb-16/05/2016
'Sub ExportaDAtosAlHISv4(lcTarde As String, lnChkExportaCPT As Integer, lnIdMesElegido As Integer, lcAnio As String, _
'                        lbExportaAnovafis As Boolean, lbAgregaOrdenesMedicas As Boolean)
'
'10            On Error GoTo ErrorProceso
'              Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'              Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
'              Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
'              Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
'              Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
'
'              Dim oConexion As New Connection
'              Dim oConexionFox As New Connection
'
'              Dim rsReporte As New Recordset
'              Dim oRsTmp As New Recordset
'              Dim oRsFox1 As New Recordset
'              Dim oRsFox2 As New Recordset
'              Dim oRsFox3 As New Recordset
'              Dim oRsFox4 As New Recordset
'              Dim oRsFox5 As New Recordset
'              Dim oRsEdadHis As New Recordset
'              Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
'              Dim lcTDx1 As String, lcTDx2 As String, lcTDx3 As String, lcTDx4 As String, lcTDx5 As String, lcTDx6 As String
'              Dim lcCDx1 As String, lcCDx2 As String, lcCDx3 As String, lcCDx4 As String, lcCDx5 As String, lcCDx6 As String
'              Dim lcLDx1 As String, lcLDx2 As String, lcLDx3 As String, lcLDx4 As String, lcLDx5 As String, lcLDx6 As String
'              Dim lnTot_reg As Long, lnEstancia As Integer
'              Dim lcFec1 As String, lcFec2 As String
'              Dim lcMes As String, lnEdadEnAtencion As Integer, lnAsc As Integer
'              Dim lnUltimoDiaMes As Integer, lbContinuar As Boolean
'              Dim lnIdEdad As Long, lnDesde As Integer, lnHasta As Integer
'              Dim lcDistritoDomicilio As String, lcCodigoServicioHIS As String
'              Dim lnTotal As Long, lnIdEmpleado As Long, lnNumPag As Long, lnReg As Long
'              Dim lnTTot_pag As Long, lnTTot_reg As Long, lnMaxPacientesXhoja As Long, lnMaxHISxLOTE As Long
'              Dim lcCod_dig As String, lcCod_2000 As String, lcMt As String, lcNomLote As String
'              Dim ldFecha As Date, lbAumentoPagina As Boolean, lbEsCentroSalud As Boolean
'              Dim lcCodif As String, lcPlaza As String, lcCod_servsa As String, lcTipoDocumentoHisDoctor As String
'              Dim lcFichaOhistoria As String, lcDNIpaciente As String, lcFuenteFinanciamientoPaciente As String, lcEtniaPaciente As String
'
'              Dim lnGrupo6DxCPT As Integer
'              Dim lnRestoDxCPT As Integer
'              Dim lnIndiceDxCPT As Integer
'              Dim lnContador As Integer
'              Dim lnContDxCpt As Integer
'              Const lcCharInicio As String = "0"
'              Dim oRsImpHIStodos As New Recordset, lnoRsImpHIStodos As Long
'              Dim rsDx1 As New Recordset
'              Dim oConexionExterna As New Connection
'              Dim oRsNovafis_edades As New Recordset
'              Dim lcAD040 As String
'              Dim lcRutaExportar As String, lcPaisPaciente As String, lcTipoDocumentoPaciente As String
'              Dim lnEdadEnTablaEdades As Long, lcGestante As String, lbNuevoRegistro As Boolean, lcDNIpaciente1 As String
'              Dim lnDNImedico1 As String
'              Dim lnAscFinal1 As Integer, lnAscFinal2 As Integer
'              '
'20            lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'              '
'30            oConexionExterna.CursorLocation = adUseClient
'40            oConexionExterna.CommandTimeout = 150
'50            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'              '
'              If lbExportaAnovafis = True Then
'60               Set oRsNovafis_edades = mo_reglasComunes.novafis_edadesSeleccionarTodos
'              End If
'70            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
'              '
'80            lcMes = Right("0" & Trim(Str(lnIdMesElegido)), 2)
'              'lcAnio = cmbAnio.Text
'90            lcFec1 = ("01/" & lcMes & "/" & lcAnio)
'100           lnUltimoDiaMes = DevuelveUltimoDiaDelMes(Val(lcMes), Val(lcAnio))
'110           lcFec2 = Right("0" & Trim(Str(lnUltimoDiaMes)), 2) & "/" & lcMes & "/" & lcAnio
'120           lnMaxPacientesXhoja = Val(lcBuscaParametro.SeleccionaFilaParametro(272))
'130           lnMaxHISxLOTE = Val(lcBuscaParametro.SeleccionaFilaParametro(271))
'140           If lbExportaAnovafis = True Then
'                 '***********exporta al NOVAFIS ****************
'150              lcCod_dig = Val(lcBuscaParametro.SeleccionaFilaParametro(270))
'160           Else
'                 '***********exporta al HIS ****************
'170              lcCod_dig = Val(Right(lcBuscaParametro.SeleccionaFilaParametro(270), 3))
'180           End If
'190           lbEsCentroSalud = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
'              '
'200           oConexion.CommandTimeout = 300
'210           oConexion.CursorLocation = adUseClient
'220           oConexion.Open sighentidades.CadenaConexion
'              '
'230           oConexionFox.CommandTimeout = 300
'240           oConexionFox.Open "DSN=his"
'
'              '
'250           mo_reglasComunes.PreparaTablasDBFhis oRsFox1, oRsFox2, oRsFox3, oConexionFox
'260           If lbExportaAnovafis = True Then
'                 '***********exporta al NOVAFIS ****************
'270              Set oRsFox2 = Nothing
'280              With oRsFox2
'290                   .Fields.Append "COD_ESTAB", adVarChar, 9, adFldIsNullable
'300                   .Fields.Append "ANO", adInteger
'310                   .Fields.Append "MES", adInteger
'320                   .Fields.Append "NOM_LOTE", adVarChar, 3, adFldIsNullable
'330                   .Fields.Append "tot_pag", adInteger
'340                   .CursorType = adOpenDynamic
'350                   .LockType = adLockOptimistic
'360                   .Open
'370              End With
'380              Set oRsFox1 = Nothing
'390              With oRsFox1
'400                   .Fields.Append "COD_ESTAB", adVarChar, 9, adFldIsNullable
'410                   .Fields.Append "ANO", adInteger
'420                   .Fields.Append "MES", adInteger
'430                   .Fields.Append "NOM_LOTE", adVarChar, 3, adFldIsNullable
'440                   .Fields.Append "NUM_PAG", adInteger
'450                   .Fields.Append "NUM_REG", adInteger
'460                   .Fields.Append "DIA", adInteger
'470                   .Fields.Append "FICHAFAM", adVarChar, 10, adFldIsNullable
'480                   .Fields.Append "UBIGEO", adVarChar, 6, adFldIsNullable
'490                   .Fields.Append "EDAD", adInteger
'500                   .Fields.Append "TIP_EDAD", adVarChar, 1, adFldIsNullable
'510                   .Fields.Append "SEXO", adVarChar, 1, adFldIsNullable
'520                   .Fields.Append "ESTABLEC", adVarChar, 1, adFldIsNullable
'530                   .Fields.Append "SERVICIO", adVarChar, 1, adFldIsNullable
'540                   .Fields.Append "TD1", adVarChar, 1, adFldIsNullable
'550                   .Fields.Append "LC1", adVarChar, 3, adFldIsNullable
'560                   .Fields.Append "DX1", adVarChar, 6, adFldIsNullable
'570                   .Fields.Append "TD2", adVarChar, 1, adFldIsNullable
'580                   .Fields.Append "LC2", adVarChar, 3, adFldIsNullable
'590                   .Fields.Append "DX2", adVarChar, 6, adFldIsNullable
'600                   .Fields.Append "TD3", adVarChar, 1, adFldIsNullable
'610                   .Fields.Append "LC3", adVarChar, 3, adFldIsNullable
'620                   .Fields.Append "DX3", adVarChar, 6, adFldIsNullable
'630                   .Fields.Append "TD4", adVarChar, 1, adFldIsNullable
'640                   .Fields.Append "LC4", adVarChar, 3, adFldIsNullable
'650                   .Fields.Append "DX4", adVarChar, 6, adFldIsNullable
'660                   .Fields.Append "TD5", adVarChar, 1, adFldIsNullable
'670                   .Fields.Append "LC5", adVarChar, 3, adFldIsNullable
'680                   .Fields.Append "DX5", adVarChar, 6, adFldIsNullable
'690                   .Fields.Append "TD6", adVarChar, 1, adFldIsNullable
'700                   .Fields.Append "LC6", adVarChar, 3, adFldIsNullable
'710                   .Fields.Append "DX6", adVarChar, 6, adFldIsNullable
'720                   .Fields.Append "CODPSAL", adVarChar, 11, adFldIsNullable
'730                   .Fields.Append "DIGITA", adVarChar, 11, adFldIsNullable
'740                   .Fields.Append "COD_SERVSA", adVarChar, 6, adFldIsNullable
'750                   .Fields.Append "MT", adVarChar, 1, adFldIsNullable
'760                   .Fields.Append "CF", adVarChar, 2, adFldIsNullable
'770                   .Fields.Append "ET", adVarChar, 3, adFldIsNullable
'780                   .Fields.Append "PAIS", adVarChar, 3, adFldIsNullable
'790                   .Fields.Append "COD_DOC", adVarChar, 1, adFldIsNullable
'800                   .Fields.Append "FAMILIA", adVarChar, 2, adFldIsNullable
'810                   .Fields.Append "DNI", adVarChar, 11, adFldIsNullable
'820                   .Fields.Append "EST", adVarChar, 3, adFldIsNullable
'830                   .Fields.Append "ED", adInteger
'840                   .Fields.Append "ID", adVarChar, 30, adFldIsNullable
'850                   .CursorType = adOpenDynamic
'860                   .LockType = adLockOptimistic
'870                   .Open
'880                   .Sort = "cod_estab,ano,mes,nom_lote"
'890              End With
'900           End If
'              '
'910           CrearTablaTemporalDxCPT 'Frank 12092014
'              '
'920           Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion(lcMes, lcAnio, oConexion)
'      'rsReporte.Filter = "HISfuenteFinanciamiento='1'"
'930           lnTotal = rsReporte.RecordCount
'940           If lnTotal = 0 Then
'950              MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
'960           Else
'970              mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
'980              lnTotal = 0
'990              lcCod_2000 = Right("00000000000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 9)
'1000             ms_MensajeError = ""
'
'1010             rsReporte.MoveFirst
'1020             Do While Not rsReporte.EOF
'1030                  If IsNull(rsReporte.Fields!loteHis) Then
'1040                     ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
'1050                  ElseIf lbExportaAnovafis = True And UCase(Left(rsReporte!loteHis, 1)) = "V" Then
'1060                     ms_MensajeError = ms_MensajeError & "No puede empezar con la letra   V     el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
'1070                  ElseIf lbExportaAnovafis = True And InStr("0123456789", Left(rsReporte!loteHis, 1)) > 0 Then
'1080                     ms_MensajeError = ms_MensajeError & "No puede empezar con   NUMERO  el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
'1090                  End If
'1100                  lcMt = IIf(rsReporte.Fields!HoraIngreso >= lcTarde, "T", "M")
'1110                  lcNomLote = rsReporte.Fields!loteHis & lcCharInicio    'rsReporte.Fields!LoteHis & lcMt
'1120                  lnIdEmpleado = rsReporte.Fields!IdEmpleado
'1130                  lnTTot_pag = 0: lnTTot_reg = 0
'1140                  If IsNull(rsReporte.Fields!codigoServicioHIS) Then
'1150                     MsgBox "No se pudo procesar" & Chr(13) & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio
'1160                     Exit Do
'1170                  End If
'1180                  lbAumentoPagina = False
'1190                  Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado
'                          '
'1200                      lcMt = IIf(rsReporte.Fields!HoraIngreso >= lcTarde, "T", "M")
'1210                      lnAsc = Asc(lcCharInicio)
'                          lnAscFinal1 = 90: lnAscFinal2 = 90
'1220                      Do While True
'1230                         If InStr("01234567890ABCDEFGHIJKLMNOPQRSTUVWYZabcdefghijklmnopqrstuvwyz", Right(lcNomLote, 1)) = 0 Then
'1240                            If Second(Time) <= 10 Then
'1250                               lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
'1260                            ElseIf Second(Time) > 10 And Second(Time) <= 20 Then
'1270                               lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
'1280                            ElseIf Second(Time) > 20 And Second(Time) <= 30 Then
'1290                               lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
'1300                            Else
'1310                               lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
'1320                            End If
'1330                            lnAsc = lnAsc + 1
'1340                            If lnAsc >= 123 Then
'1350                               lnAsc = Asc(lcCharInicio)
'1360                            End If
'1370                         Else
'1380                              On Error Resume Next
'1390                              mo_reglasComunes.HisLoteXfiltro oRsFox4, oConexionFox, "ano= " & lcAnio & " and mes=" & lcMes & " and nom_lote='" & lcNomLote & "'"
'1400                              If oRsFox4.RecordCount = 0 Then
'1410                                 On Error GoTo ErrorProceso
'1420                                 Exit Do
'1430                              Else
'1440                                 lcSql = " "
'1450                                  If Second(Time) <= 10 Then
'1460                                     lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
'1470                                  ElseIf Second(Time) > 10 And Second(Time) <= 20 Then
'1480                                     lcNomLote = UCase(Left(lcNomLote, 2)) & Chr(lnAsc)
'1490                                  ElseIf Second(Time) > 20 And Second(Time) <= 30 Then
'1500                                     lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
'1510                                  Else
'1520                                     lcNomLote = UCase(Mid(lcNomLote, 1, 1)) & UCase(Mid(lcNomLote, 2, 1)) & Chr(lnAsc)
'1530                                  End If
'1540                                  lnAsc = lnAsc + 1
'1550                                  If lnAsc >= 123 Then
'
'                                            '*************
'                                            lcNomLote = Chr(lnAscFinal1) & Chr(lnAscFinal2)
'                                            lnAscFinal2 = lnAscFinal2 - 1
'                                            If lnAscFinal2 = 65 Then
'                                               lnAscFinal1 = lnAscFinal1 - 1
'                                               lnAscFinal2 = 90
'                                            End If
'                                             '******
'
'                                            lnAsc = Asc(lcCharInicio)
'1570                                  End If
'1580                              End If
'1590                              On Error GoTo ErrorProceso
'1600                         End If
'1610                      Loop
'                          '
'1620                      lnTTot_pag = 0: lnTTot_reg = 0
'                          '
'1630                      ldFecha = rsReporte.Fields!FechaIngreso
'                          '
'1640                      If IsNull(rsReporte.Fields!codigoServicioHIS) Then
'1650                          ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsReporte.Fields!DServicio & Chr(13)
'1660                          MsgBox ms_MensajeError, vbInformation, "Mensajes"
'1670                          Exit Sub
'1680                          lcCodigoServicioHIS = ""
'1690                      Else
'1700                          lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS
'1710                      End If
'1720                      lcCod_servsa = lcCodigoServicioHIS
'1730                      Set oRsImpHIStodos = mo_ReglasAdmision.ServiciosAtenSimultaneaImpHISxUPS(lcCodigoServicioHIS)
'1740                      lnoRsImpHIStodos = oRsImpHIStodos.RecordCount
'
'                          '
'1750                      lnDNImedico1 = "00000000"
'1760                      If IsNull(rsReporte.Fields!DNI) Or IsNull(rsReporte.Fields!idColegioHIS) Then
'1770                          ms_MensajeError = ms_MensajeError & "El profesional de Salud : " & rsReporte.Fields!dEmpleado & " deberá tener 'Colegio Médico' y 'DNI'" & Chr(13)
'1780                          lcCodif = ""
'1790                      Else
'1800                          Set oRsTmp = mo_ReglasAdmision.TiposDocIdentidadXfiltro("idDocIdentidad=" & rsReporte.Fields!idTipoDocumento, oConexion)
'1810                          lcTipoDocumentoHisDoctor = "1"
'1820                          If oRsTmp.RecordCount > 0 Then
'1830                             lcTipoDocumentoHisDoctor = oRsTmp.Fields!CodigoHIS
'1840                          End If
'1850                          oRsTmp.Close
'1860                          lcCodif = lcTipoDocumentoHisDoctor & Left(rsReporte.Fields!DNI, 8) & rsReporte.Fields!idColegioHIS
'1870                          lnDNImedico1 = Left(rsReporte.Fields!DNI, 8)
'1880                      End If
'                          '
'1890                      lcPlaza = lcCodif
'                          '
'                          '
'1900                      If lbAumentoPagina = True Then
'1910                         lnNumPag = lnNumPag + 1
'1920                      Else
'1930                         lnNumPag = 1
'1940                      End If
'1950                      lnTot_reg = 0: lnReg = 0
'1960                      Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso
'1970                          lcFuenteFinanciamientoPaciente = IIf(IsNull(rsReporte.Fields!HISfuenteFinanciamiento), "", rsReporte.Fields!HISfuenteFinanciamiento)
'1980                          lcEtniaPaciente = IIf(IsNull(rsReporte.Fields!IdEtnia), "", rsReporte.Fields!IdEtnia)
'                              '1112333333334455
'1990                          lcDNIpaciente = Space$(3): lcPaisPaciente = Space$(3)                     '1)Pais de Procedencia
'2000                          If Not IsNull(rsReporte.Fields!IdPaisProcedencia) Then
'2010                              Set oRsTmp = mo_ReglasServGeograf.PaisesXfiltro("idPais=" & rsReporte.Fields!IdPaisProcedencia)
'2020                              If oRsTmp.RecordCount > 0 Then
'2030                                 lcDNIpaciente = Right(Space$(3) & oRsTmp.Fields!Codigo, 3)
'2040                                 lcPaisPaciente = Right(Space$(3) & oRsTmp.Fields!Codigo, 3)
'2050                              End If
'2060                              oRsTmp.Close
'2070                          End If
'2080                          lcTipoDocumentoPaciente = Space$(1)
'2090                          If Not IsNull(rsReporte.Fields!HIStipoDocumento) Then    '2)Tipo de Documento de Identidad
'2100                             lcDNIpaciente = lcDNIpaciente & rsReporte.Fields!HIStipoDocumento
'2110                             lcTipoDocumentoPaciente = rsReporte.Fields!HIStipoDocumento
'2120                          Else
'2130                             lcDNIpaciente = lcDNIpaciente & IIf(IsNull(rsReporte.Fields!NroDocumento), Space$(1), "1")
'2140                          End If
'2150                          lcDNIpaciente1 = "00000000"
'2160                          If Not IsNull(rsReporte.Fields!NroDocumento) Then          '3)Numero de Documento de identidad
'2170                             lcDNIpaciente = lcDNIpaciente & Right(Trim(rsReporte.Fields!NroDocumento), 8)
'                                    If Val(lcTipoDocumentoPaciente) > 0 Then 'solo novafis
'2180                                  lcDNIpaciente1 = Left(rsReporte.Fields!NroDocumento, 8)
'                                    End If
'2190                          Else
'2200                             lcDNIpaciente = lcDNIpaciente & Space$(8)
'2210                          End If
'2220                          lcDNIpaciente = lcDNIpaciente & Space$(4)                  '4)CUATRO espacios en BLANCO
'                                                                                        '5)Numero de Hijos
'2230                          Set oRsTmp = mo_ReglasAdmision.atencionesDatosAdicionalesXfiltro("idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
'2240                          If oRsTmp.RecordCount > 0 Then
'2250                             If Not IsNull(oRsTmp.Fields!NumeroDeHijos) Then
'2260                                 lcDNIpaciente = lcDNIpaciente & Right("0" & Trim(Str(oRsTmp.Fields!NumeroDeHijos)), 2)
'2270                             Else
'2280                                 lcDNIpaciente = lcDNIpaciente & "00"
'2290                             End If
'2300                          Else
'2310                             lcDNIpaciente = lcDNIpaciente & "00"
'2320                          End If
'2330                          oRsTmp.Close
'
'                              '
'2340                          If rsReporte.Fields!NroHistoriaClinica = 511302 Then
'2350                          lcFichaOhistoria = ""
'2360                          End If
'
'2370                          If IsNull(rsReporte.Fields!NroHistoriaClinica) And IsNull(rsReporte.Fields!FichaFamiliar) Then
'2380                             MsgBox "El Paciente No existe para la CUENTA N° " & rsReporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
'2390                             Exit Sub
'2400                          End If
'2410                          lcFichaOhistoria = Trim(Str(rsReporte.Fields!NroHistoriaClinica))
'
'2420                          If Not IsNull(rsReporte.Fields!FichaFamiliar) And lbEsCentroSalud = True Then
'2430                             If Len(Trim(rsReporte.Fields!FichaFamiliar)) > 2 Then
'2440                                lcFichaOhistoria = Trim(rsReporte.Fields!FichaFamiliar)
'2450                             End If
'2460                          End If
'                              '
'2470                          lnIndiceDxCPT = 0
'2480                          If lnoRsImpHIStodos = 0 Then
'2490                              LimpiarTablaDxCpt
'2500                              If oRsTmp.State = 1 Then oRsTmp.Close
'2510                              Set oRsTmp = mo_ReglasAdmision.AtencionesDiagnosticosXidAtencion(rsReporte.Fields!idAtencion, oConexion)
'2520                              If oRsTmp.RecordCount > 0 Then
'2530                                 oRsTmp.MoveFirst
'2540                                 Do While Not oRsTmp.EOF
'2550                                    If IsNull(oRsTmp.Fields!DescripcionTipoDx) Then
'2560                                       MsgBox "No existe DIAGNOSTICO para la CUENTA N° " & rsReporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
'2570                                       Exit Sub
'2580                                    End If
'2590                                    lnIndiceDxCPT = lnIndiceDxCPT + 1
'2600                                    oRsDxCpt.AddNew
'2610                                    oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
'2620                                    oRsDxCpt.Fields!dx = sighentidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6))
'2630                                    oRsDxCpt.Fields!TDx = IIf(IsNull(oRsTmp.Fields!DescripcionTipoDx), "D", Left(oRsTmp.Fields!DescripcionTipoDx, 1))
'2640                                    oRsDxCpt.Fields!CDx = IIf(IsNull(oRsTmp.Fields!ClaseDxHIS), "1", Left(oRsTmp.Fields!ClaseDxHIS, 1))
'2650                                    oRsDxCpt.Fields!LDx = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
'2660                                    oRsDxCpt.Update
'2670                                    oRsTmp.MoveNext
'2680                                 Loop
'2690                              End If
'2700                              oRsTmp.Close
'                                  'CPT
'2710                              If lbAgregaOrdenesMedicas = True Or lnChkExportaCPT = 1 Then
'2720                                   If oRsTmp.State = 1 Then oRsTmp.Close
'2730                                   Set oRsTmp = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte.Fields!idCuentaAtencion)
'                                       If lbAgregaOrdenesMedicas = False Then
'                                          oRsTmp.Filter = "idPuntoCarga=1"
'                                       End If
'2740                                   If oRsTmp.RecordCount > 0 Then
'2750                                        oRsTmp.MoveFirst
'2760                                        Do While Not oRsTmp.EOF
'2770                                          lnIndiceDxCPT = lnIndiceDxCPT + 1
'2780                                          oRsDxCpt.AddNew
'2790                                          oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
'2800                                          oRsDxCpt.Fields!dx = oRsTmp.Fields("codigo").Value
'2810                                          oRsDxCpt.Fields!TDx = "D"
'2820                                          oRsDxCpt.Fields!CDx = ""
'2830                                          oRsDxCpt.Fields!LDx = oRsTmp.Fields("labConfHIS").Value
'2840                                          oRsDxCpt.Update
'2850                                          oRsTmp.MoveNext
'2860                                        Loop
'2870                                   End If
'2880                                   oRsTmp.Close
'2890                              End If
'2900                          Else
'
'2910                              If rsDx1.State = 1 Then rsDx1.Close
'2920                              Set rsDx1 = mo_ReglasAdmision.BuscaAtencionesDxCEparaFormatoHIS(rsReporte.Fields!idAtencion)
'2930                              If lbAgregaOrdenesMedicas = True Or lnChkExportaCPT = 1 Then
'                                       'CPT
'2940                                   If oRsTmp.State = 1 Then oRsTmp.Close
'2950                                   Set oRsTmp = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsReporte.Fields!idCuentaAtencion)
'                                       If lbAgregaOrdenesMedicas = False Then
'                                          oRsTmp.Filter = "idPuntoCarga=1"
'                                       End If
'
'2960                              End If
'2970                              mo_ReglasAdmision.CreaTemporalYcargaDxYcptXgrupo oRsDxCpt, rsDx1, oRsTmp, _
'                                                                                  rsReporte!idAtencion, False, lcAD040
'2980                              lnIndiceDxCPT = oRsDxCpt.RecordCount
'2990                          End If
'                              '
'3000                          lbContinuar = True
'3010                          If oRsDxCpt.RecordCount = 0 Then
'3020                              lbContinuar = False
'3030                          End If
'                              '
'3040                          If lbContinuar = True Then
'3050                              lcDistritoDomicilio = Space(6)
'3060                              If Not IsNull(rsReporte.Fields!IdDistrito) Then
'3070                                 lcDistritoDomicilio = Right("0" & Trim(Str(rsReporte.Fields!IdDistrito)), 6)
'3080                              End If
'                                  '
'3090                              lnTot_reg = lnTot_reg + 1
'3100                              lnReg = lnReg + 1
'3110                              lnEdadEnAtencion = rsReporte.Fields!Edad
'3120                              If lnEdadEnAtencion > 99 Then
'3130                                 lnEdadEnAtencion = 99
'3140                              End If
'3150                              lnGrupo6DxCPT = IIf(Round(oRsDxCpt.RecordCount / 6) < oRsDxCpt.RecordCount / 6, Round(oRsDxCpt.RecordCount / 6) + 1, Round(oRsDxCpt.RecordCount / 6))
'3160                              lnRestoDxCPT = IIf(oRsDxCpt.RecordCount Mod 6 = 0, 6, oRsDxCpt.RecordCount Mod 6)
'
'3170                              For lnContador = 0 To lnGrupo6DxCPT - 1
'3180                                  lnIndiceDxCPT = lnContador * 6
'3190                                  If lbExportaAnovafis = True Then
'                                          '***********exporta al NOVAFIS ****************
'3200                                      lnEdadEnTablaEdades = 0
'3210                                      oRsNovafis_edades.Filter = "tipo_edad='" & UCase(rsReporte.Fields!EdadCodigo) & "' and " & _
'                                                                     "edad=" & lnEdadEnAtencion
'3220                                      If oRsNovafis_edades.RecordCount > 0 Then
'3230                                         lnEdadEnTablaEdades = oRsNovafis_edades!ed
'3240                                      End If
'                                          '
'3250                                      oRsFox1.AddNew
'3260                                      oRsFox1.Fields!cod_estab = lcCod_2000
'3270                                      oRsFox1.Fields!ano = Val(lcAnio)
'3280                                      oRsFox1.Fields!Mes = Val(lcMes)
'3290                                      oRsFox1.Fields!nom_lote = lcNomLote
'3300                                      oRsFox1.Fields!Num_pag = lnNumPag
'3310                                      oRsFox1.Fields!num_reg = lnReg
'3320                                      oRsFox1.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
'3330                                      oRsFox1.Fields!fichafam = lcFichaOhistoria
'3340                                      oRsFox1.Fields!Ubigeo = lcDistritoDomicilio
'3350                                      oRsFox1.Fields!Edad = lnEdadEnAtencion
'3360                                      oRsFox1.Fields!tip_edad = rsReporte.Fields!EdadCodigo
'3370                                      oRsFox1.Fields!Sexo = IIf(rsReporte.Fields!idTipoSexo = 1, "M", "F")
'3380                                      oRsFox1.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
'3390                                      oRsFox1.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
'                                          'Coloca por defecto los Dx vacios
'3400                                      oRsFox1.Fields!dx1 = ""
'3410                                      oRsFox1.Fields!lc1 = ""
'3420                                      oRsFox1.Fields!td1 = ""
'3430                                      oRsFox1.Fields!dx2 = ""
'3440                                      oRsFox1.Fields!lc2 = ""
'3450                                      oRsFox1.Fields!td2 = ""
'3460                                      oRsFox1.Fields!dx3 = ""
'3470                                      oRsFox1.Fields!lc3 = ""
'3480                                      oRsFox1.Fields!td3 = ""
'3490                                      oRsFox1.Fields!dx4 = ""
'3500                                      oRsFox1.Fields!lc4 = ""
'3510                                      oRsFox1.Fields!td4 = ""
'3520                                      oRsFox1.Fields!dx5 = ""
'3530                                      oRsFox1.Fields!lc5 = ""
'3540                                      oRsFox1.Fields!td5 = ""
'3550                                      oRsFox1.Fields!dx6 = ""
'3560                                      oRsFox1.Fields!lc6 = ""
'3570                                      oRsFox1.Fields!td6 = ""
'3580                                      lcGestante = "0"
'3590                                      For lnContDxCpt = 1 To 6
'3600                                          oRsDxCpt.MoveFirst
'3610                                          oRsDxCpt.Find "IdDxCpt=" & lnIndiceDxCPT + lnContDxCpt
'3620                                          If Not oRsDxCpt.EOF Then
'3630                                              Select Case lnContDxCpt
'                                                  Case 1
'3640                                                  oRsFox1.Fields!td1 = oRsDxCpt.Fields!TDx
'3650                                                  oRsFox1.Fields!lc1 = oRsDxCpt.Fields!LDx
'3660                                                  oRsFox1.Fields!dx1 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'3670                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'3680                                                     lcGestante = "G"
'3690                                                  End If
'3700                                              Case 2
'3710                                                  oRsFox1.Fields!td2 = oRsDxCpt.Fields!TDx
'3720                                                  oRsFox1.Fields!lc2 = oRsDxCpt.Fields!LDx
'3730                                                  oRsFox1.Fields!dx2 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'3740                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'3750                                                     lcGestante = "G"
'3760                                                  End If
'3770                                              Case 3
'3780                                                  oRsFox1.Fields!td3 = oRsDxCpt.Fields!TDx
'3790                                                  oRsFox1.Fields!lc3 = oRsDxCpt.Fields!LDx
'3800                                                  oRsFox1.Fields!dx3 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'3810                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'3820                                                     lcGestante = "G"
'3830                                                  End If
'3840                                              Case 4
'3850                                                  oRsFox1.Fields!td4 = oRsDxCpt.Fields!TDx
'3860                                                  oRsFox1.Fields!lc4 = oRsDxCpt.Fields!LDx
'3870                                                  oRsFox1.Fields!dx4 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'3880                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'3890                                                     lcGestante = "G"
'3900                                                  End If
'3910                                              Case 5
'3920                                                  oRsFox1.Fields!td5 = oRsDxCpt.Fields!TDx
'3930                                                  oRsFox1.Fields!lc5 = oRsDxCpt.Fields!LDx
'3940                                                  oRsFox1.Fields!dx5 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'3950                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'3960                                                     lcGestante = "G"
'3970                                                  End If
'3980                                              Case 6
'3990                                                  oRsFox1.Fields!td6 = oRsDxCpt.Fields!TDx
'4000                                                  oRsFox1.Fields!lc6 = oRsDxCpt.Fields!LDx
'4010                                                  oRsFox1.Fields!dx6 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4020                                                  If mo_reglasComunes.novafis_ElDxEsDeEmbarazada(oRsDxCpt!dx, rsReporte.Fields!idTipoSexo) = True Then
'4030                                                     lcGestante = "G"
'4040                                                  End If
'4050                                              End Select
'4060                                          End If
'4070                                      Next
'4080                                      oRsFox1.Fields!codpsal = lnDNImedico1
'4090                                      oRsFox1.Fields!digita = Right("0" & lcCod_dig, 8)
'4100                                      oRsFox1.Fields!cod_servsa = lcCod_servsa
'4110                                      oRsFox1.Fields!mt = lcMt
'4120                                      oRsFox1.Fields!cf = Left(lcFuenteFinanciamientoPaciente & " ", 2)
'4130                                      oRsFox1.Fields!et = lcEtniaPaciente
'4140                                      oRsFox1.Fields!pais = lcPaisPaciente
'4150                                      oRsFox1.Fields!cod_doc = lcTipoDocumentoPaciente
'4160                                      oRsFox1.Fields!familia = "00"
'4170                                      oRsFox1.Fields!DNI = lcDNIpaciente1
'4180                                      oRsFox1.Fields!est = lcGestante
'4190                                      oRsFox1.Fields!ed = lnEdadEnTablaEdades
'4200                                      oRsFox1.Fields!Id = Space(30)
'4210                                      lnEdadEnTablaEdades = 0
'4220                                  Else
'                                          '***********exporta al HIS ****************
'4230                                      oRsFox1.AddNew
'4240                                      oRsFox1.Fields!Cod_2000 = lcCod_2000
'4250                                      oRsFox1.Fields!ano = Val(lcAnio)
'4260                                      oRsFox1.Fields!Mes = Val(lcMes)
'4270                                      oRsFox1.Fields!nom_lote = lcNomLote
'4280                                      oRsFox1.Fields!Num_pag = lnNumPag
'4290                                      oRsFox1.Fields!num_reg = lnReg
'4300                                      oRsFox1.Fields!dia = Day(rsReporte.Fields!FechaIngreso)
'4310                                      oRsFox1.Fields!fichafam = lcFichaOhistoria
'4320                                      oRsFox1.Fields!Cod_Dpto = Left(lcDistritoDomicilio, 2)
'4330                                      oRsFox1.Fields!Cod_Prov = Mid(lcDistritoDomicilio, 3, 2)
'4340                                      oRsFox1.Fields!Cod_Dist = Right(lcDistritoDomicilio, 2)
'4350                                      oRsFox1.Fields!Edad = lnEdadEnAtencion
'4360                                      oRsFox1.Fields!tip_edad = rsReporte.Fields!EdadCodigo
'4370                                      oRsFox1.Fields!Sexo = IIf(rsReporte.Fields!idTipoSexo = 1, "M", "F")
'4380                                      oRsFox1.Fields!establec = IIf(rsReporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
'4390                                      oRsFox1.Fields!Servicio = IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsReporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
'
'                                          'Coloca por defecto los Dx vacios
'4400                                      oRsFox1.Fields!diagnost1 = ""
'4410                                      oRsFox1.Fields!labconf1 = ""
'4420                                      oRsFox1.Fields!codigo1 = ""
'4430                                      oRsFox1.Fields!diagnost2 = ""
'4440                                      oRsFox1.Fields!labconf2 = ""
'4450                                      oRsFox1.Fields!codigo2 = ""
'4460                                      oRsFox1.Fields!diagnost3 = ""
'4470                                      oRsFox1.Fields!labconf3 = ""
'4480                                      oRsFox1.Fields!codigo3 = ""
'4490                                      oRsFox1.Fields!diagnost4 = ""
'4500                                      oRsFox1.Fields!labconf4 = ""
'4510                                      oRsFox1.Fields!codigo4 = ""
'4520                                      oRsFox1.Fields!diagnost5 = ""
'4530                                      oRsFox1.Fields!labconf5 = ""
'4540                                      oRsFox1.Fields!codigo5 = ""
'4550                                      oRsFox1.Fields!diagnost6 = ""
'4560                                      oRsFox1.Fields!labconf6 = ""
'4570                                      oRsFox1.Fields!codigo6 = ""
'
'4580                                      For lnContDxCpt = 1 To 6
'4590                                          oRsDxCpt.MoveFirst
'4600                                          oRsDxCpt.Find "IdDxCpt=" & lnIndiceDxCPT + lnContDxCpt
'4610                                          If Not oRsDxCpt.EOF Then
'4620                                              Select Case lnContDxCpt
'                                                  Case 1
'4630                                                  oRsFox1.Fields!diagnost1 = oRsDxCpt.Fields!TDx
'4640                                                  oRsFox1.Fields!labconf1 = oRsDxCpt.Fields!LDx
'4650                                                  oRsFox1.Fields!codigo1 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4660                                              Case 2
'4670                                                  oRsFox1.Fields!diagnost2 = oRsDxCpt.Fields!TDx
'4680                                                  oRsFox1.Fields!labconf2 = oRsDxCpt.Fields!LDx
'4690                                                  oRsFox1.Fields!codigo2 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4700                                              Case 3
'4710                                                  oRsFox1.Fields!diagnost3 = oRsDxCpt.Fields!TDx
'4720                                                  oRsFox1.Fields!labconf3 = oRsDxCpt.Fields!LDx
'4730                                                  oRsFox1.Fields!codigo3 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4740                                              Case 4
'4750                                                  oRsFox1.Fields!diagnost4 = oRsDxCpt.Fields!TDx
'4760                                                  oRsFox1.Fields!labconf4 = oRsDxCpt.Fields!LDx
'4770                                                  oRsFox1.Fields!codigo4 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4780                                              Case 5
'4790                                                  oRsFox1.Fields!diagnost5 = oRsDxCpt.Fields!TDx
'4800                                                  oRsFox1.Fields!labconf5 = oRsDxCpt.Fields!LDx
'4810                                                  oRsFox1.Fields!codigo5 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4820                                              Case 6
'4830                                                  oRsFox1.Fields!diagnost6 = oRsDxCpt.Fields!TDx
'4840                                                  oRsFox1.Fields!labconf6 = oRsDxCpt.Fields!LDx
'4850                                                  oRsFox1.Fields!codigo6 = IIf(Len(oRsDxCpt.Fields!dx) > 6, Left(oRsDxCpt.Fields!dx, 6), oRsDxCpt.Fields!dx)
'4860                                              End Select
'4870                                          End If
'4880                                      Next
'4890                                      oRsFox1.Fields!esta_reg = "2"
'4900                                      oRsFox1.Fields!mt = lcMt
'4910                                      oRsFox1.Fields!DNI = lcDNIpaciente
'4920                                      oRsFox1.Fields!FI = lcFuenteFinanciamientoPaciente
'4930                                      oRsFox1.Fields!et = lcEtniaPaciente
'4940                                      oRsFox1.Fields!ST = "1"
'4950                                      oRsFox1.Update
'4960                                  End If
'4970                              Next
'                                  '
'4980                          End If
'4990                          lnTotal = lnTotal + 1
'                              oRsFox5.Open "update parametros set valorint=" & Trim(Str(lnTotal)) & " where idparametro=2", oConexion, adOpenKeyset, adLockOptimistic
'                              ml_ValorActual = Trim(Str(lnTotal))
'5000                          mo_ProgressRpt1.Value = lnTotal
'5010                          DoEvents
'If lnTotal = 1719 Then
'ml_ValorActual = 0
'End If
'
'5020                          rsReporte.MoveNext
'5030                          If rsReporte.EOF Then
'5040                             Exit Do
'5050                          End If
'5060                          If lnReg >= lnMaxPacientesXhoja And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoServicioHIS And ldFecha = rsReporte.Fields!FechaIngreso Then
'5070                             lbAumentoPagina = True
'5080                             lnReg = 0
'5090                             Exit Do
'5100                          Else
'5110                             lbAumentoPagina = False
'5120                          End If
'5130                      Loop
'5140                      If lnTot_reg > 0 Then
'5150                          If lbExportaAnovafis = True Then
'                                  '***********exporta al NOVAFIS ****************
'5160                              lbNuevoRegistro = True
'5170                              If oRsFox2.RecordCount > 0 Then
'5180                                 oRsFox2.MoveFirst
'5190                                 Do While Not oRsFox2.EOF
'5200                                    If oRsFox2!cod_estab = lcCod_2000 And oRsFox2!ano = Val(lcAnio) And _
'                                                          oRsFox2!Mes = Val(lcMes) And oRsFox2!nom_lote = lcNomLote Then
'5210                                       lbNuevoRegistro = False
'5220                                       Exit Do
'5230                                    End If
'5240                                    oRsFox2.MoveNext
'5250                                 Loop
'5260                              End If
'5270                              If lbNuevoRegistro = True Then
'5280                                  oRsFox2.AddNew
'5290                                  oRsFox2.Fields!cod_estab = lcCod_2000
'5300                                  oRsFox2.Fields!ano = Val(lcAnio)
'5310                                  oRsFox2.Fields!Mes = Val(lcMes)
'5320                                  oRsFox2.Fields!nom_lote = lcNomLote
'5330                                  oRsFox2.Fields!tot_pag = lnNumPag
'5340                              Else
'5350                                  oRsFox2.Fields!tot_pag = oRsFox2.Fields!tot_pag + lnNumPag
'5360                              End If
'5370                          Else
'                                  '***********exporta al HIS ****************
'5380                              oRsFox2.AddNew
'5390                              oRsFox2.Fields!Cod_2000 = lcCod_2000
'5400                              oRsFox2.Fields!ano = Val(lcAnio)
'5410                              oRsFox2.Fields!Mes = Val(lcMes)
'5420                              oRsFox2.Fields!nom_lote = lcNomLote
'5430                              oRsFox2.Fields!Num_pag = lnNumPag
'5440                              oRsFox2.Fields!Codif = lcCodif
'5450                              oRsFox2.Fields!cod_servsa = lcCod_servsa
'5460                              oRsFox2.Fields!plaza = lcCodif
'5470                              oRsFox2.Fields!Esta_pag = "2"
'5480                              oRsFox2.Fields!Tot_reg = lnTot_reg
'5490                              oRsFox2.Fields!FlagEnvio = Space(1)
'5500                              oRsFox2.Fields!mt = lcMt
'5510                              oRsFox2.Fields!ST = "1"
'5520                              oRsFox2.Update
'5530                          End If
'                              '
'5540                          lnTTot_pag = lnTTot_pag + lnNumPag
'5550                          lnTTot_reg = lnTTot_reg + lnTot_reg
'                              '
'5560                          oRsFox3.AddNew
'5570                          oRsFox3.Fields!Cod_2000 = lcCod_2000
'5580                          oRsFox3.Fields!ano = Val(lcAnio)
'5590                          oRsFox3.Fields!Mes = Val(lcMes)
'5600                          oRsFox3.Fields!nom_lote = lcNomLote
'5610                          If lbExportaAnovafis = True Then
'                                  '***********exporta al NOVAFIS ****************
'5620                              oRsFox3.Fields!cod_dig = "1"
'5630                          Else
'                                  '***********exporta al HIS ****************
'5640                              oRsFox3.Fields!cod_dig = lcCod_dig
'5650                          End If
'5660                          oRsFox3.Fields!esta_lote = "LC"
'5670                          oRsFox3.Fields!tot_pag = lnTTot_pag
'5680                          oRsFox3.Fields!Tot_reg = lnTTot_reg
'5690                          oRsFox3.Fields!Tot_rmalos = 0
'5700                          oRsFox3.Fields!tot_pgcarg = 0
'5710                          oRsFox3.Fields!esta_local = Space(1)
'5720                          oRsFox3.Fields!nomfile = Space(14)
'5730                          oRsFox3.Fields!flagReport = "3"
'5740                          oRsFox3.Fields!mt = lcMt
'5750                          oRsFox3.Update
'5760                      End If
'5770                      If rsReporte.EOF Then
'5780                         Exit Do
'5790                      End If
'5800                  Loop
'5810             Loop
'                 '
'5820             If lbExportaAnovafis = True Then
'                     '***********exporta al NOVAFIS ****************
'5830                 lnTTot_pag = oRsFox1.RecordCount
'5840                 If lnTTot_pag > 0 Then
'5850                      oRsFox1.MoveFirst
'5860                      Do While Not oRsFox1.EOF
'5870                         lnNumPag = 0
'5880                         oRsFox2.MoveFirst
'5890                         Do While Not oRsFox2.EOF
'5900                              If oRsFox2!cod_estab = oRsFox1!cod_estab And oRsFox2!ano = oRsFox1!ano And _
'                                                    oRsFox2!Mes = oRsFox1!Mes And oRsFox2!nom_lote = oRsFox1!nom_lote Then
'5910                                 lnNumPag = oRsFox2!tot_pag
'5920                                 Exit Do
'5930                              End If
'5940                              oRsFox2.MoveNext
'5950                         Loop
'
'5960                         lcMt = oRsFox1!cod_estab & Trim(Str(oRsFox1!ano)) & Right("0" & Trim(Str(oRsFox1!Mes)), 2) & _
'                                   oRsFox1!nom_lote & Right("00" & Trim(Str(lnNumPag)), 3) & Right("0" & oRsFox1!Num_pag, 2) & _
'                                   Right("0" & Trim(Str(oRsFox1!num_reg)), 2)
'5970                         oRsFox1!Id = lcMt
'5980                         oRsFox1.Update
'5990                         oRsFox1.MoveNext
'6000                      Loop
'                          '
'6010                      If ms_MensajeError = "" Then
'6020                          GeneraTXT_una_linea oRsFox1, lcRutaExportar & "envio.txt", lnTTot_pag, 1, Chr(9), Chr(10)
'6030                          Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al NOVAFIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
'6040                          MsgBox "Se exportó correctamente el archivo: " & lcRutaExportar & "envio.txt"
'6050                      End If
'6060                 End If
'6070             Else
'                     '***********exporta al HIS ****************
'6080                 Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al HIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
'6090             End If
'                 '
'
'6100          End If
'6110          If oRsTmp.State = 1 Then oRsTmp.Close
'6120          oRsFox1.Close
'6130          oRsFox2.Close
'6140          oRsFox3.Close
'6150          oConexionFox.Close
'6160          oConexion.Close
'6170          oConexionExterna.Close
'              If lbExportaAnovafis = True Then
'6180             oRsNovafis_edades.Close
'              End If
'6190          If ms_MensajeError <> "" Then
'6200             MsgBox ms_MensajeError, vbInformation, "Mensajes"
'6210          End If
'
'6220          Set mo_ReglasFarmacia = Nothing
'6230          Set mo_reglasComunes = Nothing
'6240          Set mo_ReglasSeguridad = Nothing
'6250          Set mo_ReglasAdmision = Nothing
'6260          Set mo_ReglasServGeograf = Nothing
'6270          Set oConexion = Nothing
'6280          Set oConexionFox = Nothing
'6290          Set rsReporte = Nothing
'6300          Set oRsTmp = Nothing
'6310          Set oRsFox1 = Nothing
'6320          Set oRsFox2 = Nothing
'6330          Set oRsFox3 = Nothing
'6340          Set oRsFox4 = Nothing
'6350          Set oRsEdadHis = Nothing
'6360          Set oRsImpHIStodos = Nothing
'6370          Set rsDx1 = Nothing
'6380          Set oConexionExterna = Nothing
'6390          Set oRsNovafis_edades = Nothing
'6400      Exit Sub
'
'ErrorProceso:
'6410      If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
'6420         Resume Next
'6430      Else
'6440         MsgBox Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
'                                                                             "Sub ExportaDAtosAlHISv4", "procesos.cls")  'debb-02/05/2016
'6450         Resume
'6460      End If
'
'End Sub

Function DevuelveDESDEdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "(" & Trim(Str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = "-" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveDESDEdeTablaEdad000 = Val(lcRango)
End Function

Function DevuelveHASTAdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "-" & Trim(Str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = ")" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveHASTAdeTablaEdad000 = Val(lcRango)
End Function






'********************************************************* SEM *****************************
Sub ExportaSEMversion2009()
'  If (txtFechaInicio.Text = "" Or Not IsDate(txtFechaInicio.Text)) Or (txtFechaFin.Text = "" Or Not IsDate(txtFechaFin.Text)) Or ssoEmergencias.Value = False And ssoHospitalizacion.Value = False Or txtNCorrelativo.Text = "" Then
'    MsgBox "Debe completar los datos requeridos", vbInformation, Me.Caption
'    txtFechaInicio.SetFocus
'    Exit Sub
'  End If
'
'  If MsgBox("Esta seguro de exportar datos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
'    Me.MousePointer = 11
'    Frame4.Enabled = False
'    btnAceptar.Enabled = False
'
'    'On Error GoTo ErrorProceso
'    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
'    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
'    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
'    Dim lcBuscaParametro As New SIGHDatos.Parametros
'    Dim rsReporte As New Recordset
'    Dim oRsTmp As New Recordset
'    Dim oRsFox As New Recordset
'    Dim oConexionFox As New Connection
'    Dim lcDx1 As String, lcDx2 As String
'    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
'    Dim lcFec1 As String, lcFec2 As String
'    Dim oConexion As New Connection
'        '
'    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
'        '
'    oConexionFox.CommandTimeout = 300
'    oConexionFox.Open "DSN=his"
'
'    lcSql = "SELECT *  FROM " & ml_Tabla & " WHERE month(fecIng)>=" & Month(CDate(txtFechaInicio.Text)) & " and month(fecIng)<=" & Month(CDate(txtFechaFin.Text))
'    oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'    If oRsFox.RecordCount > 0 Then
'      oRsFox.MoveFirst
'      Do While Not oRsFox.EOF
'        If oRsFox.Fields!fecIng >= CDate(txtFechaInicio.Text) And oRsFox.Fields!fecIng <= CDate(txtFechaFin.Text) Then
'          oRsFox.Delete
'          oRsFox.Update
'        End If
'        oRsFox.MoveNext
'      Loop
'    End If
'    oRsFox.Close
'
'    oRsFox.Open "SELECT * FROM " & ml_Tabla & " ORDER BY numMov", oConexionFox, adOpenKeyset, adLockOptimistic
'    If oRsFox.RecordCount = 0 Then
'      'txtNCorrelativo.Text = "1"
'    Else
'      oRsFox.MoveLast
'      txtNCorrelativo.Text = oRsFox.Fields!numMov + 1
'    End If
'    DoEvents
'
'    Set rsReporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(txtFechaInicio.Text), CDate(txtFechaFin.Text), "FECHAEGRESO")
'    rsReporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
'    If rsReporte.RecordCount = 0 Then
'      MsgBox "No existe Información  con esos Datos", vbInformation, Me.Caption
'    Else
'      lnNumMov = Val(txtNCorrelativo.Text)
'      rsReporte.MoveFirst
'      Do While Not rsReporte.EOF
'        'solo Hospitalizados
'        lcDx1 = "": lcDx2 = ""
'        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 3, oConexion)
'        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
'        If oRsTmp.RecordCount > 0 Then
'          '*****vivos
'          oRsTmp.MoveFirst
'          lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'          oRsTmp.Filter = "IdTipoDiagnostico=302"
'          If oRsTmp.RecordCount > 0 Then
'            oRsTmp.MoveFirst
'            lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'          End If
'        Else
'          '******fallecidos
'          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsReporte.Fields!idAtencion, 4, oConexion)
'          If oRsTmp.RecordCount > 0 Then
'                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'                End If
'                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
'                If oRsTmp.RecordCount > 0 Then
'                  oRsTmp.MoveFirst
'                  lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
'                End If
'          End If
'        End If
'
'        '
'        lnEstancia = DiasDeEstanciaEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
'        lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
'        If lnEstancia = 0 Then lnEstancia = 1
'        '
'        oRsFox.AddNew
'        oRsFox.Fields!numMov = lnNumMov
'        oRsFox.Fields!cod_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
'        oRsFox.Fields!cod_red = lcBuscaParametro.SeleccionaFilaParametro(240)
'        oRsFox.Fields!cod_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
'        oRsFox.Fields!cod_est = Trim(Str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
'        If Not IsNull(rsReporte.Fields!NroHistoriaClinica) Then
'          oRsFox.Fields!numhc = Left(Trim(Str(rsReporte.Fields!NroHistoriaClinica)), 7)
'        Else
'          oRsFox.Fields!numhc = ""
'        End If
'        oRsFox.Fields!apenom = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsReporte.Fields!SegundoNombre), "", Trim(rsReporte.Fields!SegundoNombre)), 50)
'        oRsFox.Fields!Sexo = Trim(Str(rsReporte.Fields!idTipoSexo))
'        oRsFox.Fields!Edad = rsReporte.Fields!Edad
'        oRsFox.Fields!TipoEdad = rsReporte.Fields!tipoEdadSEM
'        oRsFox.Fields!ubigeo = lcBuscaParametro.SeleccionaFilaParametro(242)
'        oRsFox.Fields!fecIng = rsReporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
'        oRsFox.Fields!fecegr = rsReporte.Fields!FechaEgreso    'Format(rsReporte.Fields!FechaEgreso, "mm/dd/yyyy")
'        oRsFox.Fields!totalest = lnEstancia
'        oRsFox.Fields!Servicio = IIf(IsNull(rsReporte.Fields!codigoServicioSEM), "xx", rsReporte.Fields!codigoServicioSEM)
'        oRsFox.Fields!condicion = rsReporte.Fields!codigoCondicionAltaSEM
'        oRsFox.Fields!seguro = rsReporte.Fields!CodigoFuenteFinanciamientoSEM
'        oRsFox.Fields!coddiag = lcDx1
'        oRsFox.Fields!coddiag2 = lcDx2
'        oRsFox.Fields!ubicacionh = IIf(IsNull(rsReporte.Fields!ubicacionSEM), "xxxxx", rsReporte.Fields!ubicacionSEM)
'        oRsFox.Fields!medicoh = IIf(IsNull(rsReporte.Fields!Colegiatura), "xxxxx", Left(Trim(rsReporte.Fields!Colegiatura), 5))
'        oRsFox.Fields!estado = 1
'        oRsFox.Fields!fechareg = Format(Date, "mm/dd/yyyy")
'        oRsFox.Update
'        lnNumMov = lnNumMov + 1
'        rsReporte.MoveNext
'      Loop
'      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM: " & txtFechaInicio.Text & " al " & txtFechaFin.Text)       '500+ ListBarReporte.idReporte
'    End If
'    oRsFox.Close
'    oConexionFox.Close
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oConexionFox = Nothing
'    Set oRsFox = Nothing
'    MsgBox "Exportación de datos finalizado correctamente" & Chr(13) & "Valor del Último NumMov: " & lnNumMov - 1, vbInformation, Me.Caption
'    Frame4.Enabled = True
'    btnAceptar.Enabled = True
'    txtFechaInicio.SetFocus
'    Me.MousePointer = 1
'    'Me.Visible = False
'    Exit Sub
'  End If
'  Exit Sub
'ErrorProceso:
'  MsgBox Err.Description, vbInformation, Me.Caption
'  'Resume
End Sub
Sub ExportaSEM(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, ml_Valor As Integer, ml_Tabla As String)
    On Error GoTo ErrorProceso
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    
    Dim rsreporte As New Recordset
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsFox As New Recordset
    Dim oConexionFox As New Connection
    Dim lcDx1 As String, lcDx2 As String
    Dim lnNumMov As Long, lcSql As String, lnEstancia As Integer
    Dim lcFec1 As String, lcFec2 As String
    Dim lcMotivo As String, lcDestinoAtencion As String, lcReferido As String, lcUsuario As String
    Dim lcDireccionDomicilio As String, lcNombreAcompaniante As String
    Dim lnEsObservacionEmergencia As Integer, lcServicio As String, lcMedico As String, lcServicioObs As String
    Dim lcMedicoObs As String, ldFechaIngObs As Date, lcHoraIngObs As String, lcCamaObs As String, ldFechaSalObs As Date
    Dim lcHoraSalObs As String, lcDxObs As String, lcHoraEstanciaMax As String, lnTotal As Long
    Dim oConexion As New Connection
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
        '
    oConexion.CommandTimeout = 300
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
        '
    oConexionFox.CommandTimeout = 300
    lcParaMsgbox = " (Eliminar Datos)"
    If ml_Valor = 2 Then
        '*********************************Emergencia
        
        oConexionFox.Open "DSN=SEMNUEVO"
        ml_Tabla = "emergencias"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!em_fecha >= CDate(lcFechaInicio) And oRsFox.Fields!em_fecha <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", "em_codigo"
        If oRsFox.RecordCount = 0 Then
          'lcCorrelativo = "1"
        Else
          oRsFox.MoveLast
          lcCorrelativo = oRsFox.Fields!em_codigo + 1
        End If
    Else
        '*********************************Hospitalizacion
        oConexionFox.Open "DSN=SEM"
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "month(fecIng)>=" & Month(CDate(lcFechaInicio)) & " and month(fecIng)<=" & Month(CDate(lcFechaFinal)), ""
        If oRsFox.RecordCount > 0 Then
          oRsFox.MoveFirst
          Do While Not oRsFox.EOF
            If oRsFox.Fields!fecIng >= CDate(lcFechaInicio) And oRsFox.Fields!fecIng <= CDate(lcFechaFinal) Then
              oRsFox.Delete
              oRsFox.Update
            End If
            oRsFox.MoveNext
          Loop
        End If
        oRsFox.Close
        '
        mo_ReglasComunes.AbreTablaFox oRsFox, oConexionFox, ml_Tabla, "", "numMov"
        If oRsFox.RecordCount = 0 Then
          'lcCorrelativo = "1"
        Else
          oRsFox.MoveLast
          lcCorrelativo = oRsFox.Fields!numMov + 1
        End If
    End If
 '   DoEvents
    '
    lcUsuario = " "
    Set oRsTmp1 = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario)
    If oRsTmp1.RecordCount > 0 Then
       If Not IsNull(oRsTmp1.Fields!Usuario) Then
          lcUsuario = oRsTmp1.Fields!Usuario
       End If
    End If
    oRsTmp1.Close
    '
    Set rsreporte = mo_ReglasAdmision.AtencionesSeleccionarHospEmergPorFechasDeEgresoMedico(CDate(lcFechaInicio), CDate(lcFechaFinal), "FECHAEGRESO")
    rsreporte.Filter = "codigoCondicionAltaSEM<>null and idTipoServicio=" & ml_Valor
    lnTotal = rsreporte.RecordCount
    If lnTotal = 0 Then
      MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
    Else
      mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotal + 2: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen:
      
      
      lnNumMov = Val(lcCorrelativo)
      rsreporte.MoveFirst
      Do While Not rsreporte.EOF
If rsreporte.Fields!idCuentaAtencion = 50829 Then
lcDx1 = ""
End If
        lcParaMsgbox = " -->Cuenta: " & rsreporte.Fields!idCuentaAtencion
        'solo Hospitalizados
        lcDx1 = "": lcDx2 = ""
        Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsreporte.Fields!idAtencion, 3, oConexion)
        oRsTmp.Filter = "IdTipoDiagnostico=301"   'vivos-princial
        If oRsTmp.RecordCount > 0 Then
          '*****vivos
          oRsTmp.MoveFirst
          lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
          oRsTmp.Filter = "IdTipoDiagnostico=302"
          If oRsTmp.RecordCount > 0 Then
            oRsTmp.MoveFirst
            lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
          End If
        Else
          '******fallecidos
          Set oRsTmp = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(rsreporte.Fields!idAtencion, 4, oConexion)
          If oRsTmp.RecordCount > 0 Then
                oRsTmp.Filter = "IdTipoDiagnostico=303"    'Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx1 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
                End If
                oRsTmp.Filter = "IdTipoDiagnostico=304 or IdTipoDiagnostico=305"    'fallecidos - Final
                If oRsTmp.RecordCount > 0 Then
                  oRsTmp.MoveFirst
                  lcDx2 = Left(oRsTmp.Fields!CodigoCIE2004, 5)
                End If
          End If
        End If
    
        'oRsTmp.Close
        '
        lcMotivo = " "
        Set oRsTmp1 = mo_AdminAdmision.AtencionesEmergenciaXidAtencion(rsreporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!motivoSEM) Then
              lcMotivo = oRsTmp1.Fields!motivoSEM
           End If
        End If
        oRsTmp1.Close
        '
        lcDireccionDomicilio = "": lcNombreAcompaniante = ""
        Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsreporte.Fields!idAtencion, oConexion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
              lcDireccionDomicilio = oRsTmp1.Fields!DireccionDomicilio
           End If
           If Not IsNull(oRsTmp1.Fields!NombreAcompaniante) Then
              lcNombreAcompaniante = oRsTmp1.Fields!NombreAcompaniante
           End If
        End If
        oRsTmp1.Close
        '
        lcDestinoAtencion = " "
        If Not IsNull(rsreporte.Fields!IdDestinoAtencion) Then
            Set oRsTmp1 = mo_ReglasAdmision.TiposDestinoAtencionXidDestinoAtencion(rsreporte.Fields!IdDestinoAtencion, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               If IsNull(oRsTmp1.Fields!destinoSEM) Then
                    MsgBox "Existen Servicios de Emergencia sin CODIGO SEM", vbInformation, "Mensaje"  'debb 05/12/2012
                    Exit Sub
               Else
                    lcDestinoAtencion = oRsTmp1.Fields!destinoSEM
               End If
            End If
            oRsTmp1.Close
        End If
        '
        lcReferido = " "
        If Not IsNull(rsreporte.Fields!idEstablecimientoDestino) Then
            Set oRsTmp1 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro("idEstablecimiento=" & rsreporte.Fields!idEstablecimientoDestino, oConexion)
            If oRsTmp1.RecordCount > 0 Then
               lcReferido = Left(Trim(Str(Val(oRsTmp1.Fields!Codigo))), 4)
            End If
            oRsTmp1.Close
        End If
        '
        lnEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsreporte!FechaIngreso, rsreporte!HoraIngreso, rsreporte!fechaEgreso, rsreporte!HoraEgreso, lcHoraEstanciaMax)
        If lnEstancia = 0 Then lnEstancia = 1
        '
        If ml_Valor = 2 Then
            '*********************************Emergencia
            lcServicio = IIf(IsNull(rsreporte.Fields!codigoServicioSEM), ".", rsreporte.Fields!codigoServicioSEM)
            lcMedico = IIf(IsNull(rsreporte.Fields!Colegiatura), ".", Left(rsreporte.Fields!Colegiatura, 5))
            lnEsObservacionEmergencia = 0
            If Not IsNull(rsreporte.Fields!EsObservacionEmergencia) Then
               If rsreporte.Fields!EsObservacionEmergencia = True Then
                    lnEsObservacionEmergencia = 1
                    Set oRsTmp1 = mo_ReglasAdmision.AtencionesEstanciaHospitalariaXidAtencionIdServicioIngreso(rsreporte.Fields!IdServicioIngreso, rsreporte.Fields!idAtencion, oConexion)
                    If oRsTmp1.RecordCount > 0 Then
                       If oRsTmp1.Fields!EsObservacionEmergencia = True Then
                          '***El Servicio de Ingreso y el Servicio de Alta son OBSERVACIONES DE EMERGENCIA
                          '***Solo se muestra datos del Servicio de Alta ya no es necesario mostrar los datos
                          '***de OBSERVACIONES
                          lnEsObservacionEmergencia = 0
                       Else
                          '***El servicio de Ingreso es un CONSULTORIO y el servicio de alta es OBSERVACION
                          '***Hubo transferencia desde CONSULTORIO hacia OBSERVACION
                          '***Muestra datos del CONSULTORIO y datos DE OBSERVACION
                          lcServicioObs = lcServicio
                          lcMedicoObs = lcMedico
                          lcDxObs = lcDx1
                          ldFechaIngObs = oRsTmp1.Fields!FechaDesocupacion
                          lcHoraIngObs = oRsTmp1.Fields!HoraDesocupacion
                          'debb 05/12/2012
                          If IsNull(oRsTmp1.Fields!cama) Then  'debb 05/12/2012
                             MsgBox "Para la cuenta. " & rsreporte.Fields!idCuentaAtencion & " no se registró cama, siendo un SERVICIO DE OBSERVACION", vbInformation, "Mensaje"
                             Exit Sub
                          Else
                             lcCamaObs = oRsTmp1.Fields!cama
                          End If
                          'debb 05/12/2012
                          ldFechaSalObs = rsreporte.Fields!fechaEgreso
                          lcHoraSalObs = rsreporte.Fields!HoraEgreso
                          lcServicio = IIf(IsNull(oRsTmp1.Fields!codigoServicioSEM), ".", oRsTmp1.Fields!codigoServicioSEM)
                          lcMedico = IIf(IsNull(oRsTmp1.Fields!Colegiatura), ".", Left(oRsTmp1.Fields!Colegiatura, 5))
                       End If
                    End If
                    oRsTmp1.Close
               End If
            End If
            oRsFox.AddNew
            oRsFox.Fields!em_codigo = lnNumMov
            oRsFox.Fields!em_codest = Trim(Str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
            oRsFox.Fields!em_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
            oRsFox.Fields!em_red = lcBuscaParametro.SeleccionaFilaParametro(240)
            oRsFox.Fields!em_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
            If Not IsNull(rsreporte.Fields!NroHistoriaClinica) Then
               oRsFox.Fields!em_registro = Trim(Str(rsreporte.Fields!NroHistoriaClinica))
            Else
               oRsFox.Fields!em_registro = "."
            End If
            oRsFox.Fields!em_fecha = rsreporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
            If IsNull(rsreporte.Fields!IdTipoGravedad) Then
               oRsFox.Fields!em_conding = 1
            Else
               oRsFox.Fields!em_conding = IIf(rsreporte.Fields!IdTipoGravedad = 5, 2, 1)
            End If
            oRsFox.Fields!em_apenom = Left(Trim(rsreporte.Fields!ApellidoPaterno) & " " & Trim(rsreporte.Fields!ApellidoMaterno) & " " & Trim(rsreporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsreporte.Fields!SegundoNombre), "", Trim(rsreporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!em_sexo = Trim(Str(rsreporte.Fields!idTipoSexo))
            oRsFox.Fields!em_edad = rsreporte.Fields!Edad
            oRsFox.Fields!em_tedad = rsreporte.Fields!tipoEdadSEM
            oRsFox.Fields!em_tipopac = rsreporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!em_direccion = Left(lcDireccionDomicilio, 50)
            If Not IsNull(rsreporte.Fields!IdDistritoProcedencia) Then
               oRsFox.Fields!em_procedencia = Right("0" & Trim(Str(rsreporte.Fields!IdDistritoProcedencia)), 6)
            Else
               oRsFox.Fields!em_procedencia = " "
            End If
            If Not IsNull(rsreporte.Fields!IdDistritoDomicilio) Then
               oRsFox.Fields!em_residencia = Right("0" & Trim(Str(rsreporte.Fields!IdDistritoDomicilio)), 6)
            Else
               oRsFox.Fields!em_residencia = " "
            End If
            oRsFox.Fields!em_acompa = Left(lcNombreAcompaniante, 40)
            oRsFox.Fields!em_horate = rsreporte.Fields!HoraIngreso
            oRsFox.Fields!em_diagp = lcDx1
            oRsFox.Fields!em_diags = lcDx2
            oRsFox.Fields!em_medico = lcMedico
            oRsFox.Fields!em_motivo = lcMotivo
            oRsFox.Fields!em_servicio = lcServicio
            oRsFox.Fields!em_observ = lnEsObservacionEmergencia
            If lnEsObservacionEmergencia = 1 Then
                oRsFox.Fields!em_obshin = lcHoraIngObs
                oRsFox.Fields!em_obshsa = lcHoraSalObs
                oRsFox.Fields!em_ofechin = ldFechaIngObs
                oRsFox.Fields!em_ofechsa = ldFechaSalObs
                oRsFox.Fields!em_omedico = lcMedicoObs
                oRsFox.Fields!em_ocama = lcCamaObs
                oRsFox.Fields!em_ocdx = lcDxObs
            End If
            oRsFox.Fields!em_destino = lcDestinoAtencion
            oRsFox.Fields!em_referido = lcReferido
            oRsFox.Fields!em_condeg = rsreporte.Fields!codigoCondicionAltaSEM
            oRsFox.Fields!em_fechareg = Format(Date, "mm/dd/yyyy")
            oRsFox.Fields!em_usuario = lcUsuario
            oRsFox.Update
        Else
            '*********************************Hospitalizacion
            oRsFox.AddNew
            oRsFox.Fields!numMov = lnNumMov
            oRsFox.Fields!cod_disa = lcBuscaParametro.SeleccionaFilaParametro(239)
            oRsFox.Fields!cod_red = lcBuscaParametro.SeleccionaFilaParametro(240)
            oRsFox.Fields!cod_mred = lcBuscaParametro.SeleccionaFilaParametro(241)
            oRsFox.Fields!cod_est = Trim(Str(Val(lcBuscaParametro.SeleccionaFilaParametro(208))))
            If Not IsNull(rsreporte.Fields!NroHistoriaClinica) Then
              oRsFox.Fields!numHc = Left(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), 7)
            Else
              oRsFox.Fields!numHc = ""
            End If
            oRsFox.Fields!apenom = Left(Trim(rsreporte.Fields!ApellidoPaterno) & " " & Trim(rsreporte.Fields!ApellidoMaterno) & " " & Trim(rsreporte.Fields!PrimerNombre) & " " & IIf(IsNull(rsreporte.Fields!SegundoNombre), "", Trim(rsreporte.Fields!SegundoNombre)), 50)
            oRsFox.Fields!Sexo = Trim(Str(rsreporte.Fields!idTipoSexo))
            oRsFox.Fields!Edad = rsreporte.Fields!Edad
            oRsFox.Fields!TipoEdad = rsreporte.Fields!tipoEdadSEM
            oRsFox.Fields!Ubigeo = lcBuscaParametro.SeleccionaFilaParametro(242)
            oRsFox.Fields!fecIng = rsreporte.Fields!FechaIngreso  'Format(rsReporte.Fields!FechaIngreso, "mm/dd/yyyy")
            oRsFox.Fields!fecegr = rsreporte.Fields!fechaEgreso    'Format(rsReporte.Fields!FechaEgreso, "mm/dd/yyyy")
            oRsFox.Fields!totalest = lnEstancia
            oRsFox.Fields!Servicio = IIf(IsNull(rsreporte.Fields!codigoServicioSEM), "xx", rsreporte.Fields!codigoServicioSEM)
            oRsFox.Fields!condicion = rsreporte.Fields!codigoCondicionAltaSEM
            oRsFox.Fields!seguro = rsreporte.Fields!CodigoFuenteFinanciamientoSEM
            oRsFox.Fields!coddiag = lcDx1
            oRsFox.Fields!coddiag2 = lcDx2
            oRsFox.Fields!ubicacionh = IIf(IsNull(rsreporte.Fields!ubicacionSEM), "xxxxx", rsreporte.Fields!ubicacionSEM)
            oRsFox.Fields!medicoh = IIf(IsNull(rsreporte.Fields!Colegiatura), "xxxxx", Left(Trim(rsreporte.Fields!Colegiatura), 5))
            oRsFox.Fields!estado = 1
            oRsFox.Fields!FechaReg = Format(Date, "mm/dd/yyyy")
            oRsFox.Update
        End If
        lnNumMov = lnNumMov + 1
        mo_ProgressRpt1.Value = mo_ProgressRpt1.Value + 1: DoEvents
        rsreporte.MoveNext
      Loop
      Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 140, mo_lcNombrePc, "Exportar al SEM: " & lcFechaInicio & " al " & lcFechaFinal)       '500+ ListBarReporte.idReporte
    End If
    oRsFox.Close
    oConexionFox.Close
    oConexion.Close
    Set oConexion = Nothing
    Set oConexionFox = Nothing
    Set oRsFox = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description & Chr(13) & lcParaMsgbox    'debb 05/12/2012
    Resume
End Sub




'******************************************** SIP2000 ********************************************
'solo se exporta la especliadad=OBSTETRICIA, ya sea en Consultorios Externos, Hospitalización, Emergencia
'debb-07/12/2012
Sub ExportaSIP2000(lcFechaInicio As String, lcFechaFinal As String, lcHoraInicio As String, lcHoraFinal As String)
        On Error GoTo ErrorProceso
        
        Dim oConexion As New Connection
        Dim oConexionFox As New Connection
        Dim oConexionExterna As New Connection
        
        Dim lcSql As String
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oRsFox1 As New Recordset
        Dim lcCodigoRenaes As String
        Dim lnContador As Long, lcRenaesDestino As String, lcRenaesOrigen As String, lnPeso As Double
        Dim lnTalla As Long, lcDxHosp As String, lcServ_vdrl As String, ldSer_rpr_f As Date
        Dim lnHem_hbgrs As Long, ldDserv_vdrl_f As Date, lcSer_rpr As String, ldHem_fecha As Date
        Dim lnRn_hc As Long, lcRn_nom As String, ldRn_fecha_n As Date, lnRn_peso As Double
        Dim lnRn_talla As Long, lnRn_sexo As Long, lnApgar_1 As Long, lnApgar_5 As Long
        Dim lnIdVivo As Long
        Dim ldFUM As Date, ldFUMfecha As Date
        Dim lcDxEgreso As String, lcDxFallecido As String, lnIdEstablecimientoTraslado As Long
        Dim ldFechaHosp As Date, lnEstadoCivil As Long, lnEstudios As Long, lcDistrito As String
        Dim oDODiagnostico As New DODiagnostico
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
        Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
        Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
        
        '
        lcCodigoRenaes = lcBuscaParametro.SeleccionaFilaParametro(280)
        '
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.CursorLocation = adUseClient
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=his"
        '
        mo_ReglasComunes.PreparaTablasDBFsip2000 oRsFox1, oConexionFox
        '
        Set oRsTmp1 = mo_ReglasAdmision.AtencionesParaSip2000(lcFechaInicio, lcFechaFinal, oConexion)
        
        lnContador = oRsTmp1.RecordCount
        If lnContador > 0 Then
           mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnContador: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
           lnContador = 1
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              mo_ProgressRpt1.Value = lnContador: lnContador = lnContador + 1
              DoEvents
              '
              If CDate(oRsTmp1.Fields!FechaIngreso & " " & oRsTmp1.Fields!HoraIngreso) >= CDate(lcFechaInicio & " " & lcHoraInicio) And CDate(oRsTmp1.Fields!FechaIngreso & " " & oRsTmp1.Fields!HoraIngreso) <= CDate(lcFechaFinal & " " & lcHoraFinal) Then
                    lcRenaesOrigen = ""
                    If Not IsNull(oRsTmp1.Fields!IdEstablecimientoOrigen) And oRsTmp1.Fields!IdEstablecimientoOrigen > 0 Then
                       Set oRsTmp2 = mo_ReglasComunes.HIS_ESTABseleccionarXcodigo(oRsTmp1.Fields!IdEstablecimientoOrigen, oConexion)
                       If oRsTmp2.RecordCount > 0 Then
                          lcRenaesOrigen = oRsTmp2.Fields!cod_estab
                       End If
                    End If
                    '
                    lcRenaesDestino = ""
                    If Not IsNull(oRsTmp1.Fields!idEstablecimientoDestino) And oRsTmp1.Fields!idEstablecimientoDestino > 0 Then
                       Set oRsTmp2 = mo_ReglasComunes.HIS_ESTABseleccionarXcodigo(oRsTmp1.Fields!idEstablecimientoDestino, oConexion)
                       oRsTmp2.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                       If oRsTmp2.RecordCount > 0 Then
                          lcRenaesDestino = oRsTmp2.Fields!cod_estab
                       End If
                    End If
                    '
                    lnPeso = 0: lnTalla = 0
                    Set oRsTmp2 = mo_ReglasAdmision.AtencionCESeleccionarPorIdAtencion(oRsTmp1.Fields!idAtencion, oConexionExterna)
                    If oRsTmp2.RecordCount <> 0 Then
                       If Not oRsTmp2.EOF Then
                            lnPeso = IIf(IsNull(oRsTmp2.Fields!triajePeso), 0, oRsTmp2.Fields!triajePeso)
                            lnTalla = IIf(IsNull(oRsTmp2.Fields!triajeTalla), 0, oRsTmp2.Fields!triajeTalla)
                       End If
                    End If
                    '
                    lcDxHosp = "": ldFechaHosp = 0
                    If oRsTmp1.Fields!idTipoServicio = 3 Then
                        Set oRsTmp2 = mo_ReglasAdmision.AtencionesDiagnosticosParaSip2000(oRsTmp1.Fields!idAtencion, oConexion)
                        If oRsTmp2.RecordCount > 0 Then
                           lcDxHosp = oRsTmp2.Fields!CodigoCIE2004
                        End If
                        ldFechaHosp = oRsTmp1.Fields!FechaIngreso
                    End If
                    '
                    lcDxEgreso = "": lcDxFallecido = ""
                    Set oDODiagnostico = mo_ReglasFacturacion.DevuelveDxAltaMedica(oRsTmp1.Fields!idAtencion, oRsTmp1.Fields!idTipoServicio)
                    If oDODiagnostico.descripcion <> "" Then
                       If oRsTmp1.Fields!IdCondicionAlta = 4 Then   'Fallecido
                          lcDxFallecido = oDODiagnostico.CodigoCIE2004
                       Else
                          lcDxEgreso = oDODiagnostico.CodigoCIE2004
                       End If
                    End If
                    '
                    lcServ_vdrl = ""
                    ldDserv_vdrl_f = 0
                    lcSer_rpr = ""
                    ldSer_rpr_f = 0
                    lnHem_hbgrs = 0
                    ldHem_fecha = 0
                    Set oRsTmp2 = mo_ReglasLaboratorio.LabResultadoParaSip200(oRsTmp1.Fields!idCuentaAtencion, oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                       oRsTmp2.MoveFirst
                       Do While Not oRsTmp2.EOF
                          Select Case oRsTmp2.Fields!idAnalisis
                          Case "86592"   'Prueba cualitativa de sífilis
                               If InStr(oRsTmp2.Fields!resultadoAnalisis, "\") > 1 Then
                                    lcServ_vdrl = Left(oRsTmp2.Fields!resultadoAnalisis, InStr(oRsTmp2.Fields!resultadoAnalisis, "\") - 1)
                                    If Not IsNull(oRsTmp2.Fields!fecha) Then
                                       ldDserv_vdrl_f = oRsTmp2.Fields!fecha
                                    End If
                               End If
                          Case "86593"   'RPR Sifilis (Test Cuantitativo para Sífilis)
                               If InStr(oRsTmp2.Fields!resultadoAnalisis, "\") > 1 Then
                                    lcSer_rpr = Left(oRsTmp2.Fields!resultadoAnalisis, InStr(oRsTmp2.Fields!resultadoAnalisis, "\") - 1)
                                    If Not IsNull(oRsTmp2.Fields!fecha) Then
                                       ldSer_rpr_f = oRsTmp2.Fields!fecha
                                    End If
                               End If
                          Case "85018"   'Hemoglobina
                               If InStr(oRsTmp2.Fields!resultadoAnalisis, "\") > 1 Then
                                    lnHem_hbgrs = Left(oRsTmp2.Fields!resultadoAnalisis, InStr(oRsTmp2.Fields!resultadoAnalisis, "\") - 1)
                                    If Not IsNull(oRsTmp2.Fields!fecha) Then
                                       ldHem_fecha = oRsTmp2.Fields!fecha
                                    End If
                               End If
                          End Select
                          oRsTmp2.MoveNext
                       Loop
                    End If
                    '
                    ldFUM = 0
                    ldFUMfecha = 0
                    Set oRsTmp2 = mo_ReglasImagenes.ImagMovimientoParaSip2000(oRsTmp1.Fields!idCuentaAtencion, oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                       ldFUM = oRsTmp2.Fields!Eo_FUM
                       ldFUMfecha = oRsTmp2.Fields!fecha
                    End If
                    '
                    ldRn_fecha_n = 0
                    lnRn_peso = 0
                    lnRn_talla = 0
                    lnRn_sexo = 0
                    lnApgar_1 = 99
                    lnApgar_5 = 99
                    lnIdVivo = 0
                    Set oRsTmp2 = mo_ReglasAdmision.AtencionesNacimientosXidAtencion(oRsTmp1.Fields!idAtencion, oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                       oRsTmp2.MoveFirst
                       Do While Not oRsTmp2.EOF
                          If Not IsNull(oRsTmp2.Fields!FechaNacimiento) Then
                             ldRn_fecha_n = oRsTmp2.Fields!FechaNacimiento
                          End If
                          If Not IsNull(oRsTmp2.Fields!Peso) Then
                             lnRn_peso = oRsTmp2.Fields!Peso
                          End If
                          If Not IsNull(oRsTmp2.Fields!Talla) Then
                             lnRn_talla = oRsTmp2.Fields!Talla
                          End If
                          If Not IsNull(oRsTmp2.Fields!idTipoSexo) Then
                             lnRn_sexo = oRsTmp2.Fields!idTipoSexo
                          End If
                          If Not IsNull(oRsTmp2.Fields!apgar_1) Then
                             lnApgar_1 = oRsTmp2.Fields!apgar_1
                          End If
                          If Not IsNull(oRsTmp2.Fields!apgar_5) Then
                             lnApgar_5 = oRsTmp2.Fields!apgar_5
                          End If
                          If Not IsNull(oRsTmp2.Fields!idCondicionRN) Then
                             lnIdVivo = oRsTmp2.Fields!idCondicionRN
                          End If
                          oRsTmp2.MoveNext
                       Loop
                    End If
                    '
                    lnEstadoCivil = 0
                    If Not IsNull(oRsTmp1.Fields!IdEstadoCivil) Then
                        Set oRsTmp2 = mo_ReglasComunes.TiposEstadoCivilXidEstadoCivil(oRsTmp1.Fields!IdEstadoCivil, oConexion)
                        If oRsTmp2.RecordCount > 0 Then
                           lnEstadoCivil = oRsTmp2.Fields!sip2000
                        End If
                    End If
                    '
                    lnEstudios = 0
                    If Not IsNull(oRsTmp1.Fields!IdGradoInstruccion) Then
                        Set oRsTmp2 = mo_ReglasComunes.TiposGradoInstruccionXidGradoInstruccion(oRsTmp1.Fields!IdGradoInstruccion, oConexion)
                        If oRsTmp2.RecordCount > 0 Then
                           lnEstudios = oRsTmp2.Fields!sip2000
                        End If
                    End If
                    '
                    lnRn_hc = 0
                    lcRn_nom = ""
                    '
                    oRsFox1.AddNew
                    oRsFox1.Fields!idAtencion = oRsTmp1.Fields!idAtencion
                    oRsFox1.Fields!cod_renaes = lcCodigoRenaes
                    oRsFox1.Fields!fecha = oRsTmp1.Fields!FechaIngreso
                    oRsFox1.Fields!Hora = oRsTmp1.Fields!HoraIngreso
                    oRsFox1.Fields!tiposerv = IIf(oRsTmp1.Fields!idTipoServicio = 1, "C", IIf(oRsTmp1.Fields!idTipoServicio = 3, "H", "E"))
                    oRsFox1.Fields!nro_hc = IIf(IsNull(oRsTmp1.Fields!NroHistoriaClinica), 0, oRsTmp1.Fields!NroHistoriaClinica)
                    oRsFox1.Fields!nom_pac = oRsTmp1.Fields!nPaciente & " " & IIf(IsNull(oRsTmp1.Fields!SegundoNombre), "", oRsTmp1.Fields!SegundoNombre)
                    If oRsTmp1.Fields!IdDocIdentidad = 1 And Not (IsNull(oRsTmp1.Fields!nrodocumento)) Then
                       oRsFox1.Fields!DNI = oRsTmp1.Fields!nrodocumento
                    End If
                    oRsFox1.Fields!reanes_d = lcRenaesDestino
                    oRsFox1.Fields!renaes_o = lcRenaesOrigen
                    If Not IsNull(oRsTmp1.Fields!DireccionDomicilio) Then
                       oRsFox1.Fields!Direccion = oRsTmp1.Fields!DireccionDomicilio
                    End If
                    If Not IsNull(oRsTmp1.Fields!IdCentroPobladoDomicilio) Then
                       oRsFox1.Fields!localidad = oRsTmp1.Fields!IdCentroPobladoDomicilio
                    End If
                    oRsFox1.Fields!Edad = oRsTmp1.Fields!Edad
                    oRsFox1.Fields!estado_civ = lnEstadoCivil
                    If Not IsNull(oRsTmp1.Fields!TELEFONO) Then
                       oRsFox1.Fields!TELEFONO = oRsTmp1.Fields!TELEFONO
                    End If
                    If Not IsNull(oRsTmp1.Fields!IdDistritoDomicilio) Then
                       lcDistrito = Right("000" & Trim(Str(oRsTmp1.Fields!IdDistritoDomicilio)), 6)
                       oRsFox1.Fields!dpto = Val(Left(lcDistrito, 2))
                       oRsFox1.Fields!provincia = Val(Left(lcDistrito, 4))
                       oRsFox1.Fields!Distrito = Val(lcDistrito)
                    End If
                    oRsFox1.Fields!estudios = lnEstudios
                    oRsFox1.Fields!Peso = lnPeso
                    oRsFox1.Fields!Talla = lnTalla
                    If Not IsNull(oRsTmp1.Fields!GrupoSanguineo) Then
                       oRsFox1.Fields!grupo_sang = oRsTmp1.Fields!GrupoSanguineo
                    End If
                    If Not IsNull(oRsTmp1.Fields!FactorRh) Then
                       oRsFox1.Fields!rh = oRsTmp1.Fields!FactorRh
                    End If
                    If oRsTmp1.Fields!idTipoServicio = 3 Then
                        oRsFox1.Fields!hosp_diag = lcDxHosp
                    End If
                    oRsFox1.Fields!ser_vdrl = lcServ_vdrl
                    oRsFox1.Fields!ser_vdrl_f = ldDserv_vdrl_f
                    oRsFox1.Fields!ser_rpr = lcSer_rpr
                    oRsFox1.Fields!ser_rpr_f = ldSer_rpr_f
                    oRsFox1.Fields!hem_hbgrs = lnHem_hbgrs
                    oRsFox1.Fields!hem_fecha = ldHem_fecha
                    oRsFox1.Fields!rn_fecha_n = ldRn_fecha_n
                    oRsFox1.Fields!rn_peso = lnRn_peso
                    oRsFox1.Fields!rn_talla = lnRn_talla
                    oRsFox1.Fields!rn_sexo = lnRn_sexo
                    oRsFox1.Fields!rn_vivo = lnIdVivo
                    oRsFox1.Fields!apgar_1 = lnApgar_1
                    oRsFox1.Fields!apgar_5 = lnApgar_5
                    If ldFUM <> 0 Then
                        oRsFox1.Fields!FUM = ldFUM
                        oRsFox1.Fields!FUMfecha = ldFUMfecha
                    End If
                    oRsFox1.Fields!egre_dx = lcDxEgreso
                    oRsFox1.Fields!egre_dxfal = lcDxFallecido
                    oRsFox1.Update
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
        Exit Sub
ErrorProceso:
    MsgBox Err.Description
    Resume
End Sub





Private Function isValidField(obj_Field As ADODB.Field) As Boolean
    With obj_Field

        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function

'debb-29/04/2016
Sub GeneraTXT_una_linea(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, _
                        lcSeparadorCampo As String, lcSeparadorSaltoLinea As String)
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long
        
        
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        Select Case lnBarraNumero
        Case 1
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnBarraTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
        Case 2
            mo_ProgressRpt2.Min = 0: mo_ProgressRpt2.Max = lnBarraTotal: mo_ProgressRpt2.ShowText = True: mo_ProgressRpt2.Color = vbGreen
        Case 3
            mo_ProgressRpt3.Min = 0: mo_ProgressRpt3.Max = lnBarraTotal: mo_ProgressRpt3.ShowText = True: mo_ProgressRpt3.Color = vbGreen
        Case 4
            mo_ProgressRpt4.Min = 0: mo_ProgressRpt4.Max = lnBarraTotal: mo_ProgressRpt4.ShowText = True: mo_ProgressRpt4.Color = vbGreen
        Case 5
            mo_ProgressRpt5.Min = 0: mo_ProgressRpt5.Max = lnBarraTotal: mo_ProgressRpt5.ShowText = True: mo_ProgressRpt5.Color = vbGreen
        Case 6
            mo_ProgressRpt6.Min = 0: mo_ProgressRpt6.Max = lnBarraTotal: mo_ProgressRpt6.ShowText = True: mo_ProgressRpt6.Color = vbGreen
        End Select
        '
        oRsTmp1.MoveFirst
        lcLinea = ""
        Do While Not oRsTmp1.EOF
           Select Case lnBarraNumero
           Case 1
                mo_ProgressRpt1.Value = lnContador
           Case 2
                mo_ProgressRpt2.Value = lnContador
           Case 3
                mo_ProgressRpt3.Value = lnContador
           Case 4
                mo_ProgressRpt4.Value = lnContador
           Case 5
                mo_ProgressRpt5.Value = lnContador
           Case 6
                mo_ProgressRpt6.Value = lnContador
           End Select
           DoEvents
           lnContador = lnContador + 1
           '
           
           For lnCont = 0 To lnUltimoCampo - 1
               lcCampoValor = ""
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                    Case adDouble, adInteger
                         lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If lnUltimoCampo - 1 = lnCont Then
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorSaltoLinea
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
           Next
           
           oRsTmp1.MoveNext
        Loop
        lcLinea = Left(lcLinea, Len(lcLinea) - 1)
        Print #lniFreeFile, lcLinea
        Close #lniFreeFile
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
       ' Resume
End Sub


'debb-22/04/2016
Sub GeneraTXT(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, _
                                                lcSeparadorCampo As String, lcSeparadorSaltoLinea As String)
        
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long, lcDNI As String
        'Const lcSeparadorCampo As String = "|":  Dim lcSeparadorSaltoLinea As String
        Dim lcDisa As String, lcFormato As String, lcNumero As String
        Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
        Dim oConexionExterna As New Connection
        Dim oRsAfiliadosSIS As New Recordset
        If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
            oConexionExterna.CommandTimeout = 300
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            oConexionExterna.CursorLocation = adUseClient
        End If
        
        'lcSeparadorSaltoLinea = Chr(10)
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        Select Case lnBarraNumero
        Case 1
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnBarraTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
        Case 2
            mo_ProgressRpt2.Min = 0: mo_ProgressRpt2.Max = lnBarraTotal: mo_ProgressRpt2.ShowText = True: mo_ProgressRpt2.Color = vbGreen
        Case 3
            mo_ProgressRpt3.Min = 0: mo_ProgressRpt3.Max = lnBarraTotal: mo_ProgressRpt3.ShowText = True: mo_ProgressRpt3.Color = vbGreen
        Case 4
            mo_ProgressRpt4.Min = 0: mo_ProgressRpt4.Max = lnBarraTotal: mo_ProgressRpt4.ShowText = True: mo_ProgressRpt4.Color = vbGreen
        Case 5
            mo_ProgressRpt5.Min = 0: mo_ProgressRpt5.Max = lnBarraTotal: mo_ProgressRpt5.ShowText = True: mo_ProgressRpt5.Color = vbGreen
        Case 6
            mo_ProgressRpt6.Min = 0: mo_ProgressRpt6.Max = lnBarraTotal: mo_ProgressRpt6.ShowText = True: mo_ProgressRpt6.Color = vbGreen
        End Select
        '
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           Select Case lnBarraNumero
           Case 1
                mo_ProgressRpt1.Value = lnContador
           Case 2
                mo_ProgressRpt2.Value = lnContador
           Case 3
                mo_ProgressRpt3.Value = lnContador
           Case 4
                mo_ProgressRpt4.Value = lnContador
           Case 5
                mo_ProgressRpt5.Value = lnContador
           Case 6
                mo_ProgressRpt6.Value = lnContador
           End Select
           DoEvents: lnContador = lnContador + 1
           '
           lcLinea = ""
           lcDNI = ""
           For lnCont = 0 To lnUltimoCampo - 1
               If UCase(oRsTmp1.Fields(lnCont).Name) = "CABNROENVIOALSIS" Then
                  If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                     oRsTmp1.Fields(lnCont).Value = 1
                  Else
                     oRsTmp1.Fields(lnCont).Value = Trim(Str(Val(oRsTmp1.Fields(lnCont).Value) + 1))
                  End If
                  oRsTmp1.Update
               End If
               lcCampoValor = ""
               
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                         If UCase(oRsTmp1.Fields(lnCont).Name) = "DXCODIGO" Or UCase(oRsTmp1.Fields(lnCont).Name) = "CODIGO" And oRsTmp1.Fields(lnCont).DefinedSize = 5 Then
                             lcCampoValor = sighEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp1.Fields(lnCont).Value, 5))
                         End If
                         If lcCampoValor = lcSeparadorCampo Then
                            lcCampoValor = "_"
                         End If
                    Case adDouble, adInteger
                         lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If (lnUltimoCampo - 1) = lnCont Then
                  If lcSeparadorSaltoLinea = "" Then
                     lcLinea = lcLinea & lcCampoValor
                  Else
                     lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
                  End If
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
               If Right(lcArchivoTXT, 13) = "Pacientes.txt" And UCase(oRsTmp1.Fields(lnCont).Name) = "NRODOCUMENTO" Then
                  lcDNI = lcCampoValor
               End If
           Next
           '
           If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
                lcDisa = "": lcFormato = "": lcNumero = ""
                If Len(lcDNI) = 8 Then
                     Set oRsAfiliadosSIS = mo_ReglasSISgalenhos.SisFiliacionesXdni(lcDNI, oConexionExterna)
                     If oRsAfiliadosSIS.RecordCount > 0 Then
                        lcDisa = Trim(oRsAfiliadosSIS.Fields!AfiliacionDisa)
                        lcFormato = Trim(oRsAfiliadosSIS.Fields!AfiliacionTipoFormato)
                        lcNumero = Trim(oRsAfiliadosSIS.Fields!AfiliacionNroFormato)
                     End If
                     oRsAfiliadosSIS.Close
                End If
                lcLinea = lcLinea & lcDisa & lcSeparadorCampo & lcFormato & lcSeparadorCampo & lcNumero & lcSeparadorCampo
           End If
           Print #lniFreeFile, lcLinea
           oRsTmp1.MoveNext
        Loop
        Close #lniFreeFile
        Set oConexionExterna = Nothing
        Set oRsAfiliadosSIS = Nothing
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
        'Resume
End Sub


Sub CrearReporteObservaciones_excel(mrs_Tmp As Recordset, lnHwnd As Long)
'Dim oExcel As Excel.Application
'Dim oWorkBookPlantilla As Workbook
'Dim oWorkBook As Workbook
'Dim oWorkSheet As Worksheet
'Dim mo_ReporteUtil As New sighentidades.ReporteUtil
'Dim iFila As Long
'Dim lnTotal As Double
'        On Error GoTo ManejadorErrorExcel
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        'Crea nueva hoja
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
'        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'        iFila = 2
'        oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
'        iFila = 4
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
'        oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
'        oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
'        oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
'        oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
'        oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
'        oWorkSheet.Cells(iFila, 8).Value = "Servicio"
'        oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
'        oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
'        oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        iFila = 6: lnTotal = 0
'        mrs_Tmp.MoveFirst
'        Do While Not mrs_Tmp.EOF
'           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
'           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
'           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
'           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
'           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
'           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
'           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
'           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
'           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("FechaEgreso").Value
'           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("observaciones").Value
'           iFila = iFila + 1
'           mrs_Tmp.MoveNext
'        Loop
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
'        oExcel.Visible = True
'        oWorkSheet.PrintPreview
'        'oWorkSheet.PrintOut
'    Set oWorkSheet = Nothing
'    Set oExcel = Nothing
'        Exit Sub
'ManejadorErrorExcel:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'        Resume
'    End Select
'    Exit Sub
Dim mo_ReporteUtil As New sighEntidades.ReporteUtil
Dim iFila As Long
Dim lnTotal As Double
Dim lbEsOpenOffice As Boolean

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
        
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
        iFila = 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Lista de Cuentas Observadas (No fueron grabados sus FUAS)")
            Else
                oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
            End If
        iFila = 4
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50

            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Cuenta")
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Apellido Paterno")
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Apellido Materno")
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula("Primer Nombre")
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula("N° Historia Clinica")
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula("Tipo Servicio")
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula("Servicio")
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Fecha Ingreso")
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula("Fecha Alta")
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("Observaciones")
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
            oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
            oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
            oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
            oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
            oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
            oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
            oWorkSheet.Cells(iFila, 8).Value = "Servicio"
            oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
            oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
            oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
        End If
        
        iFila = 6: lnTotal = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("idCuentaAtencion").Value)
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("ApellidoPaterno").Value)
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("apellidoMaterno").Value)
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("PrimerNombre").Value)
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp.Fields("tipoServicio").Value)
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("Servicio").Value)
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp.Fields("FechaIngreso").Value)
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mrs_Tmp.Fields("FechaEgreso").Value)
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula(mrs_Tmp.Fields("observaciones").Value)
        Else
            oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
            oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
            oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
            oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
            oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
            oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
            oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
            oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
            oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("FechaEgreso").Value
            oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("observaciones").Value
        End If
           iFila = iFila + 1
           mrs_Tmp.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
            oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
        End If
    If lbEsOpenOffice = True Then
            Call Document.getCurrentController.GetFrame.getContainerWindow.SetVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If

        Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select
    Exit Sub

End Sub

Function ChequeaCuentasSinFUAgrabado(ByVal mrs_Tmp As Recordset, oConexionExterna As Connection, lcFechaInicio As String, lcFechaFinal As String) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lnTotalReg As Long, lnIdCuentaInicial As Long, lbSeImprime As Boolean
        Dim mo_ReglasAdmision As New ReglasAdmision
        Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
        On Error GoTo ErrorProceso1
        Set oRsTmp1 = mo_ReglasAdmision.AtencionesSisPorFechas(lcFechaInicio, lcFechaFinal)
        lnTotalReg = oRsTmp1.RecordCount
        'Chequea si hay Cuentas Observadas (No se grabaron en tablas SIS FUA)
        If lnTotalReg > 0 Then
           mo_ProgressRpt7.Min = 0: mo_ProgressRpt7.Max = lnTotalReg: mo_ProgressRpt7.ShowText = True: mo_ProgressRpt7.Color = vbGreen
           lnIdCuentaInicial = 1
           With mrs_Tmp
                  .Fields.Append "idCuentaAtencion", adInteger
                  .Fields.Append "ApellidoPaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "apellidoMaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "PrimerNombre", adVarChar, 50, adFldIsNullable
                  .Fields.Append "nroHistoriaClinica", adInteger
                  .Fields.Append "tipoServicio", adVarChar, 50, adFldIsNullable
                  .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                  .Fields.Append "FechaIngreso", adDate
                  .Fields.Append "FechaEgreso", adDate, 8, adFldIsNullable
                  .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
                  .LockType = adLockOptimistic
                  .Open
           End With
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              mo_ProgressRpt7.Value = lnIdCuentaInicial: lnIdCuentaInicial = lnIdCuentaInicial + 1
              Set oRsTmp2 = mo_ReglasSISgalenhos.SisFuaAtencionSeleccionarPorCuenta(oRsTmp1.Fields!idCuentaAtencion, oConexionExterna)
              If oRsTmp2.RecordCount = 0 Then
                 lbSeImprime = False
                 If oRsTmp1.Fields!idTipoServicio = 1 Then
                    lbSeImprime = True
                    lcSql = "Cuenta de Consultorios externos, no se ha impreso FICHA FUA"
                 Else
                    If Not IsNull(oRsTmp1.Fields!fechaEgreso) Then
                       lbSeImprime = True
                       lcSql = "Cuenta de Hospitalización/emergencia, tiene ALTA MEDICA pero no se ha impreso FICHA FUA"
                    End If
                 End If
                 If lbSeImprime = True Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
                    mrs_Tmp.Fields!ApellidoPaterno = oRsTmp1.Fields!ApellidoPaterno
                    mrs_Tmp.Fields!ApellidoMaterno = oRsTmp1.Fields!ApellidoMaterno
                    mrs_Tmp.Fields!PrimerNombre = oRsTmp1.Fields!PrimerNombre
                    mrs_Tmp.Fields!NroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!TipoServicio = oRsTmp1.Fields!TipoServicio
                    mrs_Tmp.Fields!Servicio = oRsTmp1.Fields!Servicio
                    mrs_Tmp.Fields!FechaIngreso = oRsTmp1.Fields!FechaIngreso
                    mrs_Tmp.Fields!fechaEgreso = oRsTmp1.Fields!fechaEgreso
                    mrs_Tmp.Fields!Observaciones = lcSql
                    mrs_Tmp.Update
                 End If
              End If
              oRsTmp2.Close
              oRsTmp1.MoveNext
           Loop
        End If
        Set ChequeaCuentasSinFUAgrabado = oRsTmp1
        Exit Function
ErrorProceso1:
        MsgBox Err.Description
        'Resume
End Function


Sub CreandoRutaEnDisco(lcRutaTemporal As String)
    Dim lcSql As String
    On Error Resume Next
    lcSql = "c:"
    sighEntidades.ejecutarComando lcSql
    lcSql = "cd \"
    sighEntidades.ejecutarComando lcSql
    lcSql = "md " & lcRutaTemporal
    sighEntidades.ejecutarComando lcSql
    lcSql = "cd " & lcRutaTemporal
    sighEntidades.ejecutarComando lcSql
    lcSql = "del *.*"

End Sub
Function LeerArchivoTexto(nombreFichero As String) As String
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' Obtengo el siguiente número libre de archivo
    numlib = FreeFile()
    Open nombreFichero For Input As #numlib
    ' Se ha abierto el fichero sin problemas
    isOpen = True
    ' Leo todo el contenido en una única operación
    LeerArchivoTexto = Input(LOF(numlib), numlib)
    ' Cierro el archivo
Manejador_Error:
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

'debb-22/04/2016
Function ExportaCitasWeb(ldFechaInicial As Date, ldFechaFinal As Date, oRsCuposWebSeleccionado As Recordset, _
         lnIdUsuario As Long, lnIdPaciente As Long, lbEnviarServicio As Boolean, lbEnviarPacientes As Boolean, _
         lbEnviarFteFinanciamiento As Boolean, lbEnviarMedicos As Boolean, lbEnviarTurnos As Boolean, lnIdServicio As Long) As Boolean
       On Error GoTo ErrCitasWebExp
       ExportaCitasWeb = False
       Dim oConexion As New Connection, oConexionExterna As New Connection
       Dim oRsCitasWebCupos As New Recordset, oRsCitaBloqueada As New Recordset
       Dim oRsPacientes As New Recordset, oRsFuenteFinanciamiento As New Recordset
       Dim oRsCitaBloqueada1 As New Recordset, oRsTurnos As New Recordset
       Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
       Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
       Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
       Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
       Dim oCitasBloqueadas As New CitasBloqueadas, oDoCitaBloqueada As New DOCitaBloqueada
       Dim oCitasWebCupos As New CitasWebCupos, oDOCitasWebCupos As New DOCitasWebCupos
       Dim lcRutaExportar As String, lnTotalRegistros As Long, lnIdCitaBloqueada As Long, lnIdTurno As Long
       Dim lcTelefono As String, lcEmail As String, lcUbigeo As Long
       Dim lcFuenteFinanciamiento As Long
       Dim cont As Integer
       cont = 0
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighEntidades.CadenaConexion
       oConexion.BeginTrans
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       oConexionExterna.BeginTrans
       'Elimina CUPOS RESERVADOS del mes                   *****:usar el PA pero filtrar por FEcha,IdSERVICIO
       If lnIdServicio = 0 Then
          Set oRsCitasWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal, , , oConexionExterna, 2)
       Else
          Set oRsCitasWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal, , lnIdServicio, oConexionExterna, 2)
       End If
       Set oRsCitaBloqueada = mo_ReglasDeProgMedica.CitasBloqueadasTodas(oConexion)   '*****:nuevo PA para filtrar por FEcha,IdSERVICIO
       If lnIdServicio > 0 Then
          oRsCitaBloqueada.Filter = "fecha='" & ldFechaInicial & "'"
       End If
       If oRsCitasWebCupos.RecordCount > 0 Then
          Set oCitasBloqueadas.Conexion = oConexion
          oRsCitasWebCupos.MoveFirst
          Do While Not oRsCitasWebCupos.EOF
             oDoCitaBloqueada.idCitaBloqueada = IIf(IsNull(oRsCitasWebCupos.Fields!idCitaBloqueada), 0, oRsCitasWebCupos.Fields!idCitaBloqueada)
             oDoCitaBloqueada.IdUsuarioAuditoria = lnIdUsuario
             oRsCitasWebCupos.MoveNext
          Loop
       End If
       oRsCitasWebCupos.Close
       
       '*****: agregar proc.alm. que eliminar por Fecha,IdServicio
       If mo_ReglasDeProgMedica.CitasWebCuposEliminarPorFechas(IIf(lnIdServicio = 0, CDate("01/01/2000"), ldFechaInicial), _
                                                            ldFechaFinal, oConexionExterna, lnIdServicio) = False Then
          lcParaMsgbox = mo_ReglasDeProgMedica.MensajeError: GoTo ErrCitasWebExp
       End If
       
       'Elimina CITAS BLOQUEADAS del mes
       If lnIdServicio = 0 Then
          oRsCuposWebSeleccionado.Filter = ""
       Else
          oRsCuposWebSeleccionado.Filter = "idServicio=" & Trim(Str(lnIdServicio)) & _
                                           " and fecha = '" & ldFechaInicial & "'"
       End If
       lnTotalRegistros = oRsCuposWebSeleccionado.RecordCount
       If lnTotalRegistros > 0 Then
            Set oRsTurnos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotalRegistros: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
            '
            Set oCitasBloqueadas.Conexion = oConexion
            Set oCitasWebCupos.Conexion = oConexionExterna
            oRsCuposWebSeleccionado.MoveFirst
            Do While Not oRsCuposWebSeleccionado.EOF

                mo_ProgressRpt1.Value = mo_ProgressRpt1.Value + 1
                DoEvents
                'Bloquea Cita
                lnIdCitaBloqueada = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idCitaBloqueada), 0, oRsCuposWebSeleccionado.Fields!idCitaBloqueada)
                If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoLlenadoEnCitaGalenHos And oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Then
                    If oRsCitaBloqueada.RecordCount > 0 Then
                        oRsCitaBloqueada.MoveFirst
                    End If
                    Do While Not oRsCitaBloqueada.EOF
                        If oRsCuposWebSeleccionado.Fields!fecha = oRsCitaBloqueada.Fields!fecha And oRsCuposWebSeleccionado.Fields!HoraInicio = oRsCitaBloqueada.Fields!HoraInicio And oRsCuposWebSeleccionado.Fields!idMedico = oRsCitaBloqueada.Fields!idMedico Then
                            lnIdCitaBloqueada = IIf(IsNull(oRsCitaBloqueada.Fields!idCitaBloqueada), 0, oRsCitaBloqueada.Fields!idCitaBloqueada)
                            GoTo Existe
                        End If
                        oRsCitaBloqueada.MoveNext
                    Loop

                    oDoCitaBloqueada.idUsuario = lnIdUsuario
                    oDoCitaBloqueada.fecha = oRsCuposWebSeleccionado.Fields!fecha
                    oDoCitaBloqueada.HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                    oDoCitaBloqueada.HoraFin = oRsCuposWebSeleccionado.Fields!HoraFinal
                    oDoCitaBloqueada.FechaBloqueo = Date
                    oDoCitaBloqueada.HoraBloqueo = "Web"
                    oDoCitaBloqueada.idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                    oDoCitaBloqueada.IdUsuarioAuditoria = lnIdUsuario
                    If oCitasBloqueadas.Insertar(oDoCitaBloqueada) = False Then
                    End If
                    lnIdCitaBloqueada = oDoCitaBloqueada.idCitaBloqueada
                    GoTo Siguiente
Existe:
                    oDoCitaBloqueada.idUsuario = lnIdUsuario
                    oDoCitaBloqueada.fecha = oRsCitaBloqueada.Fields!fecha
                    oDoCitaBloqueada.HoraInicio = oRsCitaBloqueada.Fields!HoraInicio
                    oDoCitaBloqueada.HoraFin = oRsCitaBloqueada.Fields!HoraFin
                    oDoCitaBloqueada.FechaBloqueo = Date
                    oDoCitaBloqueada.HoraBloqueo = "Web"
                    oDoCitaBloqueada.idMedico = oRsCitaBloqueada.Fields!idMedico
                    oDoCitaBloqueada.IdUsuarioAuditoria = lnIdUsuario
                    oDoCitaBloqueada.idCitaBloqueada = oRsCitaBloqueada.Fields!idCitaBloqueada
                    If oCitasBloqueadas.Modificar(oDoCitaBloqueada) = False Then
                    End If
                    
                Else
                    If oRsCitaBloqueada.RecordCount > 0 Then
                        oRsCitaBloqueada.MoveFirst
                    End If
                    Do While Not oRsCitaBloqueada.EOF
                        If oRsCitaBloqueada.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha And oRsCitaBloqueada.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio And oRsCitaBloqueada.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico And oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb = 1 Then
                            Dim ldFecha As Date
                            Dim lcHoraI As String
                            Dim liIdmedico As Long
    
                            ldFecha = oRsCuposWebSeleccionado.Fields!fecha
                            lcHoraI = oRsCuposWebSeleccionado.Fields!HoraInicio
                            liIdmedico = oRsCuposWebSeleccionado.Fields!idMedico
                            oDoCitaBloqueada.idCitaBloqueada = oRsCitaBloqueada.Fields!idCitaBloqueada
                            oDoCitaBloqueada.IdUsuarioAuditoria = oRsCitaBloqueada.Fields!idUsuario
    
                            If oCitasBloqueadas.Eliminar(oDoCitaBloqueada) = False Then
                            End If
                        End If
                        oRsCitaBloqueada.MoveNext
                    Loop
                End If
                '
Siguiente:
                If IsNull(oRsCuposWebSeleccionado.Fields!IdTurno) Then
                    lnIdTurno = 4
                    If oRsTurnos.RecordCount > 0 Then
                       oRsTurnos.MoveFirst
                       Do While Not oRsTurnos.EOF
                          If oRsCuposWebSeleccionado.Fields!HoraInicio >= oRsTurnos.Fields!HoraInicio And oRsCuposWebSeleccionado.Fields!HoraInicio < oRsTurnos.Fields!HoraFin Then
                             lnIdTurno = oRsTurnos.Fields!IdTurno
                             Exit Do
                          End If
                          oRsTurnos.MoveNext
                       Loop
                    End If
                Else
                    lnIdTurno = oRsCuposWebSeleccionado.Fields!IdTurno
                End If
                'Carga cupo
                
                If oRsCitaBloqueada.RecordCount > 0 Then
                    oRsCitaBloqueada.MoveFirst
                End If
                oDOCitasWebCupos.HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                oDOCitasWebCupos.HoraFinal = oRsCuposWebSeleccionado.Fields!HoraFinal
                oDOCitasWebCupos.IdServicio = oRsCuposWebSeleccionado.Fields!IdServicio
                oDOCitasWebCupos.fecha = oRsCuposWebSeleccionado.Fields!fecha
                oDOCitasWebCupos.idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                oDOCitasWebCupos.idEstadoCitaWeb = oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb
                oDOCitasWebCupos.IdTurno = lnIdTurno
                oDOCitasWebCupos.IdUsuarioAuditoria = lnIdUsuario
                oDOCitasWebCupos.idPaciente = lnIdPaciente
                If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoLlenadoEnCitaGalenHos And oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos Then
                    oDOCitasWebCupos.idCitaBloqueada = lnIdCitaBloqueada
                    oDOCitasWebCupos.DNI = ""
                    oDOCitasWebCupos.ApellidoPaterno = ""
                    oDOCitasWebCupos.ApellidoMaterno = ""
                    oDOCitasWebCupos.PrimerNombre = ""
                    oDOCitasWebCupos.SegundoNombre = ""
                    oDOCitasWebCupos.idTipoSexo = 0
                    oDOCitasWebCupos.FechaNacimiento = 0
                    oDOCitasWebCupos.Ubigeo = 0
                    oDOCitasWebCupos.IdFuenteFinanciamiento = 0
                    oDOCitasWebCupos.Email = ""
                    oDOCitasWebCupos.TELEFONO = ""
                Else
                    '
                    lcTelefono = "": lcEmail = "": lcUbigeo = 0: lcFuenteFinanciamiento = 0
                    If Not IsNull(oRsCuposWebSeleccionado.Fields!idPaciente) Then
                        Set oRsPacientes = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(oRsCuposWebSeleccionado.Fields!idPaciente, oConexion)
                        If oRsPacientes.RecordCount > 0 Then
                           If Not IsNull(oRsPacientes.Fields!TELEFONO) Then
                              lcTelefono = oRsPacientes.Fields!TELEFONO
                           End If
                           If Not IsNull(oRsPacientes.Fields!Email) Then
                              lcEmail = oRsPacientes.Fields!Email
                           End If
                           If Not IsNull(oRsPacientes.Fields!IdDistritoDomicilio) Then
                              lcUbigeo = oRsPacientes.Fields!IdDistritoDomicilio
                           End If
                        End If
                        oRsPacientes.Close
                    End If
                    '
                    If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Then
                        oDOCitasWebCupos.idCitaBloqueada = 0
                    Else
                        oDOCitasWebCupos.idCitaBloqueada = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idCitaBloqueada), 0, oRsCuposWebSeleccionado.Fields!idCitaBloqueada)
                    End If
                    oDOCitasWebCupos.DNI = IIf(IsNull(oRsCuposWebSeleccionado.Fields!DNI), "", oRsCuposWebSeleccionado.Fields!DNI)
                    oDOCitasWebCupos.ApellidoPaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoPaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoPaterno)
                    oDOCitasWebCupos.ApellidoMaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoMaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoMaterno)
                    oDOCitasWebCupos.PrimerNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!PrimerNombre), "", oRsCuposWebSeleccionado.Fields!PrimerNombre)
                    oDOCitasWebCupos.SegundoNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!SegundoNombre), "", oRsCuposWebSeleccionado.Fields!SegundoNombre)
                    oDOCitasWebCupos.idTipoSexo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idTipoSexo), 0, oRsCuposWebSeleccionado.Fields!idTipoSexo)
                    If oRsCuposWebSeleccionado.Fields!ApellidoPaterno <> "" Then
                       oDOCitasWebCupos.FechaNacimiento = IIf(IsNull(oRsCuposWebSeleccionado.Fields!FechaNacimiento), 0, oRsCuposWebSeleccionado.Fields!FechaNacimiento)
                    Else
                        oDOCitasWebCupos.FechaNacimiento = 0
                    End If
                    oDOCitasWebCupos.Ubigeo = lcUbigeo
                    Set oRsFuenteFinanciamiento = mo_ReglasDeProgMedica.CitasSeleccionarPorAtenciones(oDOCitasWebCupos.IdServicio, oDOCitasWebCupos.fecha)
                    If oRsFuenteFinanciamiento.RecordCount > 0 Then
                        oRsFuenteFinanciamiento.MoveFirst
                        Do While Not oRsFuenteFinanciamiento.EOF
                            If oRsFuenteFinanciamiento.Fields!idPaciente = oRsCuposWebSeleccionado.Fields!idPaciente Then
                                lcFuenteFinanciamiento = oRsFuenteFinanciamiento.Fields!IdFuenteFinanciamiento
                            End If
                           oRsFuenteFinanciamiento.MoveNext
                        Loop
                    End If
                    oDOCitasWebCupos.IdFuenteFinanciamiento = lcFuenteFinanciamiento
                    oDOCitasWebCupos.Email = lcEmail
                    oDOCitasWebCupos.TELEFONO = lcTelefono
                    oDOCitasWebCupos.idPaciente = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idPaciente), 0, oRsCuposWebSeleccionado.Fields!idPaciente)
                End If
                If oRsCuposWebSeleccionado!idEstadoCitaWeb <> sghCitasWebEstados.sghCWebCitaWebConfirmadoYcitado Then
                    If oCitasWebCupos.Insertar(oDOCitasWebCupos) = False Then
                    End If
                End If
                oRsCuposWebSeleccionado.MoveNext
            Loop
            'Exporta Cupos Web y Citas
            'Set oRsCitasWebCupos = mo_ReglasAdmision.CitasWebCuposSeleccionarTodas(oConexionExterna)     '*****:usar nuevo PA pero filtrar por FEcha,IdSERVICIO
            If lnIdServicio = 0 Then
               Set oRsCitasWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal, , , oConexionExterna, 2)
            Else
               Set oRsCitasWebCupos = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal, , lnIdServicio, oConexionExterna, 2)
            End If
            'GeneraTXT oRsCitasWebCupos, lcRutaExportar & "CuposWeb.txt", oRsCitasWebCupos.RecordCount, 2, "|", Chr(10)
            'Exporta a la WEB SOMEE
            Dim mo_Procesos As New SIGHProxies.Procesos
            Dim lcMensaje As String, lbSeTerminaSistema As Boolean
             If lnIdServicio = 0 Then
                mo_Procesos.MensajeError = ""
             Else
                 mo_Procesos.MensajeError = " and  day(fecha)=" & Day(ldFechaInicial) & " and month(fecha)=" & Month(ldFechaInicial) & _
                                                                " and year(fecha)=" & Year(ldFechaInicial) & " and   idServicio=" & lnIdServicio
            End If
            lcMensaje = ""
            Set mo_Procesos = Nothing
            If lcMensaje <> "" Then
               MsgBox lcMensaje, vbInformation, ""
            End If
            If lnIdServicio > 0 Then
               oRsCuposWebSeleccionado.Filter = ""
            End If
            
'            'Exporta Servicios
'            If lbEnviarServicio = True Then
'                oRsCitasWebCupos.Close
'                Set oRsCitasWebCupos = mo_ReglasArchivoClinico.ServiciosSeleccionarSoloCE
'                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Servicios.txt", oRsCitasWebCupos.RecordCount, 3, "|", Chr(10)
'            End If
'            'Exporta Medicos
'            If lbEnviarMedicos = True Then
'                oRsCitasWebCupos.Close
'                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.MedicosSeleccionarTodosParaWeb
'                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Medicos.txt", oRsCitasWebCupos.RecordCount, 4, "|", Chr(10)
'            End If
'            'Exporta Pacientes
'            If lbEnviarPacientes = True Then
'                oRsCitasWebCupos.Close
'                Set oRsCitasWebCupos = mo_ReglasAdmision.PacientesSeleccionarTodos
'                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Pacientes.txt", oRsCitasWebCupos.RecordCount, 5, "|", Chr(10)
'            End If
'            'Exporta Fuente Financiamiento
'            If lbEnviarFteFinanciamiento = True Then
'                oRsCitasWebCupos.Close
'                Set oRsCitasWebCupos = mo_reglasComunes.FuentesFinanciamientoSeleccionarTodos
'                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "ftefinan.txt", oRsCitasWebCupos.RecordCount, 6, "|", Chr(10)
'            End If
'            If lbEnviarTurnos = True Then
'                oRsCitasWebCupos.Close
'                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
'                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Turnos.txt", oRsCitasWebCupos.RecordCount, 1, "|", Chr(10)
'            End If
            MsgBox "Se exportó CITAS WEB correctamente "
       End If
       
       ExportaCitasWeb = True
ErrCitasWebExp:

       If ExportaCitasWeb = True Then
          oConexion.CommitTrans
          oConexionExterna.CommitTrans
       Else
       Resume
          If lcParaMsgbox <> "" Then
             MsgBox lcParaMsgbox
          Else
             MsgBox Err.Description
          End If
          oConexion.RollbackTrans
          oConexionExterna.RollbackTrans
       End If
       oConexion.Close
       Set oConexion = Nothing
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       Set oRsCitasWebCupos = Nothing
'       Set oRsCitaBloqueada = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       Set mo_ReglasArchivoClinico = Nothing
       Set mo_ReglasAdmision = Nothing
       Set oRsCitaBloqueada1 = Nothing
       Set oRsTurnos = Nothing
       Set oRsPacientes = Nothing
       Set oDoCitaBloqueada = Nothing
       Set oCitasBloqueadas = Nothing
       Set oCitasWebCupos = Nothing

End Function


Function ExportaCitasWebCargaTabla(ldFechaInicial As Date, ldFechaFinal As Date, ByRef oRsCitasWebCupos As Recordset) As Boolean
       ExportaCitasWebCargaTabla = False
       On Error GoTo excwct
       Dim oRsCuposWebSeleccionado As New Recordset
       Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
       Set oRsCuposWebSeleccionado = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal)
       If oRsCuposWebSeleccionado.RecordCount > 0 Then
            oRsCuposWebSeleccionado.MoveFirst
            Do While Not oRsCuposWebSeleccionado.EOF
                oRsCitasWebCupos.AddNew
                oRsCitasWebCupos.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                oRsCitasWebCupos.Fields!HoraFinal = oRsCuposWebSeleccionado.Fields!HoraFinal
                oRsCitasWebCupos.Fields!IdServicio = oRsCuposWebSeleccionado.Fields!IdServicio
                oRsCitasWebCupos.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha
                oRsCitasWebCupos.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                oRsCitasWebCupos.Fields!idEstadoCitaWeb = oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb
                oRsCitasWebCupos.Fields!idCitaBloqueada = oRsCuposWebSeleccionado.Fields!idCitaBloqueada
                oRsCitasWebCupos.Fields!DNI = IIf(IsNull(oRsCuposWebSeleccionado.Fields!DNI), "", oRsCuposWebSeleccionado.Fields!DNI)
                oRsCitasWebCupos.Fields!ApellidoPaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoPaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoPaterno)
                oRsCitasWebCupos.Fields!ApellidoMaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoMaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoMaterno)
                oRsCitasWebCupos.Fields!PrimerNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!PrimerNombre), "", oRsCuposWebSeleccionado.Fields!PrimerNombre)
                oRsCitasWebCupos.Fields!SegundoNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!SegundoNombre), "", oRsCuposWebSeleccionado.Fields!SegundoNombre)
                oRsCitasWebCupos.Fields!idTipoSexo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idTipoSexo), 0, oRsCuposWebSeleccionado.Fields!idTipoSexo)
                oRsCitasWebCupos.Fields!FechaNacimiento = IIf(IsNull(oRsCuposWebSeleccionado.Fields!FechaNacimiento), 0, oRsCuposWebSeleccionado.Fields!FechaNacimiento)
                oRsCitasWebCupos.Update
                oRsCuposWebSeleccionado.MoveNext
            Loop
       End If
       oRsCuposWebSeleccionado.Close
       Set oRsCuposWebSeleccionado = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       ExportaCitasWebCargaTabla = True
       Exit Function
excwct:
       lcParaMsgbox = Err.Description
End Function

Function ImportaCitasWeb(ByRef oRsCitasWebCupos As Recordset, lcArchivoTexto As String, _
                        lbDesdeArchivoTexto As Boolean, ml_lnHwnd As Long, lbManual As Boolean, _
                        oRsCitasWebPorImportarEnFormaManual As Recordset) As Boolean
     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
     If lbDesdeArchivoTexto = False Then
        'Importa desde la WEB SOMEE
        Dim mo_Procesos As New SIGHProxies.Procesos
        Dim lcMensaje As String, lbSeTerminaSistema As Boolean, oRsCitasWeb As New Recordset, lcUsuario As String
        lcUsuario = ""
        If lbManual = True Then
           lcUsuario = "MANUAL"
           Set oRsCitasWeb = oRsCitasWebPorImportarEnFormaManual.Clone
        End If
        lcMensaje = ""
        
        If lcMensaje = "" Then
            mo_Procesos.ImportaCitasWebMINSA ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc
            If mo_Procesos.MensajeError <> "" Then
               lcMensaje = ""
            End If
        End If
        Set oRsCitasWeb = Nothing
        Set mo_Procesos = Nothing
        Set mo_ReglasDeProgMedica = Nothing
        Exit Function
     
     
     End If
     '
     On Error GoTo ErrImpCWEB
     ImportaCitasWeb = False
     Dim mo_ReporteUtil As New sighEntidades.ReporteUtil
     Dim mo_ReglasReportes As New ReglasReportes
     Dim iFila As Long, lnIdWeb As Long
     Dim lineas() As String, lnRegistros As Long, Registro As String, Fichero As Integer, Campos() As String, Columna As Single
     Dim ldFecha As Date, lnIdServicio As Long, lnIdMedico As Long, lcHoraInicio As String
     Dim lcHoraFinal As String, lnIdEstadoCitaWeb As Long, lnIdCitaBloqueada As Long, lcDNI As String
     Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
     Dim ldFechaConfirmacion As Date, lcHoraConfirmacion As String, lnNroRegistro As Long
     Dim lcSegundoNombre As String, lnIdTipoSexo As Long, ldFechaNacimiento As Date, lnUbigeo As Long
     Dim lnDiasMaximoConfirmacion As Integer, lnDiasQfaltaPcita As Integer, lnIdFuenteFinanciamiento As Long
     Dim ldFechaHoy As Date, lcHoraHoy As String, lnHistoriaCreada As Long
     Dim lnIdTurno As Long, lcEmail As String, lcTelefono As String, lnIdAtencion As Long
     Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
     Dim oCitasWebCupos As New CitasWebCupos, oDOCitasWebCupos As New DOCitasWebCupos
     Dim oConexionExterna As New Connection, oConexion As New Connection
     Dim oRsTmp1 As New Recordset
     Dim lbEsOpenOffice As Boolean
     Dim lcSql As String
     lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
     '
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
        Dim lnHwnd As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
        Dim oRange As range
        Dim range As Excel.range
        Dim borders As Excel.borders
    End If
     
     
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        '
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        Set oExcel = GalenhosExcelApplication()
        Set oWorkBook = oExcel.Workbooks.Add
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
     End If
     iFila = 2
     If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Lista de HISTORIAS creadas en forma automática (Citas Web)")
     Else
        oWorkSheet.Cells(iFila, 2).Value = "Lista de HISTORIAS creadas en forma automática (Citas Web)"
     End If
     iFila = 4
     If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, iFila - 1).setFormula("N° Historia Clinica")
        Call Feuille.getcellbyposition(2, iFila - 1).setFormula("DNI")
        Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Apellido Paterno")
        Call Feuille.getcellbyposition(4, iFila - 1).setFormula("Apellido Materno")
        Call Feuille.getcellbyposition(5, iFila - 1).setFormula("Primer Nombre")
        Call Feuille.getcellbyposition(6, iFila - 1).setFormula("Servicio")
        Call Feuille.getcellbyposition(7, iFila - 1).setFormula("Fecha Cita")
        Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Hora Cita")
     Else
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
        oWorkSheet.Cells(iFila, 2).Value = "N° Historia Clinica"
        oWorkSheet.Cells(iFila, 3).Value = "DNI"
        oWorkSheet.Cells(iFila, 4).Value = "Apellido Paterno"
        oWorkSheet.Cells(iFila, 5).Value = "Apellido Materno"
        oWorkSheet.Cells(iFila, 6).Value = "Primer Nombre"
        oWorkSheet.Cells(iFila, 7).Value = "Servicio"
        oWorkSheet.Cells(iFila, 8).Value = "Fecha Cita"
        oWorkSheet.Cells(iFila, 9).Value = "Hora Cita"
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
     End If
     iFila = 6
     '
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighEntidades.CadenaConexion
     '
     oConexionExterna.CommandTimeout = 300
     oConexionExterna.CursorLocation = adUseClient
     oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexionExterna.BeginTrans
     Set oCitasWebCupos.Conexion = oConexionExterna
     '
     lnDiasMaximoConfirmacion = Val(lcBuscaParametro.SeleccionaFilaParametro(321))
     ldFechaHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL): lcHoraHoy = lcBuscaParametro.RetornaHoraServidorSQL
    '*************************************** Citas Web SisGalenPlus ***********************************
    lineas() = Split(LeerArchivoTexto(lcArchivoTexto), vbCrLf)
    lnRegistros = UBound(lineas)
    mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnRegistros: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
    
    Fichero = FreeFile
    Open lcArchivoTexto For Input As #Fichero
    lnNroRegistro = 1
    While Not EOF(Fichero)
       mo_ProgressRpt1.Value = lnNroRegistro: lnNroRegistro = lnNroRegistro + 1
       DoEvents
       Line Input #Fichero, Registro
       Campos = Split(Registro, "|")
       lnIdWeb = 0
       ldFecha = 0: lnIdServicio = 0: lnIdMedico = 0: lcHoraInicio = "": lcHoraFinal = "": lnIdEstadoCitaWeb = 0
       lnIdCitaBloqueada = 0: lcDNI = "": lcApellidoPaterno = "": lcApellidoMaterno = "": lcPrimerNombre = ""
       lcSegundoNombre = "": lnIdTipoSexo = 0: ldFechaNacimiento = 0: lnUbigeo = 0
       For Columna = 0 To UBound(Campos)
           Select Case Columna
           Case 0
                lnIdWeb = Val(Campos(Columna))
           Case 1
                ldFecha = CDate(Campos(Columna))
           Case 2
                lnIdServicio = Val(Campos(Columna))
           Case 3
                lnIdMedico = Val(Campos(Columna))
           Case 4
                lcHoraInicio = Campos(Columna)
           Case 5
                lcHoraFinal = Campos(Columna)
           Case 6
                lnIdEstadoCitaWeb = Val(Campos(Columna))
           Case 7
                lnIdCitaBloqueada = Val(Campos(Columna))
           Case 8
                lcDNI = Campos(Columna)
           Case 9
                lcApellidoPaterno = Campos(Columna)
           Case 10
                lcApellidoMaterno = Campos(Columna)
           Case 11
                lcPrimerNombre = Campos(Columna)
           Case 12
                lcSegundoNombre = Campos(Columna)
           Case 13
                lnIdTipoSexo = Val(Campos(Columna))
           Case 14
                If Campos(Columna) <> "" Then
                   ldFechaNacimiento = CDate(Campos(Columna))
                End If
           Case 15
                lnUbigeo = Val(Campos(Columna))
           Case 16
                If Campos(Columna) <> "" Then
                   ldFechaConfirmacion = CDate(Campos(Columna))
                End If
           Case 17
                lcHoraConfirmacion = Campos(Columna)
           Case 18
                lnIdFuenteFinanciamiento = Val(Campos(Columna))
           Case 19
                lnIdTurno = Val(Campos(Columna))
           Case 20
                lcEmail = Left(Campos(Columna), 50)
           Case 21
                lcTelefono = Left(Campos(Columna), 10)
                
           End Select
       Next
       If lnIdCitaBloqueada > 0 And ldFechaHoy <= ldFecha Then
           If lnIdEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                lnDiasQfaltaPcita = DateDiff("d", CDate(Format(ldFechaConfirmacion, sighEntidades.DevuelveFechaSoloFormato_DMY) & _
                                             " " & lcHoraConfirmacion), _
                                            CDate(Format(ldFecha, sighEntidades.DevuelveFechaSoloFormato_DMY) & _
                                            " " & lcHoraInicio))
                Set oDOCitasWebCupos = oCitasWebCupos.SeleccionarPorIdCitaBloqueada(lnIdCitaBloqueada)
                If oDOCitasWebCupos.idCitaBloqueada > 0 Then
                    If lnDiasQfaltaPcita < lnDiasMaximoConfirmacion Then
                       'elimina cupo WEB
                       oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoANULADO
                       oDOCitasWebCupos.IdUsuarioAuditoria = 0
                       If oCitasWebCupos.Modificar(oDOCitasWebCupos) = False Then
                          MsgBox oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
                       End If
                       If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(lnIdCitaBloqueada, oConexion) = False Then
                          MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                       End If
                    Else
                       'asigna al cupo reservado la cita web, incluyendo el paciente web
                       '(si es necesario crea HISTORIA)
                       oDOCitasWebCupos.idCitaBloqueada = lnIdCitaBloqueada
                       oDOCitasWebCupos.idWeb = lnIdWeb
                       oDOCitasWebCupos.idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoYconCitaEnGalenhos
                       oDOCitasWebCupos.DNI = lcDNI
                       oDOCitasWebCupos.ApellidoPaterno = lcApellidoPaterno
                       oDOCitasWebCupos.ApellidoMaterno = lcApellidoMaterno
                       oDOCitasWebCupos.PrimerNombre = lcPrimerNombre
                       oDOCitasWebCupos.SegundoNombre = lcSegundoNombre
                       oDOCitasWebCupos.idTipoSexo = lnIdTipoSexo
                       oDOCitasWebCupos.FechaNacimiento = ldFechaNacimiento
                       oDOCitasWebCupos.Ubigeo = lnUbigeo
                       oDOCitasWebCupos.FechaConfirmacion = ldFechaConfirmacion
                       oDOCitasWebCupos.HoraConfirmacion = lcHoraConfirmacion
                       oDOCitasWebCupos.IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
                       oDOCitasWebCupos.IdTurno = lnIdTurno
                       oDOCitasWebCupos.Email = lcEmail
                       oDOCitasWebCupos.TELEFONO = lcTelefono
                       If oCitasWebCupos.Modificar(oDOCitasWebCupos) = False Then
                          MsgBox oCitasWebCupos.MensajeError: GoTo ErrImpCWEB
                       End If
                       If mo_ReglasDeProgMedica.CrearCitaAutomatica(lcDNI, ldFecha, lcHoraInicio, lcHoraFinal, lnIdServicio, lnIdMedico, lcApellidoMaterno, _
                                      lcApellidoPaterno, ldFechaNacimiento, lnUbigeo, lnIdTipoSexo, lcPrimerNombre, _
                                      lcSegundoNombre, lnIdFuenteFinanciamiento, ldFechaHoy, lcHoraHoy, _
                                      lnHistoriaCreada, oConexion, ml_idUsuario, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                                      False, oDOCitasWebCupos.idPaciente, lnIdAtencion, 0, "", "", 0, "") = False Then
                          MsgBox ms_MensajeError: GoTo ErrImpCWEB
                       End If
                       If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(lnIdCitaBloqueada, oConexion) = False Then
                          MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                       End If
                       If lnHistoriaCreada > 0 Then
                          If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("'" & lnHistoriaCreada)
                                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(lcDNI)
                                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcApellidoPaterno)
                                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(lcApellidoMaterno)
                                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(lcPrimerNombre)
                                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mo_ReglasServiciosHosp.ServiciosDevuelveNombre(lnIdServicio))
                                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(ldFecha)
                                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(lcHoraInicio)

                          Else
                                oWorkSheet.Cells(iFila, 2).Value = "'" & lnHistoriaCreada
                                oWorkSheet.Cells(iFila, 3).Value = lcDNI
                                oWorkSheet.Cells(iFila, 4).Value = lcApellidoPaterno
                                oWorkSheet.Cells(iFila, 5).Value = lcApellidoMaterno
                                oWorkSheet.Cells(iFila, 6).Value = lcPrimerNombre
                                oWorkSheet.Cells(iFila, 7).Value = mo_ReglasServiciosHosp.ServiciosDevuelveNombre(lnIdServicio)
                                oWorkSheet.Cells(iFila, 8).Value = ldFecha
                                oWorkSheet.Cells(iFila, 9).Value = lcHoraInicio
                          End If
                          iFila = iFila + 1
                       End If
                    End If
                End If
                oRsTmp1.Close
           End If
        End If
    Wend
    Close #Fichero
    '
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName(mo_ReglasReportes.BuscaNombreColumna(2) & CStr(iFila) & ":" & mo_ReglasReportes.BuscaNombreColumna(11) & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
        oExcel.Visible = True
        oWorkSheet.PrintPreview
    End If
    '
    oConexionExterna.CommitTrans
    oConexionExterna.Close
    oConexion.Close
    '
    ImportaCitasWeb = True
ErrImpCWEB:
    If ImportaCitasWeb = False Then
        oConexionExterna.RollbackTrans
        oConexionExterna.Close
        oConexion.Close
    End If
    Set mo_ReglasDeProgMedica = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set oConexionExterna = Nothing
    Set oConexion = Nothing
    Set oRsTmp1 = Nothing
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        'encabezado de pagina
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
    'Resume
End Function



'Function CrearCitaAutomatica(lcDNI As String, ldFechaCita As Date, lcHoraInicio As String, lcHoraFinal As String, lnIdServicio As Long, _
'                   lnIdMedico As Long, lcApellidoMaterno, lcApellidoPaterno As String, ldFechaNacimiento As Date, _
'                   lnUbigeo As Long, lnIdTipoSexo As Long, lcPrimerNombre As String, lcSegundoNombre As String, _
'                   lnIdFuenteFinanciamiento As Long, ldFechaHoy As Date, lcHoraHoy As String, _
'                   ByRef lnHistoriaCreada As Long, oConexion As Connection) As Boolean
'     On Error GoTo ErrCrearCitaWeb
'     ms_MensajeError = ""
'     CrearCitaAutomatica = False
'     Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
'     Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
'     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
'     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
'     Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
'     Dim mo_CuentasAtencion As New DOCuentaAtencion, mo_Atenciones As New DOAtencion, mo_Pacientes As New doPaciente
'     Dim mo_DOFacturacionPaquetes As New DOFacturacionPaquetes, oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
'     Dim mo_Cita As New DOCita, mo_Historia As New DOHistoriaClinica, mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
'     Dim mo_FacturacionServicios As New Collection
'     Dim oPacientes As New Pacientes, oHistoria As New HistoriasClinicas
'     Dim oRsTmp21 As New Recordset
'     Dim lnTipoNumeracionAnterior As Long, lnIdHistoriaClinicaAnterior As Long, lbYaSeTransfirioHCdeUnServicioAotro As Boolean
'     Dim mo_Diagnosticos As New Collection, mo_Procedimientos As New Collection, mo_Examenes As New Collection
'     Dim lnIdTipoFinanciamiento As Long, lnIdProductoCpt As Long, lnIdEspecialidadMedico As Long, lnIdProgramacion As Long
'     Dim oEdad As Edad, lbNecesitaCrearHistoria As Boolean
'     Dim lnIdCondicionAlServicio As Long, lnIdCondicionAlEstablecimiento As Long
'     '
'
'     '
'     Set oRsTmp21 = mo_ReglasFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(lnIdFuenteFinanciamiento)
'     lnIdTipoFinanciamiento = 0
'     If oRsTmp21.RecordCount > 0 Then
'        lnIdTipoFinanciamiento = oRsTmp21.Fields!idTipoFinanciamiento
'     End If
'     oRsTmp21.Close
'     '
'     lnIdProductoCpt = 0
'     Set oRsTmp21 = mo_ReglasServiciosHosp.EspecialidadCESeleccionarPorIdServicio(lnIdServicio)
'     If oRsTmp21.RecordCount > 0 Then
'        lnIdProductoCpt = oRsTmp21.Fields!IdProductoConsulta
'     End If
'     oRsTmp21.Close
'     '
'     lnIdEspecialidadMedico = 0: lnIdProgramacion = 0
'     Set oRsTmp21 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(lnIdMedico, Format(ldFechaCita, "dd/mm/yyyy"), lnIdServicio)
'     If oRsTmp21.RecordCount > 0 Then
'        lnIdEspecialidadMedico = oRsTmp21.Fields!IdEspecialidad: lnIdProgramacion = oRsTmp21.Fields!IdProgramacion
'     End If
'     oRsTmp21.Close
'     '
'     Set mo_Pacientes = mo_AdminAdmision.PacientesSeleccionarPorDNI(lcDNI)
'     If mo_Pacientes.idPaciente = 0 Then
'        lbNecesitaCrearHistoria = True
'        mo_Pacientes.ApellidoMaterno = UCase(lcApellidoMaterno)
'        mo_Pacientes.ApellidoPaterno = UCase(lcApellidoPaterno)
'        mo_Pacientes.FechaNacimiento = ldFechaNacimiento
'        mo_Pacientes.IdDistritoDomicilio = lnUbigeo
'        mo_Pacientes.IdDocIdentidad = 1
'        mo_Pacientes.idPaciente = 0
'        mo_Pacientes.idTipoSexo = lnIdTipoSexo
'        mo_Pacientes.IdUsuarioAuditoria = ml_idUsuario
'        mo_Pacientes.NroDocumento = lcDNI
'        mo_Pacientes.NroHistoriaClinica = 0
'        mo_Pacientes.PrimerNombre = UCase(lcPrimerNombre)
'        mo_Pacientes.SegundoNombre = lcSegundoNombre
'        mo_Pacientes.IdTipoNumeracion = 1
'        mo_Pacientes.Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(mo_Pacientes)
'     End If
'     oEdad = CalcularEdad(mo_Pacientes.FechaNacimiento, CDate(Format(ldFechaCita, "dd/mm/yyyy") & " " & lcHoraInicio))
'     mo_ReglasComunes.TiposCondicionPacienteCondicionAlEstablecimientoYservicio lnIdCondicionAlEstablecimiento, _
'                              lnIdCondicionAlServicio, mo_Pacientes.idPaciente, _
'                              Format(ldFechaCita, sighentidades.DevuelveFechaSoloFormato_DMY), 0, lnIdServicio
'     If mo_Pacientes.idPaciente = 0 Then
'        mo_Historia.FechaCreacion = ldFechaHoy
'        mo_Historia.IdEstadoHistoria = 1
'        mo_Historia.idPaciente = 0
'        mo_Historia.IdTipoNumeracion = 1
'        mo_Historia.IdUsuarioAuditoria = ml_idUsuario
'        mo_Historia.NroHistoriaClinica = 0
'        mo_Historia.IdTipoHistoria = 1
'     Else
'        mo_Historia.NroHistoriaClinica = mo_Pacientes.NroHistoriaClinica
'        Set oHistoria.Conexion = oConexion
'        If oHistoria.SeleccionarPorId(mo_Historia) = False Then
'           ms_MensajeError = oHistoria.MensajeError: GoTo ErrCrearCitaWeb
'        End If
'     End If
'     '
'     mo_CuentasAtencion.FechaApertura = ldFechaHoy
'     mo_CuentasAtencion.FechaCreacion = ldFechaHoy
'     mo_CuentasAtencion.HoraApertura = lcHoraHoy
'     mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghAbierto
'     mo_CuentasAtencion.idPaciente = mo_Pacientes.idPaciente
'     mo_CuentasAtencion.IdUsuarioAuditoria = ml_idUsuario
'     '
'     mo_Atenciones.Edad = oEdad.Edad
'     mo_Atenciones.EsPacienteExterno = False
'     mo_Atenciones.FechaIngreso = ldFechaCita
'     mo_Atenciones.HoraIngreso = lcHoraInicio
'     mo_Atenciones.idAtencion = 0
'     mo_Atenciones.idCuentaAtencion = 0
'     mo_Atenciones.IdEspecialidadMedico = lnIdEspecialidadMedico
'     mo_Atenciones.IdEstadoAtencion = sghEstadoTabla.sghRegistrado
'     mo_Atenciones.IdFormaPago = lnIdTipoFinanciamiento
'     mo_Atenciones.idFuenteFinanciamiento = lnIdFuenteFinanciamiento
'     mo_Atenciones.IdMedicoIngreso = lnIdMedico
'     'mo_Atenciones.IdMedicoEgreso
'     mo_Atenciones.IdOrigenAtencion = 10   'domicilio
'     mo_Atenciones.idPaciente = mo_Pacientes.idPaciente
'     mo_Atenciones.IdServicioIngreso = lnIdServicio
'     mo_Atenciones.IdTipoEdad = oEdad.TipoEdad
'     mo_Atenciones.idTipoServicio = 1
'     mo_Atenciones.IdUsuarioAuditoria = ml_idUsuario
'     mo_Atenciones.IdTipoCondicionALEstab = lnIdCondicionAlEstablecimiento
'     mo_Atenciones.IdTipoCondicionAlServicio = lnIdCondicionAlServicio
'     '
'     mo_Cita.fecha = ldFechaCita
'     mo_Cita.FechaSolicitud = ldFechaCita
'     mo_Cita.HoraFin = lcHoraFinal
'     mo_Cita.HoraInicio = lcHoraInicio
'     mo_Cita.HoraSolicitud = lcHoraInicio
'     mo_Cita.idAtencion = 0
'     mo_Cita.IdCita = 0
'     mo_Cita.IdEspecialidad = lnIdEspecialidadMedico
'     mo_Cita.IdEstadoCita = 1
'     mo_Cita.idMedico = lnIdMedico
'     mo_Cita.idPaciente = mo_Pacientes.idPaciente
'     mo_Cita.idProducto = lnIdProductoCpt
'     mo_Cita.IdProgramacion = lnIdProgramacion
'     mo_Cita.idServicio = lnIdServicio
'     mo_Cita.IdUsuarioAuditoria = ml_idUsuario
'     '
'     'mo_Diagnosticos               'no hay datos
'     'mo_Procedimientos             'no hay datos
'     'mo_Examenes                   'no hay datos
'     'mo_DOFacturacionPaquetes      'no hay datos
'     'oDoSunasaPacientesHistoricos  'no hay datos
'     '
'     Set mo_FacturacionServicios = New Collection
'     mo_FacturacionServicios.Add CargarTipoDeConsulta(lnIdTipoFinanciamiento, lnIdFuenteFinanciamiento, lnIdProductoCpt)
'     '
'     lnTipoNumeracionAnterior = 1: lnIdHistoriaClinicaAnterior = 0: lbYaSeTransfirioHCdeUnServicioAotro = False
'
'     If mo_AdminAdmision.AdmisionCEAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, mo_Historia, _
'                       lnTipoNumeracionAnterior, mo_Diagnosticos, mo_Procedimientos, mo_Examenes, _
'                       lnIdHistoriaClinicaAnterior, mo_FacturacionServicios, mo_DOFacturacionPaquetes, _
'                       lbYaSeTransfirioHCdeUnServicioAotro, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
'                       "Agrega Cita Web", oDoSunasaPacientesHistoricos, mo_DoAtencionDatosAdicionales) = False Then
'        ms_MensajeError = mo_AdminAdmision.MensajeError: GoTo ErrCrearCitaWeb
'     End If
'     If lbNecesitaCrearHistoria = True Then
'        lnHistoriaCreada = mo_Pacientes.NroHistoriaClinica
'     Else
'        lnHistoriaCreada = 0
'     End If
'     CrearCitaAutomatica = True
'ErrCrearCitaWeb:
'     Set mo_AdminAdmision = Nothing
'     Set mo_ReglasFacturacion = Nothing
'     Set mo_ReglasServiciosHosp = Nothing
'     Set mo_ReglasDeProgMedica = Nothing
'     Set mo_ReglasComunes = Nothing
'     Set mo_CuentasAtencion = Nothing
'     Set mo_Atenciones = Nothing
'     Set mo_Pacientes = Nothing
'     Set mo_DOFacturacionPaquetes = Nothing
'     Set oDoSunasaPacientesHistoricos = Nothing
'     Set mo_Cita = Nothing
'     Set mo_Historia = Nothing
'     Set mo_DoAtencionDatosAdicionales = Nothing
'     Set mo_FacturacionServicios = Nothing
'     Set oPacientes = Nothing
'     Set oHistoria = Nothing
'     Set oRsTmp21 = Nothing
'End Function



'Function CargarTipoDeConsulta(lnIdTipoFinanciamiento As Long, lnIdFuenteFinanciamiento As Long, _
'                              lnIdProductoCpt As Long) As DOFacturacionServicios
'Dim oConsulta As New DOFacturacionServicios
'Dim oRsBuscaSeguro As New ADODB.Recordset
'Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
'Dim PrSeguro As Double
'        With oConsulta
'            .idAtencion = 0
'            .IdFacturacionServicio = 0
'            .idTipoFinanciamiento = lnIdTipoFinanciamiento
'            .idFuenteFinanciamiento = lnIdFuenteFinanciamiento
'            .Cantidad = 1
'            .idProducto = lnIdProductoCpt
'            .IdUsuarioAuditoria = ml_idUsuario
'            .IdEstadoFacturacion = sghEstadoFacturacion.sghRegistrado
'            .FechaAutorizaPendiente = 0
'            .FechaAutorizaSeguro = 0
'            .IdCentroCosto = 0
'            .IdUsuarioAutorizaPendiente = 0
'            .IdUsuarioAutorizaSeguro = 0
'            .PrecioUnitario = 0
'            .TotalPorPagar = 0
'            .idPuntoCarga = 6
'            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
''            If mi_Opcion = sghModificar Then
''               .IdOrden = lnIdFactServicios
''            End If
'            PrSeguro = 0
'            If lnIdTipoFinanciamiento = 9 Then
'               'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
'               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, 1)
'            Else
'               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, lnIdTipoFinanciamiento)
'            End If
'            If oRsBuscaSeguro.RecordCount > 0 Then
'               PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
'            End If
'            '.idTipoFinanciamiento = lnIdTipoFinanciamiento
'            .CantidadSIS = 0
'            .PrecioSIS = 0
'            .ImporteSIS = 0
'            .CantidadSOAT = 0
'            .PrecioSOAT = 0
'            .ImporteSOAT = 0
'            .importeEXO = 0
'            .cantidadConv = 0
'            .precConv = 0
'            .ImporteConv = 0
'            Select Case lnIdTipoFinanciamiento
'            Case 1  'Contado
'            Case 2  'SIS
'                 If PrSeguro > 0 Then
'                    .CantidadSIS = 1
'                    .PrecioSIS = PrSeguro
'                    .ImporteSIS = PrSeguro
'                    .IdEstadoFacturacion = 10
'                    .FechaAutorizaSeguro = Now
'                 Else
'                 End If
'            Case 3  'Soat
'                 If PrSeguro > 0 Then
'                    .CantidadSOAT = 1
'                    .PrecioSOAT = PrSeguro
'                    .ImporteSOAT = PrSeguro
'                    .IdEstadoFacturacion = 10
'                    .FechaAutorizaSeguro = Now
'                 Else
'                    .idTipoFinanciamiento = 1
'                 End If
'            Case 4  'Convenios
'                 If PrSeguro > 0 Then
'                    .cantidadConv = 1
'                    .precConv = PrSeguro
'                    .ImporteConv = PrSeguro
'                    .IdEstadoFacturacion = 10
'                    .PrecioUnitario = PrSeguro
'                    .FechaAutorizaSeguro = Now
'                 Else
'                    .idTipoFinanciamiento = 1
'                 End If
'            Case 9  'Exonerados
'                .importeEXO = PrSeguro
'                .IdEstadoFacturacion = 10
'                .idTipoFinanciamiento = 1
'                .FechaAutorizaEXO = Now
'            End Select
'            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
'        End With
'        Set CargarTipoDeConsulta = oConsulta
'        oRsBuscaSeguro.Close
'        Set oRsBuscaSeguro = Nothing
'        Set mo_AdminFacturacion = Nothing
'End Function

'debb-22/04/2016
Function ExportaCitasWebSoloTablas(lbEnviarServicio As Boolean, lbEnviarPacientes As Boolean, _
                                   lbEnviarFteFinanciamiento As Boolean, lbEnviarMedicos As Boolean, _
                                   lbEnviarTurnos As Boolean) As Boolean
            On Error GoTo ErrCitasWebExp1
            ExportaCitasWebSoloTablas = False
            Dim oConexion As New Connection, oConexionExterna As New Connection
            Dim oRsCitasWebCupos As New Recordset, oRsCitaBloqueada As New Recordset
            Dim oRsCitaBloqueada1 As New Recordset
            Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
            Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
            Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
            Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
            Dim lcRutaExportar As String, lnTotalRegistros As Long, lnIdCitaBloqueada As Long
            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
            'Exporta Servicios
            If lbEnviarServicio = True Then
                Set oRsCitasWebCupos = mo_ReglasArchivoClinico.ServiciosSeleccionarSoloCE
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Servicios.txt", oRsCitasWebCupos.RecordCount, 3, "|", Chr(10)
            End If
            'Exporta Medicos
            If lbEnviarMedicos = True Then
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.MedicosSeleccionarTodosParaWeb
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Medicos.txt", oRsCitasWebCupos.RecordCount, 4, "|", Chr(10)
            End If
            'Exporta Pacientes
            If lbEnviarPacientes = True Then
                Set oRsCitasWebCupos = mo_ReglasAdmision.PacientesSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Pacientes.txt", oRsCitasWebCupos.RecordCount, 5, "|", Chr(10)
            End If
            'Exporta Fuente Financiamiento
            If lbEnviarFteFinanciamiento = True Then
                Set oRsCitasWebCupos = mo_ReglasComunes.FuentesFinanciamientoSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "ftefinan.txt", oRsCitasWebCupos.RecordCount, 6, "|", Chr(10)
            End If
            If lbEnviarTurnos = True Then
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Turnos.txt", oRsCitasWebCupos.RecordCount, 1, "|", Chr(10)
            End If
            MsgBox "Se exportó correctamente los archivos textos, en la Ruta: " & lcRutaExportar
            ExportaCitasWebSoloTablas = True
ErrCitasWebExp1:
       If ExportaCitasWebSoloTablas = True Then
       Else
       End If
       Set oRsCitasWebCupos = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       Set mo_ReglasArchivoClinico = Nothing
       Set mo_ReglasAdmision = Nothing
       Set oRsCitaBloqueada1 = Nothing
End Function

Sub ExportaDAtosAlHISv4_2(lcTarde As String, lnIdMesElegido As Integer, lcAnio As String)
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
    Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
    Dim lcTDx1 As String, lcTDx2 As String, lcTDx3 As String, lcTDx4 As String, lcTDx5 As String, lcTDx6 As String
    Dim lcCDx1 As String, lcCDx2 As String, lcCDx3 As String, lcCDx4 As String, lcCDx5 As String, lcCDx6 As String
    Dim lcLDx1 As String, lcLDx2 As String, lcLDx3 As String, lcLDx4 As String, lcLDx5 As String, lcLDx6 As String
    Dim oRsTmp As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oRsTmpLote As New Recordset
    Dim oRsTmpCab As New Recordset
    Dim oRsTmp3 As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim oConexionFox As New ADODB.Connection
    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnContadorDetalle As Integer
    Dim rsreporte As New Recordset
    Dim lnTotal As Integer
    Dim Cod_Establec As Integer
    Dim Nombre_Lote As String
    Dim Num_pag As Integer
    Dim oRsFox As New Recordset
    Dim oRsFox2 As New Recordset
    Dim oRsFox3 As New Recordset
    Dim lcCod_dig As String
    Dim lcTotalRegHoja As String

    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    '
    oConexionFox.CommandTimeout = 300
    oConexionFox.Open "DSN=his"

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "His_ExportaDAtosAlHIS"
        Set oParameter = .CreateParameter("@IdMesElegido", adInteger, adParamInput, 0, lnIdMesElegido): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lcAnio): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "His_ExportaDAtosAlHISLote"
        Set oParameter = .CreateParameter("@IdMesElegido", adInteger, adParamInput, 0, lnIdMesElegido): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lcAnio): .Parameters.Append oParameter
        Set oRsTmpLote = .Execute
        Set oRsTmpLote.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "His_ExportaDAtosAlHISCabecera"
        Set oParameter = .CreateParameter("@IdMesElegido", adInteger, adParamInput, 0, lnIdMesElegido): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lcAnio): .Parameters.Append oParameter
        Set oRsTmpCab = .Execute
        Set oRsTmpCab.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing

'mo_ProgressRpt8
    lnTotal = 0
   If oRsTmp.RecordCount > 0 Or oRsTmpLote.RecordCount > 0 Or oRsTmpCab.RecordCount > 0 Then
        lnTotal = oRsTmp.RecordCount + oRsTmpLote.RecordCount + oRsTmpCab.RecordCount
        mo_ProgressRpt8.Min = 0: mo_ProgressRpt8.Max = lnTotal: mo_ProgressRpt8.ShowText = True: mo_ProgressRpt8.Color = vbGreen
        lnTotal = 0
   Else
        mo_ProgressRpt8.Min = 0
        mo_ProgressRpt8.Max = 1
        mo_ProgressRpt8.Value = 1
   End If
   Dim lcCod_Establec As String
   Dim lcNombre_Lote As String
   Dim lnNum_pag As Integer
   Dim lnNroReg As Integer
   Dim lnCodDx As Integer
   Dim lcDNIpaciente As String
   lcCod_Establec = "":  lcNombre_Lote = "":   lnNum_pag = 0: lnNroReg = 0
   
   lcCod_dig = Val(lcBuscaParametro.SeleccionaFilaParametro(270))
   lcTotalRegHoja = Val(lcBuscaParametro.SeleccionaFilaParametro(272))

   Dim lcSql As String
   If oRsTmp.RecordCount > 0 Then
        'Inicializa tabla fox
        
'    lcNombre_Lote = "C:\Archivos de programa\Digital Works Corporation\GalenHos"  'App.Path
'    If EstadoDeArchivo(lcNombre_Lote & "\Archivos\2hisLote.dbf") Then Kill (lcNombre_Lote & "\Archivos\2hisLote.dbf")
'    FileCopy lcNombre_Lote & "\Archivos\hisTem3.dbf", lcNombre_Lote & "\Archivos\2hisLote.dbf"
'    lcSql = "delete from 2hisLote"
'    oRsFox3.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'    oRsFox3.Open "select * from 2hisLote", oConexionFox, adOpenKeyset, adLockOptimistic
    
    oRsFox3.Open "select * from hisLote", oConexionFox, adOpenKeyset, adLockOptimistic 'Actualizado 25092014

'    If EstadoDeArchivo(lcNombre_Lote & "\Archivos\2his1.dbf") Then Kill (lcNombre_Lote & "\Archivos\2his1.dbf")
'    FileCopy lcNombre_Lote & "\Archivos\hisTem0.dbf", lcNombre_Lote & "\Archivos\2his1.dbf"
'    lcSql = "delete from 2his1"
'    oRsFox2.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'    oRsFox2.Open "select * from 2his1", oConexionFox, adOpenKeyset, adLockOptimistic
    
    oRsFox2.Open "select * from his1", oConexionFox, adOpenKeyset, adLockOptimistic 'Actualizado 25092014
    
'    If EstadoDeArchivo(lcNombre_Lote & "\Archivos\2hisa.dbf") Then Kill (lcNombre_Lote & "\Archivos\2hisa.dbf")
'    FileCopy lcNombre_Lote & "\Archivos\hisTem1.dbf", lcNombre_Lote & "\Archivos\2hisa.dbf"
'    lcSql = "delete from 2hisa"
'    oRsFox.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
'    oRsFox.Open "select * from 2hisA", oConexionFox, adOpenKeyset, adLockOptimistic
    
    oRsFox.Open "select * from hisA", oConexionFox, adOpenKeyset, adLockOptimistic 'Actualizado 25092014
        
        oRsTmp.MoveFirst
        lnNroReg = 1
        Do While Not oRsTmp.EOF
            oRsFox.AddNew
            oRsFox.Fields!Cod_2000 = String(9 - Len(Trim(oRsTmp.Fields!codigo_estabec)), "0") & Trim(oRsTmp.Fields!codigo_estabec)
            oRsFox.Fields!ano = oRsTmp.Fields!ano
            oRsFox.Fields!Mes = oRsTmp.Fields!Mes
            oRsFox.Fields!nom_lote = Left(oRsTmp.Fields!nom_lote, 3)
            oRsFox.Fields!Num_pag = oRsTmp.Fields!NroHojaHis
            oRsFox.Fields!num_reg = oRsTmp.Fields!NroRegistroHoja
            oRsFox.Fields!dia = oRsTmp.Fields!diaatencion
            oRsFox.Fields!fichafam = IIf(IsNull(oRsTmp.Fields!NroHC_FF), "", Mid(oRsTmp.Fields!NroHC_FF, 1, 10))
            oRsFox.Fields!Cod_Dpto = Left(oRsTmp.Fields!IdDistrito, 2)
            oRsFox.Fields!Cod_Prov = Mid(oRsTmp.Fields!IdDistrito, 3, 2)
            oRsFox.Fields!Cod_Dist = Right(oRsTmp.Fields!IdDistrito, 2)
            oRsFox.Fields!Edad = IIf(IsNull(oRsTmp.Fields!Edad), 0, oRsTmp.Fields!Edad)
            oRsFox.Fields!tip_edad = IIf(IsNull(oRsTmp.Fields!tip_edad), "", oRsTmp.Fields!tip_edad)
            oRsFox.Fields!Sexo = IIf(IsNull(oRsTmp.Fields!Edad), "", IIf(oRsTmp.Fields!Sexo = 1, "M", "F"))
            oRsFox.Fields!establec = IIf(IsNull(oRsTmp.Fields!IdTipoCondicionALEstab), "", IIf(oRsTmp.Fields!IdTipoCondicionALEstab = 1, "N", IIf(oRsTmp.Fields!IdTipoCondicionALEstab = 2, "R", "C")))
            oRsFox.Fields!Servicio = IIf(IsNull(oRsTmp.Fields!IdTipoCondicionAlServicio), "", IIf(oRsTmp.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(oRsTmp.Fields!IdTipoCondicionAlServicio = 2, "R", "C")))
             With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "His_ConsultarDxPorAtencion"
                 Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, Val(oRsTmp.Fields!IdHisDetalle)): .Parameters.Append oParameter
                 Set oRsTmp2 = .Execute
                 Set oRsTmp2.ActiveConnection = Nothing
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            lnCodDx = 0
            lcTDx1 = ""
            lcLDx1 = ""
            lcDx1 = ""
            lcTDx2 = ""
            lcLDx2 = ""
            lcDx2 = ""
            lcTDx3 = ""
            lcLDx3 = ""
            lcDx3 = ""
            lcTDx4 = ""
            lcLDx4 = ""
            lcDx4 = ""
            lcTDx5 = ""
            lcLDx5 = ""
            lcDx5 = ""
            lcTDx6 = ""
            lcLDx6 = ""
            lcDx6 = ""
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
                lnCodDx = lnCodDx + 1
                Select Case lnCodDx
                Case 1
                    lcTDx1 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx1 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx1 = oRsTmp2.Fields!codigo1
                Case 2
                    lcTDx2 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx2 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx2 = oRsTmp2.Fields!codigo1
                Case 3
                    lcTDx3 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx3 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx3 = oRsTmp2.Fields!codigo1
                Case 4
                    lcTDx4 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx4 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx4 = oRsTmp2.Fields!codigo1
                Case 5
                    lcTDx5 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx5 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx5 = oRsTmp2.Fields!codigo1
                Case 6
                    lcTDx6 = IIf(oRsTmp2.Fields!diagnost1 = 101, "P", IIf(oRsTmp2.Fields!diagnost1 = 102, "D", "R"))
                    lcLDx6 = IIf(IsNull(oRsTmp2.Fields!labconf1), "", oRsTmp2.Fields!labconf1)
                    lcDx6 = oRsTmp2.Fields!codigo1
                End Select
            oRsTmp2.MoveNext
            Loop
            
            oRsFox.Fields!diagnost1 = lcTDx1
            oRsFox.Fields!labconf1 = lcLDx1
            oRsFox.Fields!codigo1 = lcDx1
            oRsFox.Fields!diagnost2 = lcTDx2
            oRsFox.Fields!labconf2 = lcLDx2
            oRsFox.Fields!codigo2 = lcDx2
            oRsFox.Fields!diagnost3 = lcTDx3
            oRsFox.Fields!labconf3 = lcLDx3
            oRsFox.Fields!codigo3 = lcDx3
            oRsFox.Fields!diagnost4 = lcTDx4
            oRsFox.Fields!labconf4 = lcLDx4
            oRsFox.Fields!codigo4 = lcDx4
            oRsFox.Fields!diagnost5 = lcTDx5
            oRsFox.Fields!labconf5 = lcLDx5
            oRsFox.Fields!codigo5 = lcDx5
            oRsFox.Fields!diagnost6 = lcTDx6
            oRsFox.Fields!labconf6 = lcLDx6
            oRsFox.Fields!codigo6 = lcDx6
            oRsFox.Fields!esta_reg = "2"
            oRsFox.Fields!mt = oRsTmp.Fields!mt
            
            lcDNIpaciente = Space$(3)                               '1)Pais de Procedencia
            If Not IsNull(oRsTmp.Fields!IdNacionalidad) Then
                Set oRsTmp3 = mo_ReglasServGeograf.PaisesXfiltro("idPais=" & oRsTmp.Fields!IdNacionalidad)
                If oRsTmp3.RecordCount > 0 Then
                   lcDNIpaciente = Right(Space$(3) & oRsTmp3.Fields!Codigo, 3)
                End If
                oRsTmp3.Close
            End If
            If Not IsNull(oRsTmp.Fields!idTipoDocumento) Then    '2)Tipo de Documento de Identidad
               lcDNIpaciente = lcDNIpaciente & oRsTmp.Fields!idTipoDocumento
            Else
               lcDNIpaciente = lcDNIpaciente & IIf(IsNull(oRsTmp.Fields!DNI), Space$(1), "1")
            End If
            If Not IsNull(oRsTmp.Fields!DNI) Then         '3)Numero de Documento de identidad
               lcDNIpaciente = lcDNIpaciente & Right(Trim(oRsTmp.Fields!DNI), 8)
            Else
               lcDNIpaciente = lcDNIpaciente & Space$(8)
            End If
            lcDNIpaciente = lcDNIpaciente & Space$(4)                  '4)CUATRO espacios en BLANCO
                                                                      '5)Numero de Hijos
            If Trim(oRsTmp.Fields!nrohijo) = "" Or IsNull(oRsTmp.Fields!nrohijo) Then
               lcDNIpaciente = lcDNIpaciente & "00"
            Else
               lcDNIpaciente = lcDNIpaciente & Right("0" & Trim(Str(oRsTmp.Fields!nrohijo)), 2)
            End If
            
            oRsFox.Fields!DNI = lcDNIpaciente
            
'            If IsNull(oRsTmp.Fields!IdTipoDocumento) Then
'                oRsFox.Fields!DNI = ""
'            Else
'                If IsNull(oRsTmp.Fields!DNI) Then
'                    oRsFox.Fields!DNI = ""
'                Else
'                    If Trim(oRsTmp.Fields!nrohijo) = "" Or IsNull(oRsTmp.Fields!nrohijo) Then
'                        oRsFox.Fields!DNI = CStr(oRsTmp.Fields!IdTipoDocumento) & Left(oRsTmp.Fields!DNI, 8) & "00"
'                    Else
'                        oRsFox.Fields!DNI = CStr(oRsTmp.Fields!IdTipoDocumento) & Left(oRsTmp.Fields!DNI, 8) & Format(CStr(oRsTmp.Fields!nrohijo), "00")
'                    End If
'                End If
'            End If
            
            oRsFox.Fields!FI = IIf(IsNull(oRsTmp.Fields!FI), "", oRsTmp.Fields!FI)
            oRsFox.Fields!et = IIf(IsNull(oRsTmp.Fields!et), "", oRsTmp.Fields!et)
            oRsFox.Fields!ST = "1"
            oRsFox.Update
            
            lnTotal = lnTotal + 1
            mo_ProgressRpt8.Value = lnTotal
           
            oRsTmp.MoveNext
            If oRsTmp.EOF Then
               Exit Do
            End If
        Loop
          
        oRsTmpLote.MoveFirst
        Do While Not oRsTmpLote.EOF
            oRsFox3.AddNew
            oRsFox3.Fields!Cod_2000 = String(9 - Len(Trim(oRsTmpLote.Fields!Codigo)), "0") & Trim(oRsTmpLote.Fields!Codigo)
            oRsFox3.Fields!ano = oRsTmpLote.Fields!ano
            oRsFox3.Fields!Mes = oRsTmpLote.Fields!Mes
            oRsFox3.Fields!nom_lote = Trim(oRsTmpLote.Fields!nom_lote)
            oRsFox3.Fields!cod_dig = lcCod_dig
            oRsFox3.Fields!esta_lote = "LC"
            oRsFox3.Fields!tot_pag = oRsTmpLote.Fields!tot_pag
            oRsFox3.Fields!Tot_reg = lcTotalRegHoja
            oRsFox3.Fields!Tot_rmalos = 0
            oRsFox3.Fields!tot_pgcarg = 0
            oRsFox3.Fields!esta_local = Space(1)
            oRsFox3.Fields!nomfile = Space(14)
            oRsFox3.Fields!flagReport = "3"
            oRsFox3.Fields!mt = "M" 'lcMt
            oRsFox3.Update
            
            lnTotal = lnTotal + 1
            mo_ProgressRpt8.Value = lnTotal
            oRsTmpLote.MoveNext
            If oRsTmpLote.EOF Then
               Exit Do
            End If
        Loop
        
        Dim lcCodif As String
        oRsTmpCab.MoveFirst
        Do While Not oRsTmpCab.EOF
            oRsFox2.AddNew
            oRsFox2.Fields!Cod_2000 = String(9 - Len(Trim(oRsTmpCab.Fields!Codigo)), "0") & Trim(oRsTmpCab.Fields!Codigo)
            oRsFox2.Fields!ano = oRsTmpCab.Fields!ano
            oRsFox2.Fields!Mes = oRsTmpCab.Fields!Mes
            oRsFox2.Fields!nom_lote = Trim(oRsTmpCab.Fields!nom_lote)
            oRsFox2.Fields!Num_pag = oRsTmpCab.Fields!Num_pag
            If IsNull(oRsTmpCab.Fields!idColegioHIS) Then
                lcCodif = oRsTmpCab.Fields!idTipoDocumento & Mid(Trim(oRsTmpCab.Fields!DNI), 1, 8) & "00"
            Else
                lcCodif = oRsTmpCab.Fields!idTipoDocumento & Mid(Trim(oRsTmpCab.Fields!DNI), 1, 8) & oRsTmpCab.Fields!idColegioHIS
            End If
            oRsFox2.Fields!Codif = lcCodif
            oRsFox2.Fields!cod_servsa = oRsTmpCab.Fields!codigoServicioHIS 'lcCod_servsa
            oRsFox2.Fields!plaza = lcCodif
            oRsFox2.Fields!Esta_pag = "2"
            oRsFox2.Fields!Tot_reg = oRsTmpCab.Fields!Tot_reg
            oRsFox2.Fields!FlagEnvio = Space(1)
            oRsFox2.Fields!mt = "M" 'lcMt
            oRsFox2.Fields!ST = "1"
            oRsFox2.Update
            
            lnTotal = lnTotal + 1
            mo_ProgressRpt8.Value = lnTotal
            oRsTmpCab.MoveNext
            If oRsTmpCab.EOF Then
               Exit Do
            End If
        Loop
       
    End If
    If oRsFox.State = 1 Then oRsFox.Close
    If oRsFox2.State = 1 Then oRsFox2.Close
    If oRsFox3.State = 1 Then oRsFox3.Close
    oRsTmp.Close
    oRsTmpLote.Close
    oRsTmpCab.Close
End Sub

Public Function EstadoDeArchivo(ByVal Archivo As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(Archivo)) Then
    EstadoDeArchivo = True
    Else
    EstadoDeArchivo = False
    End If
End Function

'Frank 08/01/2014
Sub ExportaDAtosAlURENIS_V1(lcTarde As String, lnChkExportaCPT As Integer, lnIdMesElegido As Integer, lcAnio As String)
        On Error GoTo ErrorProceso
        Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
        Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
        Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
        
        Dim oConexion As New Connection
        Dim oConexionFox As New Connection
       
        Dim rsreporte As New Recordset
        Dim oRsTmp As New Recordset
        Dim oRsFox1 As New Recordset
        Dim oRsFox2 As New Recordset
        Dim oRsFox3 As New Recordset
        Dim oRsFox4 As New Recordset
        Dim oRsEdadHis As New Recordset
        Dim lcDx1 As String, lcDx2 As String, lcDx3 As String, lcDx4 As String, lcDx5 As String, lcDx6 As String
        Dim lcTDx1 As String, lcTDx2 As String, lcTDx3 As String, lcTDx4 As String, lcTDx5 As String, lcTDx6 As String
        Dim lcCDx1 As String, lcCDx2 As String, lcCDx3 As String, lcCDx4 As String, lcCDx5 As String, lcCDx6 As String
        Dim lcLDx1 As String, lcLDx2 As String, lcLDx3 As String, lcLDx4 As String, lcLDx5 As String, lcLDx6 As String
        Dim lnTot_reg As Long, lnEstancia As Integer
        Dim lcFec1 As String, lcFec2 As String
        Dim lcMes As String, lnEdadEnAtencion As Integer, lnAsc As Integer
        Dim lnUltimoDiaMes As Integer, lbContinuar As Boolean
        Dim lnIdEdad As Long, lnDesde As Integer, lnHasta As Integer
        Dim lcDistritoDomicilio As String, lcCodigoServicioHIS As String
        Dim lnTotal As Long, lnIdEmpleado As Long, lnNumPag As Long, lnReg As Long
        Dim lnTTot_pag As Long, lnTTot_reg As Long
'        dim lnMaxPacientesXhoja As Long, lnMaxHISxLOTE As Long,lcCod_dig As String,
        Dim lcCod_2000 As String, lcMt As String, lcNomLote As String
        Dim ldFecha As Date, lbAumentoPagina As Boolean, lbEsCentroSalud As Boolean
        Dim lcCodif As String, lcPlaza As String, lcCod_servsa As String, lcTipoDocumentoHisDoctor As String
        Dim lcFichaOhistoria As String, lcDNIpaciente As String, lcFuenteFinanciamientoPaciente As String, lcEtniaPaciente As String
        
        Dim lnGrupo6DxCPT As Integer
        Dim lnRestoDxCPT As Integer
        Dim lnIndiceDxCPT As Integer
        Dim lnContador As Integer
        Dim lnContDxCpt As Integer
        
        Const lcCharInicio As String = "0"
        '
        lcMes = Right("0" & Trim(Str(lnIdMesElegido)), 2)
        'lcAnio = cmbAnio.Text
        lcFec1 = ("01/" & lcMes & "/" & lcAnio)
        lnUltimoDiaMes = DevuelveUltimoDiaDelMes(Val(lcMes), Val(lcAnio))
        lcFec2 = Right("0" & Trim(Str(lnUltimoDiaMes)), 2) & "/" & lcMes & "/" & lcAnio
'        lnMaxPacientesXhoja = Val(lcBuscaParametro.SeleccionaFilaParametro(272))
'        lnMaxHISxLOTE = Val(lcBuscaParametro.SeleccionaFilaParametro(271))
'        lcCod_dig = Val(lcBuscaParametro.SeleccionaFilaParametro(270))
        lbEsCentroSalud = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
        '
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        '
        oConexionFox.CommandTimeout = 300
        oConexionFox.Open "DSN=urenis"
        '
        'Limpia tabla URENIS
        lcSql = "delete from urenis"
        oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
        oRsFox1.Open "select * from urenis", oConexionFox, adOpenKeyset, adLockOptimistic

        CrearTablaTemporalDxCPT 'Frank 12092014
        '
        Set rsreporte = mo_ReglasAdmision.AtencionesSeleccionarPorFechasDeIngresoYtipoServicioOconexion(lcMes, lcAnio, oConexion)
        lnTotal = rsreporte.RecordCount
        If lnTotal = 0 Then
           MsgBox "No existe Información  con esos Datos", vbInformation, "Mensaje"
        Else
           mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
           lnTotal = 0
           lcCod_2000 = Right("00000" & Trim(lcBuscaParametro.SeleccionaFilaParametro(280)), 5)
           ms_MensajeError = ""
           rsreporte.MoveFirst
           Do While Not rsreporte.EOF
'                If IsNull(rsReporte.Fields!LoteHis) Then
'                   ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'LOTE HIS' para el Médico: " & rsReporte.Fields!dEmpleado & Chr(13)
'                End If
                lcMt = IIf(rsreporte.Fields!HoraIngreso >= lcTarde, "T", "M")
                lcNomLote = rsreporte.Fields!loteHis & lcCharInicio    'rsReporte.Fields!LoteHis & lcMt
                lnIdEmpleado = rsreporte.Fields!IdEmpleado
                lbAumentoPagina = False
                   '
                    lnTTot_pag = 0: lnTTot_reg = 0
                    '
                    ldFecha = rsreporte.Fields!FechaIngreso
                    '
                    If IsNull(rsreporte.Fields!codigoServicioHIS) Then
                        ms_MensajeError = ms_MensajeError & "Tiene que ingresar el 'Codigo HIS' para el Servicio: " & rsreporte.Fields!DServicio & Chr(13)
                        MsgBox ms_MensajeError, vbInformation, "Mensajes"
                        Exit Sub
                        lcCodigoServicioHIS = ""
                    Else
                        lcCodigoServicioHIS = rsreporte.Fields!codigoServicioHIS
                    End If
                    lcCod_servsa = lcCodigoServicioHIS
                    '
                    If IsNull(rsreporte.Fields!DNI) Or IsNull(rsreporte.Fields!idColegioHIS) Then
                        ms_MensajeError = ms_MensajeError & "El profesional de Salud : " & rsreporte.Fields!dEmpleado & " deberá tener 'Colegio Médico' y 'DNI'" & Chr(13)
                        lcCodif = ""
                    Else
                        Set oRsTmp = mo_ReglasAdmision.TiposDocIdentidadXfiltro("idDocIdentidad=" & rsreporte.Fields!idTipoDocumento, oConexion)
                        lcTipoDocumentoHisDoctor = "1"
                        If oRsTmp.RecordCount > 0 Then
                           lcTipoDocumentoHisDoctor = oRsTmp.Fields!CodigoHIS
                        End If
                        oRsTmp.Close
                        lcCodif = lcTipoDocumentoHisDoctor & Left(rsreporte.Fields!DNI, 8) & rsreporte.Fields!idColegioHIS
                    End If
                    '
                    lcPlaza = lcCodif
                    '
                    '
                    If lbAumentoPagina = True Then
                       lnNumPag = lnNumPag + 1
                    Else
                       lnNumPag = 1
                    End If
                    lnTot_reg = 0: lnReg = 0
'                    Do While Not rsReporte.EOF And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoserviciohis And ldFecha = rsReporte.Fields!FechaIngreso
                        lcFuenteFinanciamientoPaciente = IIf(IsNull(rsreporte.Fields!HISfuenteFinanciamiento), "", rsreporte.Fields!HISfuenteFinanciamiento)
                        lcEtniaPaciente = IIf(IsNull(rsreporte.Fields!IdEtnia), "", rsreporte.Fields!IdEtnia)
                        '1112333333334455
'                        lcDNIpaciente = Space$(3)                               '1)Pais de Procedencia
'                        If Not IsNull(rsReporte.Fields!IdPaisProcedencia) Then
'                            Set oRsTmp = mo_ReglasServGeograf.PaisesXfiltro("idPais=" & rsReporte.Fields!IdPaisProcedencia)
'                            If oRsTmp.RecordCount > 0 Then
'                               lcDNIpaciente = Right(Space$(3) & oRsTmp.Fields!Codigo, 3)
'                            End If
'                            oRsTmp.Close
'                        End If
'                        If Not IsNull(rsReporte.Fields!HIStipoDocumento) Then    '2)Tipo de Documento de Identidad
'                           lcDNIpaciente = lcDNIpaciente & rsReporte.Fields!HIStipoDocumento
'                        Else
'                           lcDNIpaciente = lcDNIpaciente & IIf(IsNull(rsReporte.Fields!NroDocumento), Space$(1), "1")
'                        End If
'                        If Not IsNull(rsReporte.Fields!NroDocumento) Then         '3)Numero de Documento de identidad
'                           lcDNIpaciente = lcDNIpaciente & Right(Trim(rsReporte.Fields!NroDocumento), 8)
'                        Else
'                           lcDNIpaciente = lcDNIpaciente
'                        End If
'                        lcDNIpaciente = lcDNIpaciente
                                                                                  
'                        Set oRsTmp = mo_ReglasAdmision.atencionesDatosAdicionalesXfiltro("idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
'                        If oRsTmp.RecordCount > 0 Then
'                           If Not IsNull(oRsTmp.Fields!NumeroDeHijos) Then
'                               lcDNIpaciente = lcDNIpaciente & Right("0" & Trim(Str(oRsTmp.Fields!NumeroDeHijos)), 2)
'                           Else
'                               lcDNIpaciente = lcDNIpaciente & "00"
'                           End If
'                        Else
'                           lcDNIpaciente = lcDNIpaciente & "00"
'                        End If
'                        oRsTmp.Close
                        '
                        If rsreporte.Fields!NroHistoriaClinica = 511302 Then
                        lcFichaOhistoria = ""
                        End If
                        If IsNull(rsreporte.Fields!NroHistoriaClinica) And IsNull(rsreporte.Fields!FichaFamiliar) Then
                           MsgBox "El Paciente No existe para la CUENTA N° " & rsreporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
                           Exit Sub
                        End If
                        lcFichaOhistoria = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(rsreporte.Fields!NroHistoriaClinica)), False)
                         
                        If Not IsNull(rsreporte.Fields!FichaFamiliar) And lbEsCentroSalud = True Then
                           If Len(Trim(rsreporte.Fields!FichaFamiliar)) > 2 Then
                              lcFichaOhistoria = Trim(rsreporte.Fields!FichaFamiliar)
                           End If
                        End If
                        '
'                        lcDx1 = "": lcDx2 = "": lcDx3 = "": lcDx4 = "": lcDx5 = "": lcDx6 = ""
'                        lcTDx1 = "": lcTDx2 = "": lcTDx3 = "": lcTDx4 = "": lcTDx5 = "": lcTDx6 = ""
'                        lcCDx1 = "": lcCDx2 = "": lcCDx3 = "": lcCDx4 = "": lcCDx5 = "": lcCDx6 = ""
'                        lcLDx1 = "": lcLDx2 = "": lcLDx3 = "": lcLDx4 = "": lcLDx5 = "": lcLDx6 = ""
                        LimpiarTablaDxCpt
                        lnIndiceDxCPT = 0
                        
                        Set oRsTmp = mo_ReglasAdmision.AtencionesDiagnosticosXidAtencion(rsreporte.Fields!idAtencion, oConexion)
                        If oRsTmp.RecordCount > 0 Then
'                           lnUltimoDiaMes = 1
                           oRsTmp.MoveFirst
                           Do While Not oRsTmp.EOF
                              If IsNull(oRsTmp.Fields!DescripcionTipoDx) Then
                                 MsgBox "No existe DIAGNOSTICO para la CUENTA N° " & rsreporte.Fields!idCuentaAtencion & Chr(13) & "Revise la tabla ATENCIONES", vbInformation, "Mensaje"
                                 Exit Sub
                              End If
                              lnIndiceDxCPT = lnIndiceDxCPT + 1
                              oRsDxCpt.AddNew
                              oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
                              oRsDxCpt.Fields!Dx = sighEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp.Fields!CodigoCIE2004, 6))
                              oRsDxCpt.Fields!TDx = IIf(IsNull(oRsTmp.Fields!DescripcionTipoDx), "D", Left(oRsTmp.Fields!DescripcionTipoDx, 1))
                              oRsDxCpt.Fields!CDx = IIf(IsNull(oRsTmp.Fields!ClaseDxHIS), "1", Left(oRsTmp.Fields!ClaseDxHIS, 1))
                              oRsDxCpt.Fields!LDx = IIf(IsNull(oRsTmp.Fields!labConfHIS), "", oRsTmp.Fields!labConfHIS)
                              oRsDxCpt.Update
                              oRsTmp.MoveNext
                           Loop
                        End If
                        oRsTmp.Close
                        
                        'CPT
                        If lnChkExportaCPT = 1 Then
'                             Set oRsTmp = mo_ReglasAdmision.BuscaAtencionesCptCEparaFormatoHIS(rsReporte.Fields!idCuentaAtencion)
                             Set oRsTmp = mo_ReglasAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(rsreporte.Fields!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
                             If oRsTmp.RecordCount > 0 Then
                                  oRsTmp.MoveFirst
                                  Do While Not oRsTmp.EOF
                                    lnIndiceDxCPT = lnIndiceDxCPT + 1
                                    oRsDxCpt.AddNew
                                    oRsDxCpt.Fields!IdDxCpt = lnIndiceDxCPT
                                    oRsDxCpt.Fields!Dx = oRsTmp.Fields("codigo").Value
                                    oRsDxCpt.Fields!TDx = ""
                                    oRsDxCpt.Fields!CDx = ""
                                    oRsDxCpt.Fields!LDx = oRsTmp.Fields("cantidad").Value
                                    oRsDxCpt.Update
                                    oRsTmp.MoveNext
                                  Loop
                             End If
                             oRsTmp.Close
                        End If
                        '
                        lbContinuar = True
                        If oRsDxCpt.RecordCount = 0 Then
                            lbContinuar = False
                        End If
                        '
                        If lbContinuar = True Then
                            lcDistritoDomicilio = Space(6)
                            If Not IsNull(rsreporte.Fields!IdDistrito) Then
                               lcDistritoDomicilio = Right("0" & Trim(Str(rsreporte.Fields!IdDistrito)), 6)
                            End If
                            '
                            lnTot_reg = lnTot_reg + 1
                            lnReg = lnReg + 1
                            lnEdadEnAtencion = rsreporte.Fields!Edad
                            If lnEdadEnAtencion > 99 Then
                               lnEdadEnAtencion = 99
                            End If
                            lnGrupo6DxCPT = IIf(Round(oRsDxCpt.RecordCount / 6) < oRsDxCpt.RecordCount / 6, Round(oRsDxCpt.RecordCount / 6) + 1, Round(oRsDxCpt.RecordCount / 6))
                            lnRestoDxCPT = IIf(oRsDxCpt.RecordCount Mod 6 = 0, 6, oRsDxCpt.RecordCount Mod 6)
                            
                            For lnContador = 0 To 0 'lnGrupo6DxCPT - 1 'Si quiere mas de una fila Frank
                                lnIndiceDxCPT = lnContador * 6
                                
                                oRsFox1.AddNew
                                oRsFox1.Fields!renaes = lcCod_2000
                                oRsFox1.Fields!periodo = Year(rsreporte.Fields!FechaIngreso) & Month(rsreporte.Fields!FechaIngreso) & Day(rsreporte.Fields!FechaIngreso)
                                oRsFox1.Fields!id_cita = rsreporte.Fields!idAtencion
                                oRsFox1.Fields!id_ups = lcCod_servsa
                                oRsFox1.Fields!id_pac = lcFichaOhistoria
                                oRsFox1.Fields!codpais = rsreporte.Fields!IdPaisProcedencia
                                oRsFox1.Fields!tipdoc = IIf(IsNull(rsreporte.Fields!HIStipoDocumento), 99, rsreporte.Fields!HIStipoDocumento)
                                oRsFox1.Fields!nrodoc = IIf(IsNull(rsreporte.Fields!HIStipoDocumento), "", rsreporte.Fields!nrodocumento)
                                oRsFox1.Fields!Ubigeo = lcDistritoDomicilio
                                oRsFox1.Fields!Edad = lnEdadEnAtencion
                                oRsFox1.Fields!tip_edad = rsreporte.Fields!EdadCodigo
                                oRsFox1.Fields!Sexo = IIf(rsreporte.Fields!idTipoSexo = 1, "M", "F")
                                oRsFox1.Fields!conest = IIf(rsreporte.Fields!IdTipoCondicionALEstab = 1, "N", IIf(rsreporte.Fields!IdTipoCondicionALEstab = 2, "R", "C"))
                                oRsFox1.Fields!conser = IIf(rsreporte.Fields!IdTipoCondicionAlServicio = 1, "N", IIf(rsreporte.Fields!IdTipoCondicionAlServicio = 2, "R", "C"))
                                oRsFox1.Fields!id_turno = lcMt
                                oRsFox1.Fields!id_fin = lcFuenteFinanciamientoPaciente
                                oRsFox1.Fields!etnia = lcEtniaPaciente
                                
                                'Coloca por defecto los Dx vacios
                                oRsFox1.Fields!DIAG1 = ""
                                oRsFox1.Fields!CIEX1 = ""
                                oRsFox1.Fields!DIAG2 = ""
                                oRsFox1.Fields!CIEX2 = ""
                                oRsFox1.Fields!DIAG3 = ""
                                oRsFox1.Fields!CIEX3 = ""
                                oRsFox1.Fields!DIAG4 = ""
                                oRsFox1.Fields!CIEX4 = ""
                                oRsFox1.Fields!DIAG5 = ""
                                oRsFox1.Fields!CIEX5 = ""
                                oRsFox1.Fields!DIAG6 = ""
                                oRsFox1.Fields!CIEX6 = ""
                                
                                For lnContDxCpt = 1 To 6
                                    oRsDxCpt.MoveFirst
                                    oRsDxCpt.Find "IdDxCpt=" & lnIndiceDxCPT + lnContDxCpt
                                    If Not oRsDxCpt.EOF Then
                                        Select Case lnContDxCpt
                                        Case 1
                                            oRsFox1.Fields!DIAG1 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX1 = oRsDxCpt.Fields!Dx
                                        Case 2
                                            oRsFox1.Fields!DIAG2 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX2 = oRsDxCpt.Fields!Dx
                                        Case 3
                                            oRsFox1.Fields!DIAG3 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX3 = oRsDxCpt.Fields!Dx
                                        Case 4
                                            oRsFox1.Fields!DIAG4 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX4 = oRsDxCpt.Fields!Dx
                                        Case 5
                                            oRsFox1.Fields!DIAG5 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX5 = oRsDxCpt.Fields!Dx
                                        Case 6
                                            oRsFox1.Fields!DIAG6 = oRsDxCpt.Fields!TDx
                                            oRsFox1.Fields!CIEX6 = oRsDxCpt.Fields!Dx
                                        End Select
                                    End If
                                Next
                                oRsFox1.Fields!tipdocper = "1"
                                oRsFox1.Fields!nrodocper = Trim(rsreporte.Fields!DNI)
                                oRsFox1.Fields!nomesp = rsreporte.Fields!dEmpleado
                                oRsFox1.Fields!profesion = rsreporte.Fields!IdEspecialidadMedico
                                oRsFox1.Fields!colegio = rsreporte.Fields!idColegioHIS
                                oRsFox1.Update
                            Next
                            '
                            lnTotal = lnTotal + 1
                            mo_ProgressRpt1.Value = lnTotal
                            DoEvents
                        End If
                        
                        rsreporte.MoveNext
'                        If rsReporte.EOF Then
'                           Exit Do
'                        End If
'                        If lnReg >= lnMaxPacientesXhoja And lnIdEmpleado = rsReporte.Fields!IdEmpleado And lcCodigoServicioHIS = rsReporte.Fields!codigoserviciohis And ldFecha = rsReporte.Fields!FechaIngreso Then
'                           lbAumentoPagina = True
'                           lnReg = 0
'                           Exit Do
'                        Else
'                           lbAumentoPagina = False
'                        End If
'                    Loop
                    '

'                    If rsReporte.EOF Then
'                       Exit Do
'                    End If
'                Loop
           Loop
           Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "Atenciones", oConexion, 500 + 141, mo_lcNombrePc, "Exportar al URENIS: " & lcFec1 & " al " & lcFec2)     '500+ ListBarReporte.idReporte
        End If
        oRsFox1.Close
        oConexionFox.Close
        oConexion.Close
        If ms_MensajeError <> "" Then
           MsgBox ms_MensajeError, vbInformation, "Mensajes"
        End If
        Exit Sub
ErrorProceso:
    MsgBox Err.Description
    Resume
End Sub

'Actualiza valores y Devuelve percentil de la ATENCION ACTUAL DEL PACIENTE
'debb-2/3/2015
Sub CalculaPercentiles(lnPesoKg As Double, lnTallaCM As Long, ml_EdadEnMeses As Long, _
                       ml_idTipoSexo As Long, lnEdadEnAniosEnAtencion As Integer, _
                       ByRef lnPercentilPE As Long, ByRef lnPercentilTE As Long, ByRef lnPercentilPT As Long, ByRef lnPercentilIMC As Double, _
                       ByRef lnPercentilPE_Z As Double, ByRef lnPercentilTE_Z As Double, ByRef lnPercentilPT_Z As Double, ByRef lnPercentilIMC_Z As Double)
    Const lnPercentilNull As Long = 0
    lnPercentilPE = lnPercentilNull: lnPercentilTE = lnPercentilNull: lnPercentilPT = lnPercentilNull: lnPercentilIMC = lnPercentilNull
    lnPercentilPE_Z = lnPercentilNull: lnPercentilTE_Z = lnPercentilNull: lnPercentilPT_Z = lnPercentilNull: lnPercentilIMC_Z = lnPercentilNull
    If lnPesoKg > 0 And lnTallaCM > 0 Then
       On Error Resume Next
       Dim EXL As Excel.Application
       Set EXL = New Excel.Application
       Dim W As Excel.Workbook
       Dim lnEdadUSAmaxima
       lnEdadUSAmaxima = 5
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred.xls")       'usa
       Else
           Set W = EXL.Workbooks.Open(App.Path & "\Plantillas\cred who.xls")    'oms
       End If
       Dim s As Excel.Worksheet
       Dim lnEdadEnMesesMasPuntoCinco As Double, lnMinimo As Double, lnMaximo As Double, lnIMC As Double
       Dim lnTallaEnCmMasPuntoCinco As Double
       lnEdadEnMesesMasPuntoCinco = ml_EdadEnMeses + 0.5
       lnTallaEnCmMasPuntoCinco = lnTallaCM + 0.5
       'Peso Edad
       Set s = W.Sheets("P-E")
       
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If





       'Talla Edad
       Set s = W.Sheets("T-E")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(243, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(246, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(247, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaCM
          lnPercentilTE = s.Cells(255, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilTE_Z = s.Cells(254, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilTE = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaCM
          lnPercentilTE = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilTE_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       'Peso Talla
       Set s = W.Sheets("P-T")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(61, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(64, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(65, IIf(ml_idTipoSexo = 1, 4, 20)).Value = lnPesoKg
          lnPercentilPT = s.Cells(73, IIf(ml_idTipoSexo = 1, 3, 19)).Value
          lnPercentilPT_Z = s.Cells(72, IIf(ml_idTipoSexo = 1, 3, 19)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(652, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilPT = lnPercentilNull
          s.Cells(655, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnTallaEnCmMasPuntoCinco
          s.Cells(656, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnPesoKg
          lnPercentilPT = s.Cells(664, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilPT_Z = s.Cells(663, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If

       'Edad IMC
       lnIMC = Round((lnPesoKg / (lnTallaCM * lnTallaCM)), 0)
       Set s = W.Sheets("E-IMC")
       'If lnEdadEnAniosEnAtencion <= lnEdadUSAmaxima Then
       If lnEdadEnAniosEnAtencion > lnEdadUSAmaxima Then
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnMaximo = s.Cells(220, IIf(ml_idTipoSexo = 1, 2, 19)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(223, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(224, IIf(ml_idTipoSexo = 1, 4, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(232, IIf(ml_idTipoSexo = 1, 3, 20)).Value
          lnPercentilIMC_Z = s.Cells(231, IIf(ml_idTipoSexo = 1, 3, 20)).Value
       Else
          lnMinimo = s.Cells(2, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnMaximo = s.Cells(1858, IIf(ml_idTipoSexo = 1, 2, 18)).Value
          lnPercentilIMC = lnPercentilNull
          s.Cells(1861, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnEdadEnMesesMasPuntoCinco
          s.Cells(1862, IIf(ml_idTipoSexo = 1, 5, 21)).Value = lnIMC
          lnPercentilIMC = s.Cells(1870, IIf(ml_idTipoSexo = 1, 4, 20)).Value
          lnPercentilIMC_Z = s.Cells(1869, IIf(ml_idTipoSexo = 1, 4, 20)).Value
       End If
       W.Close False
       Set s = Nothing
       Set W = Nothing
       Set EXL = Nothing
    End If
End Sub


'debb-15/06/2016
Sub ExportaSEM2016(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, _
                    ml_Valor As Integer, ml_Tabla As String, lbMuestraUsuario As Boolean)
End Sub



'debb-15/06/2016
Sub ExportaSEM2017(lcFechaInicio As String, lcFechaFinal As String, lcCorrelativo As String, _
                    ml_Valor As Integer, ml_Tabla As String, lbMuestraUsuario As Boolean)
End Sub

Function CitasWebLista(lnIdEstado As sghCitaWebEstados) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set CitasWebLista = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CitasWebLista"
        Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, lnIdEstado): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CitasWebLista = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
