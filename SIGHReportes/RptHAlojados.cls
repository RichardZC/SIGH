VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptHAlojados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Alojados
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario  As Long
Const lnIdEspecialidad As Long = 23    'Neonatologia
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    MsgBox "Reporte en desarrollo..."
    Exit Sub
    
    
    Dim oFormulario As New HAlojados
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.IdEspecialidad = lnIdEspecialidad
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, ml_idServicio As Long, mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim mrs_Tmp3 As New Recordset
    Dim mrs_Tmp4 As New Recordset
    Dim mrs_Cab As New Recordset
    Dim mrs_Det As New Recordset
    Dim mrs_Shape As New Recordset
    Dim ldFecha As Date, ldFechaSolo As Date
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long, lnNroCuenta As Long, ldFechaIngreso As Date
    Dim lcHoraIngreso As String, lcPaciente As String, lcSexo As String
    Dim lcMotivo As String, lcEstancia As String, ldFechaEgreso As Date
    Dim lcHoraEgreso As String, lcDestino As String, lcRecibe As String
    Dim lcMedicoIE As String, lnImporte As Double, lcDescripcion As String
    Dim lcNombreTablaCab As String, lcNombreTablaDet As String
    Dim lnEstancia As Integer, lnIdServicio As Long, lnTotalG As Double
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
    
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Set mrs_Tmp1 = mo_AdminReportes.AtencionesSeleccionarPorFechaIngresoYespecialidadDelServicio(mda_FechaInicio, mda_FechaFin, lnIdEspecialidad)
    If ml_idServicio > 0 Then
       mrs_Tmp1.Filter = "idServicioIngreso=" & ml_idServicio
    End If
    If mrs_Tmp1.RecordCount = 0 Then
         MsgBox "No existe Datos", vbInformation, "Reporte"
    Else
         If lbEnExcel = True Then
            'Crear Tabla Temporal CabeceraDetalle para EXCEL --->>>>>>cuando se usa EXCEL
            With mrs_Tmp
               .Fields.Append "NroCuenta", adInteger
               .Fields.Append "FechaIngreso", adDate, 10, adFldIsNullable
               .Fields.Append "HoraIngreso", adVarChar, 5, adFldIsNullable
               .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
               .Fields.Append "Sexo", adVarChar, 1, adFldIsNullable
               .Fields.Append "Motivo", adVarChar, 100, adFldIsNullable
               .Fields.Append "Estancia", adVarChar, 30, adFldIsNullable
               .Fields.Append "FechaEgreso", adDate, 10, adFldIsNullable
               .Fields.Append "HoraEgreso", adVarChar, 5, adFldIsNullable
               .Fields.Append "Destino", adVarChar, 100, adFldIsNullable
               .Fields.Append "Recibe", adVarChar, 100, adFldIsNullable
               .Fields.Append "MedicoIE", adVarChar, 100, adFldIsNullable
               .Fields.Append "ConsumoDescripcion", adVarChar, 100, adFldIsNullable
               .Fields.Append "ConsumoImporte", adDouble
               .LockType = adLockOptimistic
               .Open
            End With
         Else
            mo_AdminReportes.PreparaTablasDeReportes mrs_Cab, mrs_Det
         End If
         '
         lnTotalG = 0
         mrs_Tmp1.MoveFirst
         Do While Not mrs_Tmp1.EOF

                lnNroCuenta = mrs_Tmp1.Fields!idCuentaAtencion
                ldFechaIngreso = mrs_Tmp1.Fields!FechaIngreso
                lcHoraIngreso = mrs_Tmp1.Fields!HoraIngreso
                lcPaciente = "(" & Trim(Str(mrs_Tmp1.Fields!nrohistoriaclinica)) & ") " & Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp1.Fields!PrimerNombre)
                lcSexo = IIf(mrs_Tmp1.Fields!idTipoSexo = 1, "M", "F")
                lcMotivo = mrs_Tmp1.Fields!motivo
                lcDescripcion = ""
                lnImporte = 0
                lcEstancia = ""
                ldFechaEgreso = 0
                lcHoraEgreso = ""
                lcDestino = ""
                lcRecibe = ""
                lcMedicoIE = ""
                If Not IsNull(mrs_Tmp1.Fields!FechaEgresoAdministrativo) Then
                    If mrs_Tmp1.Fields!FechaIngreso = mrs_Tmp1.Fields!FechaEgreso Then
                       lcEstancia = Trim(Str(HorasDelPacienteEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!FechaEgresoAdministrativo, mrs_Tmp1.Fields!horaEgresoAdministrativo))) & "  Hrs."
                    Else
                       lcEstancia = Trim(Str(DiasDeEstanciaEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, mrs_Tmp1.Fields!FechaEgresoAdministrativo, mrs_Tmp1.Fields!horaEgresoAdministrativo))) & "  Días"
                    End If
                    ldFechaEgreso = mrs_Tmp1.Fields!FechaEgresoAdministrativo
                    lcHoraEgreso = mrs_Tmp1.Fields!horaEgresoAdministrativo
                    lcDestino = mrs_Tmp1.Fields!destino
                    lcRecibe = mrs_Tmp1.Fields!NombreAcompaniante
                    lcMedicoIE = Trim(mrs_Tmp1.Fields!mPaterno) & " " & Trim(mrs_Tmp1.Fields!mMaterno) & " " & Trim(mrs_Tmp1.Fields!mNombres)
                    If lbEnExcel = False Then
                        mrs_Cab.AddNew
                        mrs_Cab.Fields!nroCuenta = lnNroCuenta
                        mrs_Cab.Fields!FechaIngreso = ldFechaIngreso
                        mrs_Cab.Fields!HoraIngreso = lcHoraIngreso
                        mrs_Cab.Fields!Paciente = lcPaciente
                        mrs_Cab.Fields!Sexo = lcSexo
                        mrs_Cab.Fields!motivo = lcMotivo
                        mrs_Cab.Fields!Estancia = lcEstancia
                        mrs_Cab.Fields!FechaEgreso = ldFechaEgreso
                        mrs_Cab.Fields!horaEgreso = lcHoraEgreso
                        mrs_Cab.Fields!destino = lcDestino
                        mrs_Cab.Fields!Recibe = lcRecibe
                        mrs_Cab.Fields!MedicoIE = lcMedicoIE
                        mrs_Cab.Update
                    End If
                    'servicios
                    Set mrs_Tmp2 = mo_ReglasFacturacion.FacturacionServicioPagosPorCuentaTodos(lnNroCuenta)
                    mrs_Tmp2.Filter = "idEstadoFacturacion=1"
                    If mrs_Tmp2.RecordCount > 0 Then
                       mrs_Tmp2.MoveFirst
                       Do While Not mrs_Tmp2.EOF
                            lnImporte = mrs_Tmp2.Fields!totalPorPagar
                            lcDescripcion = ""
                            Set mrs_Tmp3 = mo_reglasComunes.CatalogoServiciosSeleccionarXidentificador(mrs_Tmp2.Fields!idProducto)
                            If mrs_Tmp3.RecordCount > 0 Then
                               lcDescripcion = Trim(mrs_Tmp3.Fields!Codigo) & " " & mrs_Tmp3.Fields!Nombre
                            End If
                            mrs_Tmp3.Close
                            If lbEnExcel = True Then
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!nroCuenta = lnNroCuenta
                                mrs_Tmp.Fields!FechaIngreso = ldFechaIngreso
                                mrs_Tmp.Fields!HoraIngreso = lcHoraIngreso
                                mrs_Tmp.Fields!Paciente = lcPaciente
                                mrs_Tmp.Fields!Sexo = lcSexo
                                mrs_Tmp.Fields!motivo = lcMotivo
                                mrs_Tmp.Fields!Estancia = lcEstancia
                                mrs_Tmp.Fields!FechaEgreso = ldFechaEgreso
                                mrs_Tmp.Fields!horaEgreso = lcHoraEgreso
                                mrs_Tmp.Fields!destino = lcDestino
                                mrs_Tmp.Fields!Recibe = lcRecibe
                                mrs_Tmp.Fields!MedicoIE = lcMedicoIE
                                mrs_Tmp.Fields!ConsumoImporte = lnImporte
                                mrs_Tmp.Fields!ConsumoDescripcion = lcDescripcion
                                mrs_Tmp.Update
                            Else
                                mrs_Det.AddNew
                                mrs_Det.Fields!nroCuenta = lnNroCuenta
                                mrs_Det.Fields!ConsumoDescripcion = lcDescripcion
                                mrs_Det.Fields!ConsumoImporte = lnImporte
                                mrs_Det.Update
                            End If
                            lnTotalG = lnTotalG + lnImporte
                            mrs_Tmp2.MoveNext
                       Loop
                    End If
                    'farmacia
                    Set mrs_Tmp2 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorCuentaTodos(lnNroCuenta)
                    mrs_Tmp2.Filter = "idEstadoFacturacion=1"
                    If mrs_Tmp2.RecordCount > 0 Then
                       mrs_Tmp2.MoveFirst
                       lnImporte = lnImporte - mrs_Tmp2.Fields!ImporteExonerado
                       Do While Not mrs_Tmp2.EOF
                            lnImporte = mrs_Tmp2.Fields!TotalFinanciado
                            lcDescripcion = ""
                            Set mrs_Tmp3 = mo_reglasComunes.CatalogoBienesInsumosSeleccionarPorFiltro(" idProducto=" & mrs_Tmp2.Fields!idProducto, sghPorCodigo)
                            If mrs_Tmp3.RecordCount > 0 Then
                               lcDescripcion = Trim(mrs_Tmp3.Fields!Codigo) & " " & mrs_Tmp3.Fields!Nombre
                            End If
                            mrs_Tmp3.Close
                            If lbEnExcel = True Then
                                mrs_Tmp.AddNew
                                mrs_Tmp.Fields!nroCuenta = lnNroCuenta
                                mrs_Tmp.Fields!FechaIngreso = ldFechaIngreso
                                mrs_Tmp.Fields!HoraIngreso = lcHoraIngreso
                                mrs_Tmp.Fields!Paciente = lcPaciente
                                mrs_Tmp.Fields!Sexo = lcSexo
                                mrs_Tmp.Fields!motivo = lcMotivo
                                mrs_Tmp.Fields!Estancia = lcEstancia
                                mrs_Tmp.Fields!FechaEgreso = ldFechaEgreso
                                mrs_Tmp.Fields!horaEgreso = lcHoraEgreso
                                mrs_Tmp.Fields!destino = lcDestino
                                mrs_Tmp.Fields!Recibe = lcRecibe
                                mrs_Tmp.Fields!MedicoIE = lcMedicoIE
                                mrs_Tmp.Fields!ConsumoImporte = lnImporte
                                mrs_Tmp.Fields!ConsumoDescripcion = lcDescripcion
                                mrs_Tmp.Update
                            Else
                                mrs_Det.AddNew
                                mrs_Det.Fields!nroCuenta = lnNroCuenta
                                mrs_Det.Fields!ConsumoDescripcion = lcDescripcion
                                mrs_Det.Fields!ConsumoImporte = lnImporte
                                mrs_Det.Update
                            End If
                            lnTotalG = lnTotalG + lnImporte
                            mrs_Tmp2.MoveNext
                       Loop
                    End If
                Else
                    'Paciente no ha salido de Alta Administrativa
                    'Calcula estancia hasta la Fecha de Reporte
                    lnIdServicio = Val(lcBuscaParametro.SeleccionaFilaParametro(202))
                    If mrs_Tmp1.Fields!FechaIngreso = Date Then
                       lcEstancia = Trim(Str(HorasDelPacienteEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, Date, Format(Now, "hh:mm")))) & "  Hrs."
                       lnEstancia = HorasDelPacienteEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, Date, Format(Now, "hh:mm"))
                    Else
                       lcEstancia = Trim(Str(DiasDeEstanciaEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, Date, Format(Now, "hh:mm")))) & "  Días"
                       lnEstancia = DiasDeEstanciaEnHospitalizacionEmergencia(mrs_Tmp1.Fields!FechaIngreso, mrs_Tmp1.Fields!HoraIngreso, Date, Format(Now, "hh:mm"))
                    End If
                    Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(lnIdServicio, 1)
                    If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
                        lnImporte = Round(oDoCatalogoServicioHosp.PrecioUnitario * lnEstancia, 2)
                        lcDescripcion = "Estancia hasta la FECHA"
                    End If
                    lnTotalG = lnTotalG + lnImporte
                    '
                    If lbEnExcel = True Then
                        mrs_Tmp.AddNew
                        mrs_Tmp.Fields!nroCuenta = lnNroCuenta
                        mrs_Tmp.Fields!FechaIngreso = ldFechaIngreso
                        mrs_Tmp.Fields!HoraIngreso = lcHoraIngreso
                        mrs_Tmp.Fields!Paciente = lcPaciente
                        mrs_Tmp.Fields!Sexo = lcSexo
                        mrs_Tmp.Fields!motivo = lcMotivo
                        mrs_Tmp.Fields!Estancia = lcEstancia
                        'mrs_Tmp.Fields!FechaEgreso = ldFechaEgreso
                        mrs_Tmp.Fields!horaEgreso = lcHoraEgreso
                        mrs_Tmp.Fields!destino = lcDestino
                        mrs_Tmp.Fields!Recibe = lcRecibe
                        mrs_Tmp.Fields!MedicoIE = lcMedicoIE
                        mrs_Tmp.Fields!ConsumoImporte = lnImporte
                        mrs_Tmp.Fields!ConsumoDescripcion = lcDescripcion
                        mrs_Tmp.Update
                    Else
                        mrs_Cab.AddNew
                        mrs_Cab.Fields!nroCuenta = lnNroCuenta
                        mrs_Cab.Fields!FechaIngreso = ldFechaIngreso
                        mrs_Cab.Fields!HoraIngreso = lcHoraIngreso
                        mrs_Cab.Fields!Paciente = lcPaciente
                        mrs_Cab.Fields!Sexo = lcSexo
                        mrs_Cab.Fields!motivo = lcMotivo
                        mrs_Cab.Fields!Estancia = lcEstancia
                        'mrs_Cab.Fields!FechaEgreso = ldFechaEgreso
                        mrs_Cab.Fields!horaEgreso = lcHoraEgreso
                        mrs_Cab.Fields!destino = lcRecibe
                        mrs_Cab.Fields!Recibe = lcRecibe
                        mrs_Cab.Fields!MedicoIE = lcMedicoIE
                        mrs_Cab.Update
                        '
                        mrs_Det.AddNew
                        mrs_Det.Fields!nroCuenta = lnNroCuenta
                        mrs_Det.Fields!ConsumoDescripcion = lcDescripcion
                        mrs_Det.Fields!ConsumoImporte = lnImporte
                        mrs_Det.Update
                    End If
                End If
                mrs_Tmp1.MoveNext
        Loop
        lcPie = "N° Pacientes: " & Trim(Str(mrs_Tmp1.RecordCount))
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
            mrs_Cab.Close
            mrs_Det.Close
            lcTexto1 = " SHAPE {select NroCuenta,FechaIngreso,HoraIngreso,Paciente,sexo," & _
                                       "motivo, estancia, fechaEgreso,horaEgreso,destino," & _
                                       "recibe,medicoIE from reporte_cabecera} AS padre" & _
                       " APPEND ({select NroCuenta,ConsumoDescripcion,ConsumoImporte FROM reporte_detalle} AS hijo" & _
                       " RELATE NroCuenta TO NroCuenta) AS hijo"
            mrs_Shape.Open lcTexto1, sighEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
            With RpHAlojados
                 .Orientation = rptOrientLandscape
                 .Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                 .Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                 Set .Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                 .Sections("pie").Controls("lblPie").Caption = lcPie
                 .Sections("pie").Controls("lblTotalG").Caption = lnTotalG
                 Set .DataSource = mrs_Shape
                 .DataMember = ""
                 With .Sections("CabGrupo")
                      .Controls("txtFechaIngreso").DataMember = ""
                      .Controls("txtFechaIngreso").DataField = "FechaIngreso"
                      .Controls("txtHoraIngreso").DataMember = ""
                      .Controls("txtHoraIngreso").DataField = "HoraIngreso"
                      .Controls("txtFechaIngreso").DataMember = ""
                      .Controls("txtPaciente").DataField = "Paciente"
                      .Controls("txtPaciente").DataMember = ""
                      .Controls("txtSexo").DataField = "sexo"
                      .Controls("txtSexo").DataMember = ""
                      .Controls("txtMotivo").DataField = "motivo"
                      .Controls("txtMotivo").DataMember = ""
                      .Controls("txtEstancia").DataField = "estancia"
                      .Controls("txtEstancia").DataMember = ""
                      .Controls("txtFechaEgreso").DataField = "fechaEgreso"
                      .Controls("txtFechaEgreso").DataMember = ""
                      .Controls("txtHoraEgreso").DataField = "horaEgreso"
                      .Controls("txtHoraEgreso").DataMember = ""
                      .Controls("txtDestino").DataField = "destino"
                      .Controls("txtDestino").DataMember = ""
                      .Controls("txtFamiliar").DataField = "recibe"
                      .Controls("txtFamiliar").DataMember = ""
                      .Controls("txtMedicoIS").DataField = "medicoIE"
                      .Controls("txtMedicoIS").DataMember = ""
                 End With
                 With .Sections("DetGrupo")
                      .Controls("txtConsumoDescripcion").DataMember = "Hijo"
                      .Controls("txtConsumoDescripcion").DataField = "ConsumoDescripcion"
                      .Controls("txtConsumoImporte").DataMember = "Hijo"
                      .Controls("txtConsumoImporte").DataField = "ConsumoImporte"
                 End With
                 .Show 1
            End With
        End If
    End If
End Sub




