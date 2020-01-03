VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historia Emergencia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasServGeograf As New SIGHNegocios.ReglasServGeograf
Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica

Sub CrearReporteReferencias(lcEESSdestino As String, _
                               mo_paciente As DOPaciente, mo_atenciones As DOAtencion, _
                               mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, lbReporteEsEnPDF As Boolean)
    
    Dim oRsContraReferencias As New Recordset
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim oRsTmp1 As New Recordset
    Dim rsDx As New Recordset
    Dim oDOServicio As New DOServicio
    Dim oConexion As New Connection
    Dim lcX As String, lcFechaReferencia As String, lcDistrito As String, lcDpto As String, lcUPSfua As String, lcIdDistrito As String
    Dim oEdad As Edad
    Dim lcMedicoColegiatura As String, lcMedico As String
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    If mo_atenciones.idTipoServicio <> 1 Then
       Set mo_paciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_atenciones.IdPaciente, oConexion)
    End If
    
    
    lcMedico = "": lcMedicoColegiatura = ""
    Set rsDx = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.idMedicoIngreso, mo_atenciones.IdMedicoEgreso))
    If rsDx.RecordCount > 0 Then
       lcMedicoColegiatura = IIf(IsNull(rsDx!colegiatura), "", rsDx!colegiatura)
       lcMedico = Trim(rsDx!ApellidoPaterno) & " " & Trim(rsDx!ApellidoMaterno) & " " & rsDx!Nombres
    End If
    rsDx.Close
    
    Set rsDx = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mo_atenciones.idAtencion)
    
    lcFechaReferencia = Format(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.FechaIngreso, mo_atenciones.fechaEgreso), SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oEdad = SIGHEntidades.CalcularEdad(mo_paciente.FechaNacimiento, CDate(lcFechaReferencia))
    lcIdDistrito = Right("0" & Trim(Str(mo_paciente.IdDistritoDomicilio)), 6)
    lcDistrito = ""
    Set oRsTmp1 = mo_ReglasServGeograf.DistritoSeleccionarPorProvincia(Val(Left(lcIdDistrito, 4)))
    oRsTmp1.Filter = "idDistrito=" & mo_paciente.IdDistritoDomicilio
    If oRsTmp1.RecordCount > 0 Then
       lcDistrito = oRsTmp1!Nombre
    End If
    oRsTmp1.Close
    lcDpto = ""
    Set oRsTmp1 = mo_ReglasServGeograf.DepartamentosSeleccionarXidentificador(Val(Left(lcIdDistrito, 2)))
    If oRsTmp1.RecordCount > 0 Then
       lcDpto = oRsTmp1!Nombre
    End If
    oRsTmp1.Close
    lcUPSfua = ""
    Set oDOServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.IdServicioIngreso, mo_atenciones.IdServicioEgreso), oConexion)
    If Not oDOServicio Is Nothing Then
       lcUPSfua = oDOServicio.codigoServicioFUA
    End If
    lcX = " "
    
    
    Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
    If lbReporteEsEnPDF = True Then
       lcArchivoPDF = SIGHEntidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(237)) & _
       Trim(Str(mo_paciente.IdPaciente)) & "-" & Format(mo_atenciones.FechaIngreso, "DDMMYYYY") & _
       "-CTA" & Trim(Str(mo_atenciones.idCuentaAtencion)) & "-IMPRESIONREFERENCIA.pdf"
       If SePuedeImprimirPDF(lcArchivoPDF, True) = True Then
          lbSePuedeImprimirPDF = True
       End If
    End If
    
    Set oRsContraReferencias = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(1)
    Set FormReferencia.DataSource = oRsContraReferencias
    If lbSePuedeImprimirPDF = True And lcBuscaParametro.SeleccionaFilaParametro(582) = "S" Then
       FormReferencia.Sections("cabecera").Controls("Image2").Visible = True
    Else
       FormReferencia.Sections("cabecera").Controls("Image2").Visible = False
    End If
    FormReferencia.Sections("cabecera").Controls("etReferencia").Caption = mo_DoAtencionDatosAdicionales.NroReferenciaDestino
    FormReferencia.Sections("cabecera").Controls("etFechaRef").Caption = Left(lcFechaReferencia, 2) & Space(8) & Mid(lcFechaReferencia, 4, 2) & Space(6) & Right(lcFechaReferencia, 4)
    FormReferencia.Sections("cabecera").Controls("etEESSorigen").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    FormReferencia.Sections("cabecera").Controls("etEESSdestino").Caption = lcEESSdestino
    FormReferencia.Sections("cabecera").Controls("etAfiliacionSIS").Caption = mo_DoAtencionDatosAdicionales.sisAfiliacion
    FormReferencia.Sections("cabecera").Controls("etHistoria").Caption = mo_paciente.NroHistoriaClinica
    FormReferencia.Sections("cabecera").Controls("etApellidoPaterno").Caption = mo_paciente.ApellidoPaterno
    FormReferencia.Sections("cabecera").Controls("etApellidoMaterno").Caption = mo_paciente.ApellidoMaterno
    FormReferencia.Sections("cabecera").Controls("etPrimerNombre").Caption = mo_paciente.PrimerNombre
    FormReferencia.Sections("cabecera").Controls("etSegundoNombre").Caption = mo_paciente.SegundoNombre
    FormReferencia.Sections("cabecera").Controls("etSexo1").Caption = IIf(mo_paciente.idTipoSexo = 1, "X", "")
    FormReferencia.Sections("cabecera").Controls("etSexo2").Caption = IIf(mo_paciente.idTipoSexo = 2, "X", "")
    FormReferencia.Sections("cabecera").Controls("etEdadAnios").Caption = IIf(oEdad.TipoEdad = 1, oEdad.Edad, "")
    FormReferencia.Sections("cabecera").Controls("etEdadMeses").Caption = IIf(oEdad.TipoEdad = 2, oEdad.Edad, "")
    FormReferencia.Sections("cabecera").Controls("etEdadDias").Caption = IIf(oEdad.TipoEdad = 3, oEdad.Edad, "")
    FormReferencia.Sections("cabecera").Controls("etDireccion").Caption = mo_DoAtencionDatosAdicionales.DireccionDomicilio
    FormReferencia.Sections("cabecera").Controls("etDistrito").Caption = lcDistrito
    FormReferencia.Sections("cabecera").Controls("etDepartamento").Caption = lcDpto
'    FormReferencia.Sections("cabecera").Controls("etFechaIngreso").Caption = mo_atenciones.FechaIngreso
'    FormReferencia.Sections("cabecera").Controls("etFechaEgreso").Caption = IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.FechaIngreso, mo_atenciones.fechaEgreso)
    If rsDx.RecordCount > 0 Then
        If mo_atenciones.idTipoServicio <> 1 Then
'            rsDx.Filter = "idClasificacionDx=2"
'            If rsDx.RecordCount > 0 Then
'            FormReferencia.Sections("cabecera").Controls("etDxIng1").Caption = rsDx!dDiagnostico
'            rsDx.MoveNext
'            If Not rsDx.EOF Then
'                FormReferencia.Sections("cabecera").Controls("etDxIng2").Caption = rsDx!dDiagnostico
'                rsDx.MoveNext
'                If Not rsDx.EOF Then
'                    FormReferencia.Sections("cabecera").Controls("etDxIng3").Caption = rsDx!dDiagnostico
'                End If
'            End If
'            End If
'
             rsDx.Filter = "idClasificacionDx=3"
        Else
'            FormReferencia.Sections("cabecera").Controls("etDxIng1").Caption = lcX
'            FormReferencia.Sections("cabecera").Controls("etDxIng2").Caption = lcX
'            FormReferencia.Sections("cabecera").Controls("etDxIng3").Caption = lcX
        End If
        If Not rsDx.EOF Then
        FormReferencia.Sections("cabecera").Controls("etDxEgr1").Caption = rsDx!dDiagnostico
        FormReferencia.Sections("cabecera").Controls("etDxE1").Caption = rsDx!CodigoCIE10
        FormReferencia.Sections("cabecera").Controls("etDxD1").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
        If mo_atenciones.idTipoServicio = 1 Then
           FormReferencia.Sections("cabecera").Controls("etDxP1").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
        Else
           FormReferencia.Sections("cabecera").Controls("etDxP1").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
        End If
        FormReferencia.Sections("cabecera").Controls("etDxR1").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
        rsDx.MoveNext
        If Not rsDx.EOF Then
            FormReferencia.Sections("cabecera").Controls("etDxEgr2").Caption = rsDx!dDiagnostico
            FormReferencia.Sections("cabecera").Controls("etDxE2").Caption = rsDx!CodigoCIE10
            FormReferencia.Sections("cabecera").Controls("etDxD2").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
            If mo_atenciones.idTipoServicio = 1 Then
               FormReferencia.Sections("cabecera").Controls("etDxP2").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
            Else
               FormReferencia.Sections("cabecera").Controls("etDxP2").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
            End If
            FormReferencia.Sections("cabecera").Controls("etDxR2").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
            rsDx.MoveNext
            If Not rsDx.EOF Then
                FormReferencia.Sections("cabecera").Controls("etDxEgr3").Caption = rsDx!dDiagnostico
                FormReferencia.Sections("cabecera").Controls("etDxE3").Caption = rsDx!CodigoCIE10
                FormReferencia.Sections("cabecera").Controls("etDxD3").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
                If mo_atenciones.idTipoServicio = 1 Then
                   FormReferencia.Sections("cabecera").Controls("etDxP3").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
                Else
                   FormReferencia.Sections("cabecera").Controls("etDxP3").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
                End If
                FormReferencia.Sections("cabecera").Controls("etDxR3").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
            End If
        End If
        End If
    End If
'    FormReferencia.Sections("cabecera").Controls("etCE").Caption = IIf(mo_atenciones.idTipoServicio = 1, "X", "")
'    FormReferencia.Sections("cabecera").Controls("etEmergencia").Caption = IIf(mo_atenciones.idTipoServicio = 2, "X", "")
'    FormReferencia.Sections("cabecera").Controls("etApoyoDx").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etHospitalizacion").Caption = IIf(mo_atenciones.idTipoServicio = 3, "X", "")
'    FormReferencia.Sections("cabecera").Controls("etPediatria").Caption = IIf(lcUPSfua = "224700", "X", "")
'    FormReferencia.Sections("cabecera").Controls("etMedicina").Caption = IIf(lcUPSfua = "302303", "X", "")
'    FormReferencia.Sections("cabecera").Controls("etCirugia").Caption = IIf(lcUPSfua = "220100", "X", "")
'    FormReferencia.Sections("cabecera").Controls("etGinecoObs").Caption = IIf(lcUPSfua = "301607" Or lcUPSfua = "301608", "X", "")
'    FormReferencia.Sections("cabecera").Controls("etLaboratorio").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etImagenes").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etOtrosUPS").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etCurado").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etMejorado").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etAtendApDx").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etDesercion").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etRetiroVol").Caption = lcX
'    FormReferencia.Sections("cabecera").Controls("etFallecido").Caption = lcX
    'FormReferencia.Sections("cabecera").Controls("etMedico").Caption = lcMedico
    FormReferencia.Sections("cabecera").Controls("etColegiatura").Caption = lcX 'lcMedicoColegiatura
    FormReferencia.RightMargin = 0
    FormReferencia.TopMargin = 0
    FormReferencia.LeftMargin = 0
    FormReferencia.BottomMargin = 0
    FormReferencia.Orientation = rptOrientPortrait
    
    
    If lbSePuedeImprimirPDF = True Then
       FormReferencia.PrintReport False
       

    Else
       FormReferencia.Show 1
    End If
    SeteaOtraImpresoraDefault SIGHEntidades.ImpresoraDefaultDeEstaPC
    
    Set oRsContraReferencias = Nothing
    Set mo_AdminReportes = Nothing
    Set oRsTmp1 = Nothing
    Set rsDx = Nothing
    Set oDOServicio = Nothing
    Set oConexion = Nothing
End Sub


                            
Sub CreaReporteContrarefencias(lcEESSdestino As String, _
                               mo_paciente As DOPaciente, mo_atenciones As DOAtencion, _
                               mo_DoAtencionDatosAdicionales As DoAtencionDatosAdicionales, lbReporteEsEnPDF As Boolean)
    Dim oRsContraReferencias As New Recordset
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim oRsTmp1 As New Recordset
    Dim rsDx As New Recordset
    Dim oDOServicio As New DOServicio
    Dim oConexion As New Connection
    Dim lcX As String, lcFechaReferencia As String, lcDistrito As String, lcDpto As String, lcUPSfua As String, lcIdDistrito As String
    Dim oEdad As Edad
    Dim lcMedicoColegiatura As String, lcMedico As String
    
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    
    If mo_atenciones.idTipoServicio <> 1 Then
       Set mo_paciente = mo_AdminAdmision.PacientesSeleccionarPorId(mo_atenciones.IdPaciente, oConexion)
    End If
    
    
    lcMedico = "": lcMedicoColegiatura = ""
    Set rsDx = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.idMedicoIngreso, mo_atenciones.IdMedicoEgreso))
    If rsDx.RecordCount > 0 Then
       lcMedicoColegiatura = IIf(IsNull(rsDx!colegiatura), "", rsDx!colegiatura)
       lcMedico = Trim(rsDx!ApellidoPaterno) & " " & Trim(rsDx!ApellidoMaterno) & " " & rsDx!Nombres
    End If
    rsDx.Close
    
    Set rsDx = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mo_atenciones.idAtencion)
    
    lcFechaReferencia = Format(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.FechaIngreso, mo_atenciones.fechaEgreso), SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    oEdad = SIGHEntidades.CalcularEdad(mo_paciente.FechaNacimiento, CDate(lcFechaReferencia))
    lcIdDistrito = Right("0" & Trim(Str(mo_paciente.IdDistritoDomicilio)), 6)
    lcDistrito = ""
    Set oRsTmp1 = mo_ReglasServGeograf.DistritoSeleccionarPorProvincia(Val(Left(lcIdDistrito, 4)))
    oRsTmp1.Filter = "idDistrito=" & mo_paciente.IdDistritoDomicilio
    If oRsTmp1.RecordCount > 0 Then
       lcDistrito = oRsTmp1!Nombre
    End If
    oRsTmp1.Close
    lcDpto = ""
    Set oRsTmp1 = mo_ReglasServGeograf.DepartamentosSeleccionarXidentificador(Val(Left(lcIdDistrito, 2)))
    If oRsTmp1.RecordCount > 0 Then
       lcDpto = oRsTmp1!Nombre
    End If
    oRsTmp1.Close
    lcUPSfua = ""
    Set oDOServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.IdServicioIngreso, mo_atenciones.IdServicioEgreso), oConexion)
    If Not oDOServicio Is Nothing Then
       lcUPSfua = oDOServicio.codigoServicioFUA
    End If
    lcX = " "
    
    Dim lbSePuedeImprimirPDF As Boolean, lcArchivoPDF As String
    If lbReporteEsEnPDF = True Then
       lcArchivoPDF = SIGHEntidades.DevuelveRutaConSlashInvertida(lcBuscaParametro.SeleccionaFilaParametro(237)) & _
       Trim(Str(mo_paciente.IdPaciente)) & "-" & Format(mo_atenciones.FechaIngreso, "DDMMYYYY") & _
       "-CTA" & Trim(Str(mo_atenciones.idCuentaAtencion)) & "-IMPRESIONCONTRAREFERENCIA.pdf"
       If SePuedeImprimirPDF(lcArchivoPDF, True) = True Then
          lbSePuedeImprimirPDF = True
       End If
    End If
    
    
    
    Set oRsContraReferencias = mo_AdminReportes.ReporteAtencionesParaHistoriaClinica(1)
    Set FormContrareferencia.DataSource = oRsContraReferencias
    If lbSePuedeImprimirPDF = True And lcBuscaParametro.SeleccionaFilaParametro(582) = "S" Then
       FormContrareferencia.Sections("cabecera").Controls("Image2").Visible = True
    Else
       FormContrareferencia.Sections("cabecera").Controls("Image2").Visible = False
    End If
    
    FormContrareferencia.Sections("cabecera").Controls("etReferencia").Caption = mo_DoAtencionDatosAdicionales.NroReferenciaDestino
    FormContrareferencia.Sections("cabecera").Controls("etFechaRef").Caption = Left(lcFechaReferencia, 2) & Space(8) & Mid(lcFechaReferencia, 4, 2) & Space(6) & Right(lcFechaReferencia, 4)
    FormContrareferencia.Sections("cabecera").Controls("etEESSorigen").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    FormContrareferencia.Sections("cabecera").Controls("etEESSdestino").Caption = lcEESSdestino
    FormContrareferencia.Sections("cabecera").Controls("etAfiliacionSIS").Caption = mo_DoAtencionDatosAdicionales.sisAfiliacion
    FormContrareferencia.Sections("cabecera").Controls("etHistoria").Caption = mo_paciente.NroHistoriaClinica
    FormContrareferencia.Sections("cabecera").Controls("etApellidoPaterno").Caption = mo_paciente.ApellidoPaterno
    FormContrareferencia.Sections("cabecera").Controls("etApellidoMaterno").Caption = mo_paciente.ApellidoMaterno
    FormContrareferencia.Sections("cabecera").Controls("etPrimerNombre").Caption = mo_paciente.PrimerNombre
    FormContrareferencia.Sections("cabecera").Controls("etSegundoNombre").Caption = mo_paciente.SegundoNombre
    FormContrareferencia.Sections("cabecera").Controls("etSexo1").Caption = IIf(mo_paciente.idTipoSexo = 1, "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etSexo2").Caption = IIf(mo_paciente.idTipoSexo = 2, "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etEdadAnios").Caption = IIf(oEdad.TipoEdad = 1, oEdad.Edad, "")
    FormContrareferencia.Sections("cabecera").Controls("etEdadMeses").Caption = IIf(oEdad.TipoEdad = 2, oEdad.Edad, "")
    FormContrareferencia.Sections("cabecera").Controls("etEdadDias").Caption = IIf(oEdad.TipoEdad = 3, oEdad.Edad, "")
    FormContrareferencia.Sections("cabecera").Controls("etDireccion").Caption = mo_DoAtencionDatosAdicionales.DireccionDomicilio
    FormContrareferencia.Sections("cabecera").Controls("etDistrito").Caption = lcDistrito
    FormContrareferencia.Sections("cabecera").Controls("etDepartamento").Caption = lcDpto
    FormContrareferencia.Sections("cabecera").Controls("etFechaIngreso").Caption = mo_atenciones.FechaIngreso
    FormContrareferencia.Sections("cabecera").Controls("etFechaEgreso").Caption = IIf(mo_atenciones.idTipoServicio = 1, mo_atenciones.FechaIngreso, mo_atenciones.fechaEgreso)
    If rsDx.RecordCount > 0 Then
        If mo_atenciones.idTipoServicio <> 1 Then
            rsDx.Filter = "idClasificacionDx=2"
            If rsDx.RecordCount > 0 Then
            FormContrareferencia.Sections("cabecera").Controls("etDxIng1").Caption = rsDx!dDiagnostico
            rsDx.MoveNext
            If Not rsDx.EOF Then
                FormContrareferencia.Sections("cabecera").Controls("etDxIng2").Caption = rsDx!dDiagnostico
                rsDx.MoveNext
                If Not rsDx.EOF Then
                    FormContrareferencia.Sections("cabecera").Controls("etDxIng3").Caption = rsDx!dDiagnostico
                End If
            End If
            End If
            rsDx.Filter = "idClasificacionDx=3"
        Else
            FormContrareferencia.Sections("cabecera").Controls("etDxIng1").Caption = lcX
            FormContrareferencia.Sections("cabecera").Controls("etDxIng2").Caption = lcX
            FormContrareferencia.Sections("cabecera").Controls("etDxIng3").Caption = lcX
        End If
        If Not rsDx.EOF Then
        FormContrareferencia.Sections("cabecera").Controls("etDxEgr1").Caption = rsDx!dDiagnostico
        FormContrareferencia.Sections("cabecera").Controls("etDxE1").Caption = rsDx!CodigoCIE10
        FormContrareferencia.Sections("cabecera").Controls("etDxD1").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
        If mo_atenciones.idTipoServicio = 1 Then
           FormContrareferencia.Sections("cabecera").Controls("etDxP1").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
        Else
           FormContrareferencia.Sections("cabecera").Controls("etDxP1").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
        End If
        FormContrareferencia.Sections("cabecera").Controls("etDxR1").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
        rsDx.MoveNext
        If Not rsDx.EOF Then
            FormContrareferencia.Sections("cabecera").Controls("etDxEgr2").Caption = rsDx!dDiagnostico
            FormContrareferencia.Sections("cabecera").Controls("etDxE2").Caption = rsDx!CodigoCIE10
            FormContrareferencia.Sections("cabecera").Controls("etDxD2").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
            If mo_atenciones.idTipoServicio = 1 Then
               FormContrareferencia.Sections("cabecera").Controls("etDxP2").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
            Else
               FormContrareferencia.Sections("cabecera").Controls("etDxP2").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
            End If
            FormContrareferencia.Sections("cabecera").Controls("etDxR2").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
            rsDx.MoveNext
            If Not rsDx.EOF Then
                FormContrareferencia.Sections("cabecera").Controls("etDxEgr3").Caption = rsDx!dDiagnostico
                FormContrareferencia.Sections("cabecera").Controls("etDxE3").Caption = rsDx!CodigoCIE10
                FormContrareferencia.Sections("cabecera").Controls("etDxD3").Caption = IIf(rsDx!idSubClasificacionDx = 102 Or rsDx!idSubClasificacionDx = "301", "X", "")
                If mo_atenciones.idTipoServicio = 1 Then
                   FormContrareferencia.Sections("cabecera").Controls("etDxP3").Caption = IIf(rsDx!idSubClasificacionDx = 101, "X", "")
                Else
                   FormContrareferencia.Sections("cabecera").Controls("etDxP3").Caption = IIf(rsDx!idSubClasificacionDx <> "301", "X", "")
                End If
                FormContrareferencia.Sections("cabecera").Controls("etDxR3").Caption = IIf(rsDx!idSubClasificacionDx = 103, "X", "")
            End If
        End If
        End If
    End If
    FormContrareferencia.Sections("cabecera").Controls("etCE").Caption = IIf(mo_atenciones.idTipoServicio = 1, "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etEmergencia").Caption = IIf(mo_atenciones.idTipoServicio = 2, "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etApoyoDx").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etHospitalizacion").Caption = IIf(mo_atenciones.idTipoServicio = 3, "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etPediatria").Caption = IIf(lcUPSfua = "224700", "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etMedicina").Caption = IIf(lcUPSfua = "302303", "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etCirugia").Caption = IIf(lcUPSfua = "220100", "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etGinecoObs").Caption = IIf(lcUPSfua = "301607" Or lcUPSfua = "301608", "X", "")
    FormContrareferencia.Sections("cabecera").Controls("etLaboratorio").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etImagenes").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etOtrosUPS").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etCurado").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etMejorado").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etAtendApDx").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etDesercion").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etRetiroVol").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etFallecido").Caption = lcX
    FormContrareferencia.Sections("cabecera").Controls("etMedico").Caption = lcX  'lcMedico
    FormContrareferencia.Sections("cabecera").Controls("etColegiatura").Caption = lcX   'lcMedicoColegiatura
    FormContrareferencia.RightMargin = 0
    FormContrareferencia.TopMargin = 0
    FormContrareferencia.LeftMargin = 0
    FormContrareferencia.BottomMargin = 0
    FormContrareferencia.Orientation = rptOrientPortrait
    
    
    
    If lbSePuedeImprimirPDF = True Then
       FormContrareferencia.PrintReport False
       
      ' MsgBox "Se creó archivo : " & lcArchivoPDF
    Else
       FormContrareferencia.Show 1
    End If
    SeteaOtraImpresoraDefault SIGHEntidades.ImpresoraDefaultDeEstaPC
    
    Set oRsContraReferencias = Nothing
    Set mo_AdminReportes = Nothing
    Set oRsTmp1 = Nothing
    Set rsDx = Nothing
    Set oDOServicio = Nothing
    Set oConexion = Nothing
End Sub


