VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHhcNoLlegaAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historias que no llegan al Archivo
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Sub EjecutaFormulario()
    Dim oFormulario As New AHCNoLleganAC
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, mda_FechaInicio As String, _
                         mda_FechaFin As String, lnHwnd As Long, mda_fechaCitaMaxima As String)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim lcSql As String
    Dim lbVolvio As Boolean
    Dim lnNroHistoria As Long, lnIdAtencion As Long
    Dim lnIdArchivo As Long
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_AdminComun As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim lcPie As String
    Dim lcPaciente As String
    Dim lcServicioSalida As String
    Dim ldFechaSalida As Date
    With mrs_Tmp
        .Fields.Append "UltimoDigito", adInteger
        .Fields.Append "NroHistoria", adVarChar, 20, adFldIsNullable
        .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
        .Fields.Append "ServicioSalida", adVarChar, 150, adFldIsNullable
        .Fields.Append "FechaSalida", adDate, 10, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
    Set mrs_Tmp1 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaNoLlegaAlArchivoClinicodebb(mda_FechaInicio, mda_FechaFin)
    If mrs_Tmp1.RecordCount > 0 Then
       lnIdArchivo = mo_AdminComun.ParametrosIdServicioArchivoClinico()
       mrs_Tmp1.MoveFirst
       Do While Not mrs_Tmp1.EOF
          lnNroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
          lcPaciente = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & mrs_Tmp1.Fields!PrimerNombre
          lcServicioSalida = mrs_Tmp1.Fields!servicioSalida
          ldFechaSalida = mrs_Tmp1.Fields!FechaMovimiento
          lbVolvio = True
          Do While Not mrs_Tmp1.EOF And lnNroHistoria = mrs_Tmp1.Fields!NroHistoriaClinica
             If lnIdArchivo = mrs_Tmp1.Fields!idServicioDestino Then
                lbVolvio = True
             Else
                lbVolvio = False
             End If
             lcServicioSalida = mrs_Tmp1.Fields!servicioSalida
             lnIdAtencion = IIf(IsNull(mrs_Tmp1!idAtencion), 0, mrs_Tmp1!idAtencion)
             mrs_Tmp1.MoveNext
             If mrs_Tmp1.EOF Then
                Exit Do
             End If
          Loop
          If lbVolvio = False And lnIdAtencion > 0 Then
             Set mrs_Tmp2 = mo_ReglasFacturacion.AtencionesSeleccionarPorIdAtencion(lnIdAtencion)
             If mrs_Tmp2.RecordCount > 0 Then
                If CDate(Format(mrs_Tmp2!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_DMY) + " " + mrs_Tmp2!HoraIngreso + ":00") > CDate(mda_fechaCitaMaxima) Then
                   lbVolvio = True
                End If
             End If
             mrs_Tmp2.Close
          End If
          If lbVolvio = False Then
             mrs_Tmp.AddNew
             mrs_Tmp.Fields!ultimoDigito = Right(HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnNroHistoria)), False), 1)
             mrs_Tmp.Fields!nroHistoria = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(lnNroHistoria)), False)
             mrs_Tmp.Fields!Paciente = lcPaciente
             mrs_Tmp.Fields!servicioSalida = lcServicioSalida
             mrs_Tmp.Fields!FechaSalida = CDate(Format(ldFechaSalida, "dd/mm/yyyy"))
             mrs_Tmp.Update
          End If
       Loop
    End If
    mrs_Tmp1.Close
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No exite información con esos datos", vbInformation, "Reporte"
    Else
        lcPie = "Nro Historias: " & Trim(Str(mrs_Tmp.RecordCount))
        'For lnIdArchivo = 0 To 9
        '   mrs_Tmp.AddNew
        '   mrs_Tmp.Fields!ultimoDigito = lnIdArchivo
        '   mrs_Tmp.Update
        'Next
        mrs_Tmp.Sort = "ultimoDigito,Paciente"
        If lbEnExcel = True Then
            Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
            mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
            Set RpACHCnoLleganAC.DataSource = mrs_Tmp
            RpACHCnoLleganAC.RightMargin = 10
            RpACHCnoLleganAC.TopMargin = 10
            RpACHCnoLleganAC.LeftMargin = 10
            RpACHCnoLleganAC.BottomMargin = 10
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
            RpACHCnoLleganAC.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
            Set RpACHCnoLleganAC.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            RpACHCnoLleganAC.Sections("piePag").Controls("lblTotal").Caption = "N° Historias: " & Trim(Str(mrs_Tmp.RecordCount))
            
            RpACHCnoLleganAC.Orientation = rptOrientPortrait
            RpACHCnoLleganAC.Show 1
            'debb-27/05/2015
            Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
            mo_reglasComunes.grabaTablaAuditoria ("RpACHCnoLleganAC: " & _
                                           Mid(lcSubTitulo, IIf(InStr(lcSubTitulo, "FILTROS: ") > 0, 10, 1)))
            Set mo_reglasComunes = Nothing
            '
        End If
    End If
    Set mo_ReglasArchivoClinico = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_AdminComun = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mrs_Tmp1 = Nothing
    Set mrs_Tmp2 = Nothing
    Set mrs_Tmp = Nothing
End Sub







