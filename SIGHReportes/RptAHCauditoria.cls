VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHCauditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para auditoría en Archivo Clínico
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Consideraciones:
Option Explicit
Dim ml_idUsuario  As Long
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New AHCauditoria
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, _
                   ml_idUsuario As Long, mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim ldFecha As Date, ldFechaSolo As Date
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set mrs_Tmp1 = mo_ReglasFarmacia.AuditoriaPorFechasUsuario(ml_idUsuario, mda_FechaInicio, mda_FechaFin)
    mrs_Tmp1.Filter = "idListItem=501 or idListItem=502 or idListItem=1330"
    If mrs_Tmp1.RecordCount = 0 Then
         MsgBox "No existe Datos", vbInformation, "Reporte"
    Else
         With mrs_Tmp
            .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
            .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
            .Fields.Append "Accion", adVarChar, 10, adFldIsNullable
            .Fields.Append "Nusuario", adVarChar, 160, adFldIsNullable
            .Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
            .Fields.Append "NombrePC", adVarChar, 30, adFldIsNullable
            .Fields.Append "Observacion1", adVarChar, 250, adFldIsNullable
            .Fields.Append "Observacion2", adVarChar, 250, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
         End With
         '
         mrs_Tmp1.MoveFirst
         Do While Not mrs_Tmp1.EOF
           ldFecha = mrs_Tmp1.Fields!fechaHora
           ldFechaSolo = Format(mrs_Tmp1.Fields!fechaHora, sighentidades.DevuelveFechaSoloFormato_DMY)
           lcHoraSolo = Format(mrs_Tmp1.Fields!fechaHora, "hh:mm")
           lnIdListItem = mrs_Tmp1.Fields!idListItem
           lcNusuario = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & mrs_Tmp1.Fields!Nombres
           lcUsuario = mrs_Tmp1.Fields!Usuario
           lcAccion = IIf(mrs_Tmp1.Fields!Accion = "A", "Agregó", IIf(mrs_Tmp1.Fields!Accion = "M", "Modificó", "Anuló"))
           lcPc = IIf(IsNull(mrs_Tmp1.Fields!NombrePC), "", mrs_Tmp1.Fields!NombrePC)
           Select Case mrs_Tmp1.Fields!idListItem
           Case 502      'Movimientos
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.MovimientosHistoriaClinicaSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)), False) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   If mrs_Tmp2.Fields!idMotivo = 9 Then
                      lcTexto2 = "(Llegada al A.C.)  desde: " & mrs_Tmp2.Fields!sOrigen
                   Else
                      lcTexto2 = "(Salida del A.C.)  hacia: " & mrs_Tmp2.Fields!sDestino
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 1332      'Movimientos Formatos
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.MovimientosFormatosSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)), False) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   If mrs_Tmp2.Fields!idMotivo = 9 Then
                      lcTexto2 = "(Llegada Formato al A.C.)  desde: " & mrs_Tmp2.Fields!sOrigen
                   Else
                      lcTexto2 = "(Salida Formato del A.C.)  hacia: " & mrs_Tmp2.Fields!sDestino
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case 501      'Archivo Clinico
                Set mrs_Tmp2 = mo_ReglasArchivoClinico.HistoriaClinicasSeleccionarXnrohistoria(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mrs_Tmp2.Fields!NroHistoriaClinica)), False) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "(Mantenimiento de H.C.)    "
                   If Not IsNull(mrs_Tmp2.Fields!nroHistoriaClinicaAnterior) Then
                      lcTexto2 = lcTexto2 & "    " & "Nro anterior: " & Trim(Str(mrs_Tmp2.Fields!nroHistoriaClinicaAnterior))
                   Else
                      lcTexto2 = lcTexto2 & mrs_Tmp1.Fields!observaciones
                   End If
                ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                   lcTexto2 = mrs_Tmp1.Fields!observaciones
                End If
                mrs_Tmp2.Close
                Do While Not mrs_Tmp1.EOF And ldFecha = mrs_Tmp1.Fields!fechaHora
                   mrs_Tmp1.MoveNext
                   If mrs_Tmp1.EOF Then
                      Exit Do
                   End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields!fechaCreacion = ldFechaSolo
                mrs_Tmp.Fields!HoraCreacion = lcHoraSolo
                mrs_Tmp.Fields!Accion = lcAccion
                mrs_Tmp.Fields!Nusuario = lcNusuario
                mrs_Tmp.Fields!Usuario = lcUsuario
                mrs_Tmp.Fields!NombrePC = lcPc
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 250)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 250)
                mrs_Tmp.Update
           Case Else
               mrs_Tmp1.MoveNext
           End Select
        Loop
        lcPie = "Nro Movimientos: " & Trim(Str(mrs_Tmp.RecordCount))
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
        Else
            Set RpAHCauditoria.DataSource = mrs_Tmp
            RpAHCauditoria.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
            RpAHCauditoria.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
            RpAHCauditoria.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
            RpAHCauditoria.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
            RpAHCauditoria.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
                  
            RpAHCauditoria.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
            RpAHCauditoria.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
            Set RpAHCauditoria.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
            RpAHCauditoria.Sections("pie").Controls("lblPie").Caption = lcPie
            RpAHCauditoria.Orientation = rptOrientLandscape
            RpAHCauditoria.Show 1
        End If
    End If
    Set oConexion = Nothing
End Sub


