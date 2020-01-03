VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEmergAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Auditoría en Emergencia
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario  As Long
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New EmergAuditoria
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, ml_idUsuario As Long, mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim ldFecha As Date, ldFechaSolo As Date
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    Set mrs_Tmp1 = mo_ReglasFarmacia.AuditoriaPorFechasUsuario(ml_idUsuario, mda_FechaInicio, mda_FechaFin)
    mrs_Tmp1.Filter = "idListItem=201 or idListItem=202 or idListItem=601"
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
            .Fields.Append "Observacion1", adVarChar, 100, adFldIsNullable
            .Fields.Append "Observacion2", adVarChar, 100, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
         End With
         '
         mrs_Tmp1.MoveFirst
         Do While Not mrs_Tmp1.EOF
           ldFecha = mrs_Tmp1.Fields!fechaHora
           ldFechaSolo = Format(mrs_Tmp1.Fields!fechaHora, sighEntidades.DevuelveFechaSoloFormato_DMY)
           lcHoraSolo = Format(mrs_Tmp1.Fields!fechaHora, "hh:mm")
           lnIdListItem = mrs_Tmp1.Fields!idListItem
           lcNusuario = Trim(mrs_Tmp1.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp1.Fields!ApellidoMaterno) & " " & mrs_Tmp1.Fields!Nombres
           lcUsuario = mrs_Tmp1.Fields!Usuario
           lcAccion = IIf(mrs_Tmp1.Fields!Accion = "A", "Agregó", IIf(mrs_Tmp1.Fields!Accion = "M", "Modificó", "Anuló"))
           lcPc = IIf(IsNull(mrs_Tmp1.Fields!NombrePC), "", mrs_Tmp1.Fields!NombrePC)
           Select Case mrs_Tmp1.Fields!idListItem
           Case 201      'Pacientes
                Set mrs_Tmp2 = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!nrohistoriaclinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                   lcTexto2 = "(Mantenimiento de Pacientes desde Emergencia)"
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
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 100)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 100)
                mrs_Tmp.Update
           Case 202      'Admision Emerg
                Set mrs_Tmp2 = mo_ReglasAdmision.AtencionesSeleccionarPorIdAtencionTipoServicio(mrs_Tmp1.Fields!idRegistro, 2, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                   lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!nrohistoriaclinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre) & " (" & Trim(mrs_Tmp1.Fields!observaciones) & ")"
                   lcTexto2 = "Adm: (Fecha: " & mrs_Tmp2.Fields!FechaIngreso & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
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
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 100)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 100)
                mrs_Tmp.Update
           Case 601     'General - Reg.Procedimientos
                Set mrs_Tmp2 = mo_ReglasAdmision.FactOrdenServicioSeleccionarPorIdOrden(mrs_Tmp1.Fields!idRegistro, oConexion)
                lcTexto1 = "Tabla:  " & mrs_Tmp1.Fields!Tabla
                lcTexto2 = "IdTabla: " & mrs_Tmp1.Fields!idRegistro
                If mrs_Tmp2.RecordCount > 0 Then
                    If mrs_Tmp2.RecordCount > 0 And mrs_Tmp2.Fields!idTipoServicio <> 3 Then
                       lcTexto1 = "(" & Trim(Str(mrs_Tmp2.Fields!nrohistoriaclinica)) & ") " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre) & " " & Trim(mrs_Tmp2.Fields!SegundoNombre)
                       lcTexto2 = "Cpt: (Fecha: " & Format(mrs_Tmp2.Fields!fechaCreacion, sighEntidades.DevuelveFechaSoloFormato_DMY) & ") (N° Cuenta: " & Trim(Str(mrs_Tmp2.Fields!idCuentaAtencion)) & ") (Serv: " & Trim(mrs_Tmp2.Fields!DServicio) & ")"
                    ElseIf Not IsNull(mrs_Tmp1.Fields!observaciones) Then
                       lcTexto2 = mrs_Tmp1.Fields!observaciones
                    End If
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
                mrs_Tmp.Fields!Observacion1 = Left(lcTexto1, 100)
                mrs_Tmp.Fields!Observacion2 = Left(lcTexto2, 100)
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
            RpAHCauditoria.Sections("pie").Controls("lblPie").Caption = lcPie
            RpAHCauditoria.Orientation = rptOrientLandscape
            RpAHCauditoria.Show 1
        End If
    End If
End Sub

