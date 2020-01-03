VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clAtencionesTotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Atenciones por año
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Sub EjecutaFormulario()
    Dim oFormulario As New RpAtencionesTotales
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub PacientesYcpt(lnHwnd As Long, lcANIO As String)
    On Error GoTo ErrPacientesYcpt
    Dim oRsTmp1 As New Recordset, oRsEspecialidad As New Recordset, oRsTmp2 As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim lcPaciente As String, lnIdPaciente As Long, lcEspecialidad As String
    Dim lnEspecialidades As Long, lnNroCpt As Long
    Set oRsTmp1 = mo_ReglasAdmision.AtencionesSeleccionarXanio(lcANIO)
    If oRsTmp1.RecordCount = 0 Then
       MsgBox "No hay información", vbInformation, "Reporte"
    Else
'oRsTmp1.Filter = "historia= 510411086"
       With oRsEspecialidad
             .Fields.Append "Paciente", adVarChar, 200, adFldIsNullable
             .Fields.Append "EspecialidadCantidad", adInteger
             .Fields.Append "EspecialidadNombre", adVarChar, 100, adFldIsNullable
             .Fields.Append "CptCantidad", adInteger
             .LockType = adLockOptimistic
             .Open
       End With
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          lcPaciente = oRsTmp1!Paciente
          lnIdPaciente = oRsTmp1!idPaciente
          lcEspecialidad = oRsTmp1!especialidad
          lnEspecialidades = 0
          lnNroCpt = 0
          Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente
             Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioDespachoXcuenta(oRsTmp1!idCuentaAtencion)
             oRsTmp2.Filter = "idEstadoFacturacion<>9 and idEstadoFacturacion<>12"
             lnNroCpt = lnNroCpt + oRsTmp2.RecordCount
             lnEspecialidades = lnEspecialidades + 1
             oRsTmp1.MoveNext
             If oRsTmp1.EOF Then
                Exit Do
             End If
          Loop
          oRsEspecialidad.AddNew
          oRsEspecialidad.Fields!Paciente = lcPaciente
          oRsEspecialidad.Fields!EspecialidadCantidad = lnEspecialidades
          oRsEspecialidad.Fields!EspecialidadNombre = IIf(lnEspecialidades > 1, "", lcEspecialidad)
          oRsEspecialidad.Fields!CptCantidad = lnNroCpt
          oRsEspecialidad.Update
       Loop
       If oRsEspecialidad.RecordCount > 0 Then
           oRsEspecialidad.Sort = "paciente"
           mo_ReglasReportes.ExportarRecordSetAexcel oRsEspecialidad, "LISTA DE PACIENTES", "AÑO: " & lcANIO, "", lnHwnd
       End If
    
    End If
    Set oRsTmp1 = Nothing
    Set oRsEspecialidad = Nothing
    Set oRsTmp2 = Nothing
    Set mo_ReglasReportes = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasFacturacion = Nothing
    Exit Sub
ErrPacientesYcpt:
    MsgBox Err.Description
End Sub
        
Sub SISxDpto(lnHwnd As Long, lcANIO As String, lnOpcion As Integer)
        On Error GoTo ErrSisXdpto
        Dim oRsTmp1 As New Recordset
        Dim oRsHospital As New Recordset
        Dim oRsSIS As New Recordset
        Dim oRsSIStemp As New Recordset
        Dim oRsReferidos As New Recordset
        Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        Dim mo_ReglasReportes1 As New SIGHNegocios.ReglasReportes
        Dim lcSubTitulo As String, lcIdOrigenAtencion As String
        Dim lnIdPaciente As Long, lbNuevaFila As Boolean, lbEncontroDato As Boolean
        
        Const lcCodigosOrigenAtencion As String = "/12/12/21/22/33/34"
        Set oRsTmp1 = mo_ReglasAdmision.AtencionesSeleccionarXanioIngreso(lcANIO)
        oRsTmp1.Filter = "idTipoServicio<4"
        If oRsTmp1.RecordCount > 0 Then
            With oRsHospital
                .Fields.Append "TipoServicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Departamento", adVarChar, 100, adFldIsNullable
                .Fields.Append "EneroEgresos", adInteger
                .Fields.Append "EneroReingresos", adInteger
                .Fields.Append "FebreroEgresos", adInteger
                .Fields.Append "FebreroReingresos", adInteger
                .Fields.Append "MarzoEgresos", adInteger
                .Fields.Append "MarzoReingresos", adInteger
                .Fields.Append "AbrilEgresos", adInteger
                .Fields.Append "AbrilReingresos", adInteger
                .Fields.Append "MayoEgresos", adInteger
                .Fields.Append "MayoReingresos", adInteger
                .Fields.Append "JunioEgresos", adInteger
                .Fields.Append "JunioReingresos", adInteger
                .Fields.Append "JulioEgresos", adInteger
                .Fields.Append "JulioReingresos", adInteger
                .Fields.Append "AgostoEgresos", adInteger
                .Fields.Append "AgostoReingresos", adInteger
                .Fields.Append "SetiembreEgresos", adInteger
                .Fields.Append "SetiembreReingresos", adInteger
                .Fields.Append "OctubreEgresos", adInteger
                .Fields.Append "OctubreReingresos", adInteger
                .Fields.Append "NoviembreEgresos", adInteger
                .Fields.Append "NoviembreReingresos", adInteger
                .Fields.Append "DiciembreEgresos", adInteger
                .Fields.Append "DiciembreReingresos", adInteger
                .Fields.Append "TotalEgresos", adInteger
                .Fields.Append "TotalReingresos", adInteger
                .LockType = adLockOptimistic
                .Open
           End With
           With oRsReferidos
                .Fields.Append "TipoServicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Departamento", adVarChar, 100, adFldIsNullable
                .Fields.Append "Enero", adInteger
                .Fields.Append "Febrero", adInteger
                .Fields.Append "Marzo", adInteger
                .Fields.Append "Abril", adInteger
                .Fields.Append "Mayo", adInteger
                .Fields.Append "Junio", adInteger
                .Fields.Append "Julio", adInteger
                .Fields.Append "Agosto", adInteger
                .Fields.Append "Setiembre", adInteger
                .Fields.Append "Octubre", adInteger
                .Fields.Append "Noviembre", adInteger
                .Fields.Append "Diciembre", adInteger
                .Fields.Append "Total", adInteger
                .LockType = adLockOptimistic
                .Open
           End With
           With oRsSIS
                .Fields.Append "TipoServicio", adVarChar, 100, adFldIsNullable
                .Fields.Append "Departamento", adVarChar, 100, adFldIsNullable
                .Fields.Append "EneroSIS", adInteger
                .Fields.Append "FebreroSIS", adInteger
                .Fields.Append "MarzoSIS", adInteger
                .Fields.Append "AbrilSIS", adInteger
                .Fields.Append "MayoSIS", adInteger
                .Fields.Append "JunioSIS", adInteger
                .Fields.Append "JulioSIS", adInteger
                .Fields.Append "AgostoSIS", adInteger
                .Fields.Append "SetiembreSIS", adInteger
                .Fields.Append "OctubreSIS", adInteger
                .Fields.Append "NoviembreSIS", adInteger
                .Fields.Append "DiciembreSIS", adInteger
                .Fields.Append "TotalSIS", adInteger
                .Fields.Append "EneroNOSIS", adInteger
                .Fields.Append "FebreroNOSIS", adInteger
                .Fields.Append "MarzoNOSIS", adInteger
                .Fields.Append "AbrilNOSIS", adInteger
                .Fields.Append "MayoNOSIS", adInteger
                .Fields.Append "JunioNOSIS", adInteger
                .Fields.Append "JulioNOSIS", adInteger
                .Fields.Append "AgostoNOSIS", adInteger
                .Fields.Append "SetiembreNOSIS", adInteger
                .Fields.Append "OctubreNOSIS", adInteger
                .Fields.Append "NoviembreNOSIS", adInteger
                .Fields.Append "DiciembreNOSIS", adInteger
                .Fields.Append "TotalNOSIS", adInteger
                .Fields.Append "EneroSEMISUB", adInteger
                .Fields.Append "FebreroSEMISUB", adInteger
                .Fields.Append "MarzoSEMISUB", adInteger
                .Fields.Append "AbrilSEMISUB", adInteger
                .Fields.Append "MayoSEMISUB", adInteger
                .Fields.Append "JunioSEMISUB", adInteger
                .Fields.Append "JulioSEMISUB", adInteger
                .Fields.Append "AgostoSEMISUB", adInteger
                .Fields.Append "SetiembreSEMISUB", adInteger
                .Fields.Append "OctubreSEMISUB", adInteger
                .Fields.Append "NoviembreSEMISUB", adInteger
                .Fields.Append "DiciembreSEMISUB", adInteger
                .Fields.Append "TotalSEMISUB", adInteger
                .LockType = adLockOptimistic
                .Open
           End With
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
If oRsTmp1!idCuentaAtencion = 34232 Then
lnIdPaciente = 0
End If
              lnIdPaciente = oRsTmp1!idPaciente
              '
              If oRsSIStemp.State = 1 Then
                 Set oRsSIStemp = Nothing
              End If
              With oRsSIStemp
                     .Fields.Append "TipoServicio", adVarChar, 100, adFldIsNullable
                     .Fields.Append "Departamento", adVarChar, 100, adFldIsNullable
                     .Fields.Append "SoloSIS", adInteger, 4
                     .Fields.Append "SoloPagante", adInteger
                     .Fields.Append "Enero", adInteger, 4
                     .Fields.Append "Febrero", adInteger, 4
                     .Fields.Append "Marzo", adInteger, 4
                     .Fields.Append "Abril", adInteger, 4
                     .Fields.Append "Mayo", adInteger, 4
                     .Fields.Append "Junio", adInteger, 4
                     .Fields.Append "Julio", adInteger, 4
                     .Fields.Append "Agosto", adInteger, 4
                     .Fields.Append "Setiembre", adInteger, 4
                     .Fields.Append "Octubre", adInteger, 4
                     .Fields.Append "Noviembre", adInteger, 4
                     .Fields.Append "Diciembre", adInteger, 4
                     .Fields.Append "Total", adInteger, 4
                     .Fields.Append "Cuenta", adInteger, 4
                     .LockType = adLockOptimistic
                     .Open
              End With
              
              Do While Not oRsTmp1.EOF And lnIdPaciente = oRsTmp1!idPaciente
If oRsTmp1!idCuentaAtencion = 34232 Then
lcIdOrigenAtencion = ""
End If                    '
                    lcIdOrigenAtencion = "9999"
                    If Not IsNull(oRsTmp1!idOrigenAtencion) Then
                       lcIdOrigenAtencion = Trim(Str(oRsTmp1!idOrigenAtencion))
                    End If
                    '
                    If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                        lbNuevaFila = True
                        If oRsReferidos.RecordCount > 0 Then
                           oRsReferidos.MoveFirst
                           Do While Not oRsReferidos.EOF
                              If oRsReferidos!TipoServicio = oRsTmp1!TipoServicio And oRsReferidos!Departamento = oRsTmp1!dpto Then
                                 lbNuevaFila = False
                                 Exit Do
                              End If
                              oRsReferidos.MoveNext
                           Loop
                        End If
                        If lbNuevaFila = True Then
                           oRsReferidos.AddNew
                           oRsReferidos.Fields!TipoServicio = oRsTmp1!TipoServicio
                           oRsReferidos.Fields!Departamento = oRsTmp1!dpto
                        End If
                    End If
                    '
                    lbEncontroDato = False
                    lbNuevaFila = True
                    If oRsSIStemp.RecordCount > 0 Then
                         oRsSIStemp.MoveFirst
                         Do While Not oRsSIStemp.EOF
                            If Trim(oRsSIStemp!TipoServicio) = Trim(oRsTmp1!TipoServicio) And Trim(oRsSIStemp!Departamento) = Trim(oRsTmp1!dpto) Then
                               lbNuevaFila = False
                               Exit Do
                            End If
                            oRsSIStemp.MoveNext
                         Loop
                    End If
                    If lbNuevaFila = True Then
                         oRsSIStemp.AddNew
                         oRsSIStemp.Fields!TipoServicio = oRsTmp1!TipoServicio
                         oRsSIStemp.Fields!Departamento = oRsTmp1!dpto
                         lbEncontroDato = True
                    End If
                    '
                    Select Case Month(oRsTmp1.Fields!FechaIngreso)
                    Case 1
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!ENERO = oRsReferidos.Fields!ENERO + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!ENERO = oRsSIStemp.Fields!ENERO + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!ENERO = oRsSIStemp.Fields!ENERO + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 2
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Febrero = oRsReferidos.Fields!Febrero + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Febrero = oRsSIStemp.Fields!Febrero + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Febrero = oRsSIStemp.Fields!Febrero + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 3
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Marzo = oRsReferidos.Fields!Marzo + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Marzo = oRsSIStemp.Fields!Marzo + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Marzo = oRsSIStemp.Fields!Marzo + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 4
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Abril = oRsReferidos.Fields!Abril + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Abril = oRsSIStemp.Fields!Abril + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Abril = oRsSIStemp.Fields!Abril + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 5
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Mayo = oRsReferidos.Fields!Mayo + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Mayo = oRsSIStemp.Fields!Mayo + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Mayo = oRsSIStemp.Fields!Mayo + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 6
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Junio = oRsReferidos.Fields!Junio + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Junio = oRsSIStemp.Fields!Junio + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Junio = oRsSIStemp.Fields!Junio + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 7
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Julio = oRsReferidos.Fields!Julio + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Julio = oRsSIStemp.Fields!Julio + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Julio = oRsSIStemp.Fields!Julio + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 8
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Agosto = oRsReferidos.Fields!Agosto + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Agosto = oRsSIStemp.Fields!Agosto + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Agosto = oRsSIStemp.Fields!Agosto + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 9
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Setiembre = oRsReferidos.Fields!Setiembre + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Setiembre = oRsSIStemp.Fields!Setiembre + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Setiembre = oRsSIStemp.Fields!Setiembre + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 10
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Octubre = oRsReferidos.Fields!Octubre + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Octubre = oRsSIStemp.Fields!Octubre + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Octubre = oRsSIStemp.Fields!Octubre + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 11
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Noviembre = oRsReferidos.Fields!Noviembre + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Noviembre = oRsSIStemp.Fields!Noviembre + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Noviembre = oRsSIStemp.Fields!Noviembre + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    Case 12
                         If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                            oRsReferidos.Fields!Diciembre = oRsReferidos.Fields!Diciembre + 1
                         End If
                         '
                         If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                            oRsSIStemp.Fields!Diciembre = oRsSIStemp.Fields!Diciembre + 1
                            oRsSIStemp.Fields!soloSIS = 1
                            lbEncontroDato = True
                         ElseIf oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Or oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                            oRsSIStemp.Fields!Diciembre = oRsSIStemp.Fields!Diciembre + 1
                            oRsSIStemp.Fields!soloPagante = 1
                            lbEncontroDato = True
                         End If
                    End Select
                    '
                    If InStr(lcCodigosOrigenAtencion, lcIdOrigenAtencion) > 0 Then
                       oRsReferidos.Fields!Total = oRsReferidos.Fields!Total + 1
                       oRsReferidos.Update
                    End If
                    '
                    If lbEncontroDato = True Then
                       oRsSIStemp.Fields!cuenta = oRsTmp1.Fields!idPaciente
                       oRsSIStemp.Fields!Total = oRsSIStemp.Fields!Total + 1
                       oRsSIStemp.Update
                    End If
                    '
                    oRsTmp1.MoveNext
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
              Loop
              '
              If oRsSIStemp.RecordCount > 0 Then
                 oRsSIStemp.MoveFirst
                 Do While Not oRsSIStemp.EOF
                    If oRsSIStemp!soloPagante = 1 Or oRsSIStemp!soloSIS = 1 Then
                        lbNuevaFila = True
                        If oRsSIS.RecordCount > 0 Then
                             oRsSIS.MoveFirst
                             Do While Not oRsSIS.EOF
                                If Trim(oRsSIS!TipoServicio) = Trim(oRsSIStemp!TipoServicio) And Trim(oRsSIS!Departamento) = Trim(oRsSIStemp!Departamento) Then
                                   lbNuevaFila = False
                                   Exit Do
                                End If
                                oRsSIS.MoveNext
                             Loop
                        End If
                        If lbNuevaFila = True Then
                             oRsSIS.AddNew
                             oRsSIS.Fields!TipoServicio = oRsSIStemp!TipoServicio
                             oRsSIS.Fields!Departamento = oRsSIStemp!Departamento
                        End If
                        If oRsSIStemp.Fields!soloSIS = 1 And oRsSIStemp.Fields!soloPagante = 0 Then
                           oRsSIS.Fields!EneroSIS = oRsSIS.Fields!EneroSIS + oRsSIStemp!ENERO
                           oRsSIS.Fields!FebreroSIS = oRsSIS.Fields!FebreroSIS + oRsSIStemp!Febrero
                           oRsSIS.Fields!MarzoSIS = oRsSIS.Fields!MarzoSIS + oRsSIStemp!Marzo
                           oRsSIS.Fields!AbrilSIS = oRsSIS.Fields!AbrilSIS + oRsSIStemp!Abril
                           oRsSIS.Fields!MayoSIS = oRsSIS.Fields!MayoSIS + oRsSIStemp!Mayo
                           oRsSIS.Fields!JunioSIS = oRsSIS.Fields!JunioSIS + oRsSIStemp!Junio
                           oRsSIS.Fields!JulioSIS = oRsSIS.Fields!JulioSIS + oRsSIStemp!Julio
                           oRsSIS.Fields!AgostoSIS = oRsSIS.Fields!AgostoSIS + oRsSIStemp!Agosto
                           oRsSIS.Fields!SetiembreSIS = oRsSIS.Fields!SetiembreSIS + oRsSIStemp!Setiembre
                           oRsSIS.Fields!OctubreSIS = oRsSIS.Fields!OctubreSIS + oRsSIStemp!Octubre
                           oRsSIS.Fields!NoviembreSIS = oRsSIS.Fields!NoviembreSIS + oRsSIStemp!Noviembre
                           oRsSIS.Fields!DiciembreSIS = oRsSIS.Fields!DiciembreSIS + oRsSIStemp!Diciembre
                           oRsSIS.Fields!TotalSIS = oRsSIS.Fields!TotalSIS + oRsSIStemp!Total
                        ElseIf oRsSIStemp.Fields!soloSIS = 0 And oRsSIStemp.Fields!soloPagante = 1 Then
                           oRsSIS.Fields!EneroNOSIS = oRsSIS.Fields!EneroNOSIS + oRsSIStemp!ENERO
                           oRsSIS.Fields!FebreroNOSIS = oRsSIS.Fields!FebreroNOSIS + oRsSIStemp!Febrero
                           oRsSIS.Fields!MarzoNOSIS = oRsSIS.Fields!MarzoNOSIS + oRsSIStemp!Marzo
                           oRsSIS.Fields!AbrilNOSIS = oRsSIS.Fields!AbrilNOSIS + oRsSIStemp!Abril
                           oRsSIS.Fields!MayoNOSIS = oRsSIS.Fields!MayoNOSIS + oRsSIStemp!Mayo
                           oRsSIS.Fields!JunioNOSIS = oRsSIS.Fields!JunioNOSIS + oRsSIStemp!Junio
                           oRsSIS.Fields!JulioNOSIS = oRsSIS.Fields!JulioNOSIS + oRsSIStemp!Julio
                           oRsSIS.Fields!AgostoNOSIS = oRsSIS.Fields!AgostoNOSIS + oRsSIStemp!Agosto
                           oRsSIS.Fields!SetiembreNOSIS = oRsSIS.Fields!SetiembreNOSIS + oRsSIStemp!Setiembre
                           oRsSIS.Fields!OctubreNOSIS = oRsSIS.Fields!OctubreNOSIS + oRsSIStemp!Octubre
                           oRsSIS.Fields!NoviembreNOSIS = oRsSIS.Fields!NoviembreNOSIS + oRsSIStemp!Noviembre
                           oRsSIS.Fields!DiciembreNOSIS = oRsSIS.Fields!DiciembreNOSIS + oRsSIStemp!Diciembre
                           oRsSIS.Fields!TotalNOSIS = oRsSIS.Fields!TotalNOSIS + oRsSIStemp!Total
                        Else
If oRsSIStemp!Marzo > 0 Then
lbNuevaFila = True
End If
                           oRsSIS.Fields!EneroSEMISUB = oRsSIS.Fields!EneroSEMISUB + oRsSIStemp!ENERO
                           oRsSIS.Fields!FebreroSEMISUB = oRsSIS.Fields!FebreroSEMISUB + oRsSIStemp!Febrero
                           oRsSIS.Fields!MarzoSEMISUB = oRsSIS.Fields!MarzoSEMISUB + oRsSIStemp!Marzo
                           oRsSIS.Fields!AbrilSEMISUB = oRsSIS.Fields!AbrilSEMISUB + oRsSIStemp!Abril
                           oRsSIS.Fields!MayoSEMISUB = oRsSIS.Fields!MayoSEMISUB + oRsSIStemp!Mayo
                           oRsSIS.Fields!JunioSEMISUB = oRsSIS.Fields!JunioSEMISUB + oRsSIStemp!Junio
                           oRsSIS.Fields!JulioSEMISUB = oRsSIS.Fields!JulioSEMISUB + oRsSIStemp!Julio
                           oRsSIS.Fields!AgostoSEMISUB = oRsSIS.Fields!AgostoSEMISUB + oRsSIStemp!Agosto
                           oRsSIS.Fields!SetiembreSEMISUB = oRsSIS.Fields!SetiembreSEMISUB + oRsSIStemp!Setiembre
                           oRsSIS.Fields!OctubreSEMISUB = oRsSIS.Fields!OctubreSEMISUB + oRsSIStemp!Octubre
                           oRsSIS.Fields!NoviembreSEMISUB = oRsSIS.Fields!NoviembreSEMISUB + oRsSIStemp!Noviembre
                           oRsSIS.Fields!DiciembreSEMISUB = oRsSIS.Fields!DiciembreSEMISUB + oRsSIStemp!Diciembre
                           oRsSIS.Fields!TotalSEMISUB = oRsSIS.Fields!TotalSEMISUB + oRsSIStemp!Total
                        End If
                        oRsSIS.Update
                    End If
                    oRsSIStemp.MoveNext
                 Loop
              End If
              '
           Loop
           '
           lcSubTitulo = "Año: " & lcANIO
           If oRsReferidos.RecordCount > 0 Then
              oRsReferidos.Sort = "tipoServicio,Departamento"
              mo_ReglasReportes.ExportarRecordSetAexcel oRsReferidos, "REFERIDOS", lcSubTitulo, "", lnHwnd
           End If
           If oRsSIS.RecordCount > 0 Then
              oRsSIS.Sort = "tipoServicio,Departamento"
              mo_ReglasReportes1.ExportarRecordSetAexcel oRsSIS, "SIS/PAGANTES", lcSubTitulo, "", lnHwnd
           End If
           
           
        End If
        Exit Sub
ErrSisXdpto:
        If Err.Number > 0 Then
           MsgBox Err.Description
           Resume
        End If
        Set oRsTmp1 = Nothing
        Set mo_ReglasAdmision = Nothing
        Set oRsHospital = Nothing
        Set oRsSIS = Nothing
        Set oRsReferidos = Nothing
        Set oRsSIStemp = Nothing
        Set mo_ReglasReportes = Nothing
        Set mo_ReglasReportes1 = Nothing
End Sub



Sub ProcesaReporteParaHuelga(lnTipo As Integer, lcHospital As String, ldFechaInicial As Date, _
                             ldFechaFinal As Date, lnFiltro As Integer, lnHwnd As Long, lcTitulo As String, _
                             lcSubTitulo As String)
    On Error GoTo ErrRptHuelga
    Dim EXL As Excel.Application
    Set EXL = New Excel.Application
    Dim W As Excel.Workbook
    Dim s As Excel.Worksheet
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim oFila As Long, ldFecha As Date, lbNuevo As Boolean, lbOKmant As Boolean
    Dim ldFechaInicialHist As Date, ldFechaFinalHist As Date
    Dim lnNroConsultas As Long, lcFecha As String, lcHoraAtencion As String, lcTexto As String
    With oRsTmp1
          .Fields.Append "Fecha", adDate
          .Fields.Append "NroConsultasHistorico", adInteger
          .Fields.Append "NroConsultasHuelga", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    Select Case lnFiltro
    Case 1   'Año pasado
         ldFechaInicialHist = CDate(Format(ldFechaInicial, "dd/mm/") & Trim(Str(Year(ldFechaInicial) - 1)))
         ldFechaFinalHist = CDate(Format(ldFechaFinal, "dd/mm/") & Trim(Str(Year(ldFechaFinal) - 1)))
    Case 2   'Semana pasada
         ldFechaInicialHist = ldFechaInicial - 7
         ldFechaFinalHist = ldFechaFinal - 7
    End Select
    Select Case lnTipo
    Case 0   '************************ATENCION DEL MEDICO**************************
         'historicos (solo DEL AÑO PASADO)
         If lnFiltro = 1 Then
            Set W = EXL.Workbooks.Open(App.Path & "\Archivos\excel2013atendidos.xls")
            Set s = W.Sheets("RPT_EGRESO_HOSP")
            oFila = 4
            Do While True
               lcFecha = s.Cells(oFila, 16).Value
               If lcFecha = "" Then
                  Exit Do
               End If
               ldFecha = CDate(Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY))
               If IsDate(ldFecha) = False Then
                  Exit Do
               End If
               lcHoraAtencion = s.Cells(oFila, 19).Value
               If ldFecha >= ldFechaInicialHist And ldFecha <= ldFechaFinalHist And Len(Trim(lcHoraAtencion)) = 5 Then
                  lbNuevo = True
                  If oRsTmp1.RecordCount > 0 Then
                     oRsTmp1.MoveFirst
                     oRsTmp1.Find "fecha='" & ldFecha & "'"
                     If Not oRsTmp1.EOF Then
                        lbNuevo = False
                     End If
                  End If
                  If lbNuevo = True Then
                     oRsTmp1.AddNew
                     oRsTmp1.Fields!fecha = ldFecha
                     oRsTmp1.Fields!NroConsultasHistorico = 1
                     oRsTmp1.Fields!NroConsultasHuelga = 0
                  Else
                     lnNroConsultas = oRsTmp1.Fields!NroConsultasHistorico
                     oRsTmp1.Fields!NroConsultasHistorico = lnNroConsultas + 1
                  End If
                  oRsTmp1.Update
               End If
               oFila = oFila + 1
            Loop
         End If
         'huelga
         Set W = EXL.Workbooks.Open(App.Path & "\Archivos\excel2014atendidos.xls")
         Set s = W.Sheets("RPT_EGRESO_HOSP")
         oFila = 4
         Do While True
            lcFecha = s.Cells(oFila, 16).Value
            If lcFecha = "" Then
               Exit Do
            End If
            ldFecha = CDate(Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY))
            If IsDate(ldFecha) = False Then
               Exit Do
            End If
            lcHoraAtencion = s.Cells(oFila, 19).Value
            If ldFecha >= ldFechaInicial And ldFecha <= ldFechaFinal And Len(Trim(lcHoraAtencion)) = 5 Then
               '****** procesa Huelga
               lbNuevo = True
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  oRsTmp1.Find "fecha='" & ldFecha & "'"
                  If Not oRsTmp1.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo = True Then
                  oRsTmp1.AddNew
                  oRsTmp1.Fields!fecha = ldFecha
                  oRsTmp1.Fields!NroConsultasHistorico = 0
                  oRsTmp1.Fields!NroConsultasHuelga = 1
               Else
                  lnNroConsultas = oRsTmp1.Fields!NroConsultasHuelga
                  oRsTmp1.Fields!NroConsultasHuelga = lnNroConsultas + 1
               End If
               oRsTmp1.Update
            ElseIf ldFecha >= ldFechaInicialHist And ldFecha <= ldFechaFinalHist Then
               '****** procesa Historico (DE LA SEMANA ANTERIOR)
               lbNuevo = True
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  oRsTmp1.Find "fecha='" & ldFecha & "'"
                  If Not oRsTmp1.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo = True Then
                  oRsTmp1.AddNew
                  oRsTmp1.Fields!fecha = ldFecha
                  oRsTmp1.Fields!NroConsultasHistorico = 1
                  oRsTmp1.Fields!NroConsultasHuelga = 0
               Else
                  lnNroConsultas = oRsTmp1.Fields!NroConsultasHistorico
                  oRsTmp1.Fields!NroConsultasHistorico = lnNroConsultas + 1
               End If
               oRsTmp1.Update
            End If
            oFila = oFila + 1
            
            oRsTmp2.Open "update parametros set ValorInt=" & oFila & " where idParametro=1", sighentidades.CadenaConexion, adOpenKeyset, adLockOptimistic
If oFila > 7254 Then
lcFecha = ""
End If
            
         Loop
    Case 1   '************************MOVIMIENTOS DE HISTORIAS**************************
         Set W = EXL.Workbooks.Open(App.Path & "\Archivos\MovimientoHC.xls")
         Set s = W.Sheets("vistaParaMovimientosHistorias")
         oFila = 2
         Do While True
            lcFecha = s.Cells(oFila, 3).Value
            If lcFecha = "" Then
               Exit Do
            End If
            ldFecha = CDate(Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY))
            If IsDate(ldFecha) = False Then
               Exit Do
            End If
If oFila = 26888 Then
lbNuevo = True
End If
            If ldFecha >= ldFechaInicial And ldFecha <= ldFechaFinal Then
               '****** procesa Huelga
               lbNuevo = True
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  oRsTmp1.Find "fecha='" & ldFecha & "'"
                  If Not oRsTmp1.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo = True Then
                  oRsTmp1.AddNew
                  oRsTmp1.Fields!fecha = ldFecha
                  oRsTmp1.Fields!NroConsultasHistorico = 0
                  oRsTmp1.Fields!NroConsultasHuelga = 1
               Else
                  lnNroConsultas = oRsTmp1.Fields!NroConsultasHuelga
                  oRsTmp1.Fields!NroConsultasHuelga = lnNroConsultas + 1
               End If
               oRsTmp1.Update
            ElseIf ldFecha >= ldFechaInicialHist And ldFecha <= ldFechaFinalHist Then
               '****** procesa Historico
               lbNuevo = True
               If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  oRsTmp1.Find "fecha='" & ldFecha & "'"
                  If Not oRsTmp1.EOF Then
                     lbNuevo = False
                  End If
               End If
               If lbNuevo = True Then
                  oRsTmp1.AddNew
                  oRsTmp1.Fields!fecha = ldFecha
                  oRsTmp1.Fields!NroConsultasHistorico = 1
                  oRsTmp1.Fields!NroConsultasHuelga = 0
               Else
                  lnNroConsultas = oRsTmp1.Fields!NroConsultasHistorico
                  oRsTmp1.Fields!NroConsultasHistorico = lnNroConsultas + 1
               End If
               oRsTmp1.Update
            End If
            oFila = oFila + 1
         Loop
    Case 2   '************************CITADOS**************************
         'huelga
         Set W = EXL.Workbooks.Open(App.Path & "\Archivos\excel2014auditoria.xls")
         Set s = W.Sheets("hoja_Libre")
         oFila = 6
         Do While True
            lcFecha = Format(s.Cells(oFila, 1).Value, sighentidades.DevuelveFechaSoloFormato_DMY)
            If lcFecha = "" Then
               Exit Do
            End If
            ldFecha = CDate(Format(lcFecha, sighentidades.DevuelveFechaSoloFormato_DMY))
            If IsDate(ldFecha) = False Then
               Exit Do
            End If
            lcTexto = s.Cells(oFila, 3)  'Agregó
            If ldFecha >= ldFechaInicial And ldFecha <= ldFechaFinal And UCase(Left(lcTexto, 1)) = "A" Then
               '****** busca si sacaron CITA el mismo día
               lcTexto = s.Cells(oFila, 8)
               If InStr(lcTexto, lcFecha) > 0 Then
                    '****** procesa Huelga
                    lbNuevo = True
                    If oRsTmp1.RecordCount > 0 Then
                       oRsTmp1.MoveFirst
                       oRsTmp1.Find "fecha='" & ldFecha & "'"
                       If Not oRsTmp1.EOF Then
                          lbNuevo = False
                       End If
                    End If
                    If lbNuevo = True Then
                       oRsTmp1.AddNew
                       oRsTmp1.Fields!fecha = ldFecha
                       oRsTmp1.Fields!NroConsultasHistorico = 0
                       oRsTmp1.Fields!NroConsultasHuelga = 1
                    Else
                       lnNroConsultas = oRsTmp1.Fields!NroConsultasHuelga
                       oRsTmp1.Fields!NroConsultasHuelga = lnNroConsultas + 1
                    End If
                    oRsTmp1.Update
               End If
            End If
            '****** procesa Historico (DE LA SEMANA ANTERIOR)
            ldFecha = CDate("01/01/2000")
            lcFecha = s.Cells(oFila, 8)
            If InStr(lcFecha, "Fecha:") > 0 Then
               lcFecha = Mid(lcFecha, InStr(lcFecha, "Fecha:") + 7, 10)
               ldFecha = CDate(lcFecha)
            End If
            lcTexto = s.Cells(oFila, 3)  'Agregó
            If ldFecha >= ldFechaInicialHist And ldFecha <= ldFechaFinalHist And UCase(Left(lcTexto, 1)) = "A" Then

               lcTexto = s.Cells(oFila, 1)
               If InStr(lcTexto, lcFecha) > 0 Then
                    lbNuevo = True
                    If oRsTmp1.RecordCount > 0 Then
                       oRsTmp1.MoveFirst
                       oRsTmp1.Find "fecha='" & ldFecha & "'"
                       If Not oRsTmp1.EOF Then
                          lbNuevo = False
                       End If
                    End If
                    If lbNuevo = True Then
                       oRsTmp1.AddNew
                       oRsTmp1.Fields!fecha = ldFecha
                       oRsTmp1.Fields!NroConsultasHistorico = 1
                       oRsTmp1.Fields!NroConsultasHuelga = 0
                    Else
                       lnNroConsultas = oRsTmp1.Fields!NroConsultasHistorico
                       oRsTmp1.Fields!NroConsultasHistorico = lnNroConsultas + 1
                    End If
                    oRsTmp1.Update
               End If
            End If
            oFila = oFila + 1
            
            
            
         Loop
    End Select
    'Envia al Excel
    oRsTmp1.Sort = "fecha"
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp1, lcTitulo, lcSubTitulo, "", lnHwnd
    '
    W.Close False
    Set s = Nothing
    Set W = Nothing
    Set EXL = Nothing
    Set oRsTmp2 = Nothing
    Exit Sub
ErrRptHuelga:
    MsgBox Err.Description
    Resume
End Sub

Sub ProcesaReporte(lcANIO As String, ml_opcion As Integer, ml_opcionSub As Integer, lcTitulo As String, _
                   lnHwnd As Long)
    Dim iFila As Long, lnDiasEstancia As Double
    Dim lnEstancia As Integer, lnConsumoServicio As Double, lnConsumoFarmacia As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim oRsAtenciones As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim lcHoraEstanciaMax As String
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    Dim lcSql As String
    Dim lbEsOpenOffice As Boolean
    
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    
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
    
    
    With oRsAtenciones
          .Fields.Append "Area", adVarChar, 20, adFldIsNullable
          .Fields.Append "EneroSis", adInteger, 4, adFldIsNullable
          .Fields.Append "EneroOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "EneroSisAcum", adDouble
          .Fields.Append "EneroOtrosAcum", adDouble
          .Fields.Append "FebreroSis", adInteger, 4, adFldIsNullable
          .Fields.Append "FebreroOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "FebreroSisAcum", adDouble
          .Fields.Append "FebreroOtrosAcum", adDouble
          .Fields.Append "MarzoSis", adInteger, 4, adFldIsNullable
          .Fields.Append "MarzoOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "MarzoSisAcum", adDouble
          .Fields.Append "MarzoOtrosAcum", adDouble
          .Fields.Append "AbrilSis", adInteger, 4, adFldIsNullable
          .Fields.Append "AbrilOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "AbrilSisAcum", adDouble
          .Fields.Append "AbrilOtrosAcum", adDouble
          .Fields.Append "MayoSis", adInteger, 4, adFldIsNullable
          .Fields.Append "MayoOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "MayoSisAcum", adDouble
          .Fields.Append "MayoOtrosAcum", adDouble
          .Fields.Append "JunioSis", adInteger, 4, adFldIsNullable
          .Fields.Append "JunioOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "JunioSisAcum", adDouble
          .Fields.Append "JunioOtrosAcum", adDouble
          .Fields.Append "JulioSis", adInteger, 4, adFldIsNullable
          .Fields.Append "JulioOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "JulioSisAcum", adDouble
          .Fields.Append "JulioOtrosAcum", adDouble
          .Fields.Append "AgostoSis", adInteger, 4, adFldIsNullable
          .Fields.Append "AgostoOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "AgostoSisAcum", adDouble
          .Fields.Append "AgostoOtrosAcum", adDouble
          .Fields.Append "SetiembreSis", adInteger, 4, adFldIsNullable
          .Fields.Append "SetiembreOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "SetiembreSisAcum", adDouble
          .Fields.Append "SetiembreOtrosAcum", adDouble
          .Fields.Append "OctubreSis", adInteger, 4, adFldIsNullable
          .Fields.Append "OctubreOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "OctubreSisAcum", adDouble
          .Fields.Append "OctubreOtrosAcum", adDouble
          .Fields.Append "NoviembreSis", adInteger, 4, adFldIsNullable
          .Fields.Append "NoviembreOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "NoviembreSisAcum", adDouble
          .Fields.Append "NoviembreOtrosAcum", adDouble
          .Fields.Append "DiciembreSis", adInteger, 4, adFldIsNullable
          .Fields.Append "DiciembreOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "DiciembreSisAcum", adDouble
          .Fields.Append "DiciembreOtrosAcum", adDouble
          .Fields.Append "TotalSis", adInteger, 4, adFldIsNullable
          .Fields.Append "TotalOtros", adInteger, 4, adFldIsNullable
          .Fields.Append "TotalSisAcum", adDouble
          .Fields.Append "TotalOtrosAcum", adDouble
          .LockType = adLockOptimistic
          .Open
          .AddNew
          .Fields!Area = "Consulta Externa"
          .Fields!EneroSIS = 0
          .Fields!EneroOtros = 0
          .Fields!FebreroSIS = 0
          .Fields!FebreroOtros = 0
          .Fields!MarzoSIS = 0
          .Fields!MarzoOtros = 0
          .Fields!AbrilSIS = 0
          .Fields!AbrilOtros = 0
          .Fields!MayoSIS = 0
          .Fields!MayoOtros = 0
          .Fields!JunioSIS = 0
          .Fields!JunioOtros = 0
          .Fields!JulioSIS = 0
          .Fields!JulioOtros = 0
          .Fields!AgostoSIS = 0
          .Fields!AgostoOtros = 0
          .Fields!SetiembreSIS = 0
          .Fields!SetiembreOtros = 0
          .Fields!OctubreSIS = 0
          .Fields!OctubreOtros = 0
          .Fields!NoviembreSIS = 0
          .Fields!NoviembreOtros = 0
          .Fields!DiciembreSIS = 0
          .Fields!DiciembreOtros = 0
          .Fields!TotalSIS = 0
          .Fields!TotalOtros = 0
          .Update
          .AddNew
          .Fields!Area = "Emergencia"
          .Fields!EneroSIS = 0
          .Fields!EneroOtros = 0
          .Fields!FebreroSIS = 0
          .Fields!FebreroOtros = 0
          .Fields!MarzoSIS = 0
          .Fields!MarzoOtros = 0
          .Fields!AbrilSIS = 0
          .Fields!AbrilOtros = 0
          .Fields!MayoSIS = 0
          .Fields!MayoOtros = 0
          .Fields!JunioSIS = 0
          .Fields!JunioOtros = 0
          .Fields!JulioSIS = 0
          .Fields!JulioOtros = 0
          .Fields!AgostoSIS = 0
          .Fields!AgostoOtros = 0
          .Fields!SetiembreSIS = 0
          .Fields!SetiembreOtros = 0
          .Fields!OctubreSIS = 0
          .Fields!OctubreOtros = 0
          .Fields!NoviembreSIS = 0
          .Fields!NoviembreOtros = 0
          .Fields!DiciembreSIS = 0
          .Fields!DiciembreOtros = 0
          .Fields!TotalSIS = 0
          .Fields!TotalOtros = 0
          .Update
          .AddNew
          .Fields!Area = "Hospitalizacion"
          .Fields!EneroSIS = 0
          .Fields!EneroOtros = 0
          .Fields!FebreroSIS = 0
          .Fields!FebreroOtros = 0
          .Fields!MarzoSIS = 0
          .Fields!MarzoOtros = 0
          .Fields!AbrilSIS = 0
          .Fields!AbrilOtros = 0
          .Fields!MayoSIS = 0
          .Fields!MayoOtros = 0
          .Fields!JunioSIS = 0
          .Fields!JunioOtros = 0
          .Fields!JulioSIS = 0
          .Fields!JulioOtros = 0
          .Fields!AgostoSIS = 0
          .Fields!AgostoOtros = 0
          .Fields!SetiembreSIS = 0
          .Fields!SetiembreOtros = 0
          .Fields!OctubreSIS = 0
          .Fields!OctubreOtros = 0
          .Fields!NoviembreSIS = 0
          .Fields!NoviembreOtros = 0
          .Fields!DiciembreSIS = 0
          .Fields!DiciembreOtros = 0
          .Fields!TotalSIS = 0
          .Fields!TotalOtros = 0
          .Update
    End With
    'Procesa datos
    Select Case ml_opcion
    Case 1   'atenciones
        Set oRsTmp1 = mo_ReglasAdmision.AtencionesActivasFiltraXanioIngreso(lcANIO)
        
        
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              oRsAtenciones.MoveFirst
              If oRsTmp1.Fields!idTipoServicio > 1 Then
                 oRsAtenciones.Move (oRsTmp1.Fields!idTipoServicio - 1)
              End If
              Select Case Month(oRsTmp1.Fields!FechaIngreso)
              Case 1
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!EneroSIS = oRsAtenciones.Fields!EneroSIS + 1
                   Else
                      oRsAtenciones.Fields!EneroOtros = oRsAtenciones.Fields!EneroOtros + 1
                   End If
              Case 2
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!FebreroSIS = oRsAtenciones.Fields!FebreroSIS + 1
                   Else
                      oRsAtenciones.Fields!FebreroOtros = oRsAtenciones.Fields!FebreroOtros + 1
                   End If
              Case 3
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MarzoSIS = oRsAtenciones.Fields!MarzoSIS + 1
                   Else
                      oRsAtenciones.Fields!MarzoOtros = oRsAtenciones.Fields!MarzoOtros + 1
                   End If
              Case 4
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AbrilSIS = oRsAtenciones.Fields!AbrilSIS + 1
                   Else
                      oRsAtenciones.Fields!AbrilOtros = oRsAtenciones.Fields!AbrilOtros + 1
                   End If
              Case 5
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MayoSIS = oRsAtenciones.Fields!MayoSIS + 1
                   Else
                      oRsAtenciones.Fields!MayoOtros = oRsAtenciones.Fields!MayoOtros + 1
                   End If
              Case 6
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JunioSIS = oRsAtenciones.Fields!JunioSIS + 1
                   Else
                      oRsAtenciones.Fields!JunioOtros = oRsAtenciones.Fields!JunioOtros + 1
                   End If
              Case 7
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JulioSIS = oRsAtenciones.Fields!JulioSIS + 1
                   Else
                      oRsAtenciones.Fields!JulioOtros = oRsAtenciones.Fields!JulioOtros + 1
                   End If
              Case 8
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AgostoSIS = oRsAtenciones.Fields!AgostoSIS + 1
                   Else
                      oRsAtenciones.Fields!AgostoOtros = oRsAtenciones.Fields!AgostoOtros + 1
                   End If
              Case 9
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!SetiembreSIS = oRsAtenciones.Fields!SetiembreSIS + 1
                   Else
                      oRsAtenciones.Fields!SetiembreOtros = oRsAtenciones.Fields!SetiembreOtros + 1
                   End If
              Case 10
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!OctubreSIS = oRsAtenciones.Fields!OctubreSIS + 1
                   Else
                      oRsAtenciones.Fields!OctubreOtros = oRsAtenciones.Fields!OctubreOtros + 1
                   End If
              Case 11
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!NoviembreSIS = oRsAtenciones.Fields!NoviembreSIS + 1
                   Else
                      oRsAtenciones.Fields!NoviembreOtros = oRsAtenciones.Fields!NoviembreOtros + 1
                   End If
              Case 12
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!DiciembreSIS = oRsAtenciones.Fields!DiciembreSIS + 1
                   Else
                      oRsAtenciones.Fields!DiciembreOtros = oRsAtenciones.Fields!DiciembreOtros + 1
                   End If
              End Select
              If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                 oRsAtenciones.Fields!TotalSIS = oRsAtenciones.Fields!TotalSIS + 1
              Else
                 oRsAtenciones.Fields!TotalOtros = oRsAtenciones.Fields!TotalOtros + 1
              End If
              oRsAtenciones.Update
              oRsTmp1.MoveNext
           Loop
        Else
           MsgBox "No hay datos", vbInformation, "Reporte"
           Exit Sub
        End If
        oRsTmp1.Close
    Case 2, 3  'Promedio dias hosp, Promedio Facturacion hosp
        oRsAtenciones.MoveLast
        Set oRsTmp1 = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, CDate("01/01/" & lcANIO), CDate("31/12/" & lcANIO), 3)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If ml_opcion = 2 Then
                 lnEstancia = oBuscaDiasPaciente.DiasDelPacienteEnHospitalizacionEmergencia(oRsTmp1!FechaIngreso, oRsTmp1!HoraIngreso, oRsTmp1!FechaEgreso, oRsTmp1!horaEgreso, lcHoraEstanciaMax)
                 lnDiasEstancia = lnEstancia
              ElseIf ml_opcion = 3 Then
                  If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFParticularHospitalizado Then
                        'Consumo Servicios-Particular
                        lnConsumoServicio = mo_ReglasFacturacion.RetornaTotalConsumoServiciosParaPagantePorNroCuenta(oRsTmp1!idCuentaAtencion)
                        'Consumo Farmacia-Particular
                        lnConsumoFarmacia = mo_ReglasFacturacion.RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta(oRsTmp1!idCuentaAtencion)
                  Else
                        'Consumo Servicios-Seguros
                        lnConsumoServicio = mo_ReglasFacturacion.RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(oRsTmp1!idCuentaAtencion)
                        'Consumo Farmacia-Seguros
                        lnConsumoFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(oRsTmp1!idCuentaAtencion)
                  End If
                  Select Case ml_opcionSub
                  Case 1    'todos
                       lnDiasEstancia = lnConsumoServicio + lnConsumoFarmacia
                  Case 2    'servicios
                       lnDiasEstancia = lnConsumoServicio
                  Case 3    'farmacia
                       lnDiasEstancia = lnConsumoFarmacia
                  End Select

              End If
              Select Case Month(oRsTmp1.Fields!FechaEgreso)
              Case 1
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!EneroSIS = oRsAtenciones.Fields!EneroSIS + 1
                      oRsAtenciones.Fields!EneroSisAcum = oRsAtenciones.Fields!EneroSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!EneroOtros = oRsAtenciones.Fields!EneroOtros + 1
                      oRsAtenciones.Fields!EneroOtrosAcum = oRsAtenciones.Fields!EneroOtrosAcum + lnDiasEstancia
                   End If
              Case 2
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!FebreroSIS = oRsAtenciones.Fields!FebreroSIS + 1
                      oRsAtenciones.Fields!FebreroSisAcum = oRsAtenciones.Fields!FebreroSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!FebreroOtros = oRsAtenciones.Fields!FebreroOtros + 1
                      oRsAtenciones.Fields!FebreroOtrosAcum = oRsAtenciones.Fields!FebreroOtrosAcum + lnDiasEstancia
                   End If
              Case 3
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MarzoSIS = oRsAtenciones.Fields!MarzoSIS + 1
                      oRsAtenciones.Fields!MarzoSisAcum = oRsAtenciones.Fields!MarzoSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!MarzoOtros = oRsAtenciones.Fields!MarzoOtros + 1
                      oRsAtenciones.Fields!MarzoOtrosAcum = oRsAtenciones.Fields!MarzoOtrosAcum + lnDiasEstancia
                   End If
              Case 4
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AbrilSIS = oRsAtenciones.Fields!AbrilSIS + 1
                      oRsAtenciones.Fields!AbrilSisAcum = oRsAtenciones.Fields!AbrilSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!AbrilOtros = oRsAtenciones.Fields!AbrilOtros + 1
                      oRsAtenciones.Fields!AbrilOtrosAcum = oRsAtenciones.Fields!AbrilOtrosAcum + lnDiasEstancia
                   End If
              Case 5
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MayoSIS = oRsAtenciones.Fields!MayoSIS + 1
                      oRsAtenciones.Fields!MayoSisAcum = oRsAtenciones.Fields!MayoSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!MayoOtros = oRsAtenciones.Fields!MayoOtros + 1
                      oRsAtenciones.Fields!MayoOtrosAcum = oRsAtenciones.Fields!MayoOtrosAcum + lnDiasEstancia
                   End If
              Case 6
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JunioSIS = oRsAtenciones.Fields!JunioSIS + 1
                      oRsAtenciones.Fields!JunioSisAcum = oRsAtenciones.Fields!JunioSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!JunioOtros = oRsAtenciones.Fields!JunioOtros + 1
                      oRsAtenciones.Fields!JunioOtrosAcum = oRsAtenciones.Fields!JunioOtrosAcum + lnDiasEstancia
                   End If
              Case 7
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JulioSIS = oRsAtenciones.Fields!JulioSIS + 1
                      oRsAtenciones.Fields!JulioSisAcum = oRsAtenciones.Fields!JulioSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!JulioOtros = oRsAtenciones.Fields!JulioOtros + 1
                      oRsAtenciones.Fields!JulioOtrosAcum = oRsAtenciones.Fields!JulioOtrosAcum + lnDiasEstancia
                   End If
              Case 8
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AgostoSIS = oRsAtenciones.Fields!AgostoSIS + 1
                      oRsAtenciones.Fields!AgostoSisAcum = oRsAtenciones.Fields!AgostoSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!AgostoOtros = oRsAtenciones.Fields!AgostoOtros + 1
                      oRsAtenciones.Fields!AgostoOtrosAcum = oRsAtenciones.Fields!AgostoOtrosAcum + lnDiasEstancia
                   End If
              Case 9
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!SetiembreSIS = oRsAtenciones.Fields!SetiembreSIS + 1
                      oRsAtenciones.Fields!SetiembreSisAcum = oRsAtenciones.Fields!SetiembreSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!SetiembreOtros = oRsAtenciones.Fields!SetiembreOtros + 1
                      oRsAtenciones.Fields!SetiembreOtrosAcum = oRsAtenciones.Fields!SetiembreOtrosAcum + lnDiasEstancia
                   End If
              Case 10
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!OctubreSIS = oRsAtenciones.Fields!OctubreSIS + 1
                      oRsAtenciones.Fields!OctubreSisAcum = oRsAtenciones.Fields!OctubreSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!OctubreOtros = oRsAtenciones.Fields!OctubreOtros + 1
                      oRsAtenciones.Fields!OctubreOtrosAcum = oRsAtenciones.Fields!OctubreOtrosAcum + lnDiasEstancia
                   End If
              Case 11
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!NoviembreSIS = oRsAtenciones.Fields!NoviembreSIS + 1
                      oRsAtenciones.Fields!NoviembreSisAcum = oRsAtenciones.Fields!NoviembreSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!NoviembreOtros = oRsAtenciones.Fields!NoviembreOtros + 1
                      oRsAtenciones.Fields!NoviembreOtrosAcum = oRsAtenciones.Fields!NoviembreOtrosAcum + lnDiasEstancia
                   End If
              Case 12
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!DiciembreSIS = oRsAtenciones.Fields!DiciembreSIS + 1
                      oRsAtenciones.Fields!DiciembreSisAcum = oRsAtenciones.Fields!DiciembreSisAcum + lnDiasEstancia
                   Else
                      oRsAtenciones.Fields!DiciembreOtros = oRsAtenciones.Fields!DiciembreOtros + 1
                      oRsAtenciones.Fields!DiciembreOtrosAcum = oRsAtenciones.Fields!DiciembreOtrosAcum + lnDiasEstancia
                   End If
              End Select
              If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                 oRsAtenciones.Fields!TotalSIS = oRsAtenciones.Fields!TotalSIS + 1
                 oRsAtenciones.Fields!totalSisAcum = oRsAtenciones.Fields!totalSisAcum + lnDiasEstancia
              Else
                 oRsAtenciones.Fields!TotalOtros = oRsAtenciones.Fields!TotalOtros + 1
                 oRsAtenciones.Fields!TotalOtrosAcum = oRsAtenciones.Fields!TotalOtrosAcum + lnDiasEstancia
              End If
              oRsAtenciones.Update
              oRsTmp1.MoveNext
           Loop
            If oRsAtenciones.Fields!EneroSIS > 0 Then
               oRsAtenciones.Fields!EneroSisAcum = Round(oRsAtenciones.Fields!EneroSisAcum / oRsAtenciones.Fields!EneroSIS, 2)
            End If
            If oRsAtenciones.Fields!EneroOtros > 0 Then
               oRsAtenciones.Fields!EneroOtrosAcum = Round(oRsAtenciones.Fields!EneroOtrosAcum / oRsAtenciones.Fields!EneroOtros, 2)
            End If
            If oRsAtenciones.Fields!FebreroSIS > 0 Then
               oRsAtenciones.Fields!FebreroSisAcum = Round(oRsAtenciones.Fields!FebreroSisAcum / oRsAtenciones.Fields!FebreroSIS, 2)
            End If
            If oRsAtenciones.Fields!FebreroOtros > 0 Then
               oRsAtenciones.Fields!FebreroOtrosAcum = Round(oRsAtenciones.Fields!FebreroOtrosAcum / oRsAtenciones.Fields!FebreroOtros, 2)
            End If
            If oRsAtenciones.Fields!MarzoSIS > 0 Then
               oRsAtenciones.Fields!MarzoSisAcum = Round(oRsAtenciones.Fields!MarzoSisAcum / oRsAtenciones.Fields!MarzoSIS, 2)
            End If
            If oRsAtenciones.Fields!MarzoOtros > 0 Then
               oRsAtenciones.Fields!MarzoOtrosAcum = Round(oRsAtenciones.Fields!MarzoOtrosAcum / oRsAtenciones.Fields!MarzoOtros, 2)
            End If
            If oRsAtenciones.Fields!AbrilSIS > 0 Then
               oRsAtenciones.Fields!AbrilSisAcum = Round(oRsAtenciones.Fields!AbrilSisAcum / oRsAtenciones.Fields!AbrilSIS, 2)
            End If
            If oRsAtenciones.Fields!AbrilOtros > 0 Then
               oRsAtenciones.Fields!AbrilOtrosAcum = Round(oRsAtenciones.Fields!AbrilOtrosAcum / oRsAtenciones.Fields!AbrilOtros, 2)
            End If
            If oRsAtenciones.Fields!MayoSIS > 0 Then
               oRsAtenciones.Fields!MayoSisAcum = Round(oRsAtenciones.Fields!MayoSisAcum / oRsAtenciones.Fields!MayoSIS, 2)
            End If
            If oRsAtenciones.Fields!MayoOtros > 0 Then
               oRsAtenciones.Fields!MayoOtrosAcum = Round(oRsAtenciones.Fields!MayoOtrosAcum / oRsAtenciones.Fields!MayoOtros, 2)
            End If
            If oRsAtenciones.Fields!JunioSIS > 0 Then
               oRsAtenciones.Fields!JunioSisAcum = Round(oRsAtenciones.Fields!JunioSisAcum / oRsAtenciones.Fields!JunioSIS, 2)
            End If
            If oRsAtenciones.Fields!JunioOtros > 0 Then
               oRsAtenciones.Fields!JunioOtrosAcum = Round(oRsAtenciones.Fields!JunioOtrosAcum / oRsAtenciones.Fields!JunioOtros, 2)
            End If
            If oRsAtenciones.Fields!JulioSIS > 0 Then
               oRsAtenciones.Fields!JulioSisAcum = Round(oRsAtenciones.Fields!JulioSisAcum / oRsAtenciones.Fields!JulioSIS, 2)
            End If
            If oRsAtenciones.Fields!JulioOtros > 0 Then
               oRsAtenciones.Fields!JulioOtrosAcum = Round(oRsAtenciones.Fields!JulioOtrosAcum / oRsAtenciones.Fields!JulioOtros, 2)
            End If
            If oRsAtenciones.Fields!AgostoSIS > 0 Then
               oRsAtenciones.Fields!AgostoSisAcum = Round(oRsAtenciones.Fields!AgostoSisAcum / oRsAtenciones.Fields!AgostoSIS, 2)
            End If
            If oRsAtenciones.Fields!AgostoOtros > 0 Then
               oRsAtenciones.Fields!AgostoOtrosAcum = Round(oRsAtenciones.Fields!AgostoOtrosAcum / oRsAtenciones.Fields!AgostoOtros, 2)
            End If
            If oRsAtenciones.Fields!SetiembreSIS > 0 Then
               oRsAtenciones.Fields!SetiembreSisAcum = Round(oRsAtenciones.Fields!SetiembreSisAcum / oRsAtenciones.Fields!SetiembreSIS, 2)
            End If
            If oRsAtenciones.Fields!SetiembreOtros > 0 Then
               oRsAtenciones.Fields!SetiembreOtrosAcum = Round(oRsAtenciones.Fields!SetiembreOtrosAcum / oRsAtenciones.Fields!SetiembreOtros, 2)
            End If
            If oRsAtenciones.Fields!OctubreSIS > 0 Then
               oRsAtenciones.Fields!OctubreSisAcum = Round(oRsAtenciones.Fields!OctubreSisAcum / oRsAtenciones.Fields!OctubreSIS, 2)
            End If
            If oRsAtenciones.Fields!OctubreOtros > 0 Then
               oRsAtenciones.Fields!OctubreOtrosAcum = Round(oRsAtenciones.Fields!OctubreOtrosAcum / oRsAtenciones.Fields!OctubreOtros, 2)
            End If
            If oRsAtenciones.Fields!NoviembreSIS > 0 Then
               oRsAtenciones.Fields!NoviembreSisAcum = Round(oRsAtenciones.Fields!NoviembreSisAcum / oRsAtenciones.Fields!NoviembreSIS, 2)
            End If
            If oRsAtenciones.Fields!NoviembreOtros > 0 Then
               oRsAtenciones.Fields!NoviembreOtrosAcum = Round(oRsAtenciones.Fields!NoviembreOtrosAcum / oRsAtenciones.Fields!NoviembreOtros, 2)
            End If
            If oRsAtenciones.Fields!DiciembreSIS > 0 Then
               oRsAtenciones.Fields!DiciembreSisAcum = Round(oRsAtenciones.Fields!DiciembreSisAcum / oRsAtenciones.Fields!DiciembreSIS, 2)
            End If
            If oRsAtenciones.Fields!DiciembreOtros > 0 Then
               oRsAtenciones.Fields!DiciembreOtrosAcum = Round(oRsAtenciones.Fields!DiciembreOtrosAcum / oRsAtenciones.Fields!DiciembreOtros, 2)
            End If
            If oRsAtenciones.Fields!TotalSIS > 0 Then
               oRsAtenciones.Fields!totalSisAcum = Round(oRsAtenciones.Fields!totalSisAcum / oRsAtenciones.Fields!TotalSIS, 2)
            End If
            If oRsAtenciones.Fields!TotalOtros > 0 Then
               oRsAtenciones.Fields!TotalOtrosAcum = Round(oRsAtenciones.Fields!TotalOtrosAcum / oRsAtenciones.Fields!TotalOtros, 2)
            End If
            oRsAtenciones.Update
        Else
           MsgBox "No hay datos", vbInformation, "Reporte"
           Exit Sub
        End If
    Case 4   'Nro vivos y muertos hosp
        Set oRsTmp1 = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, CDate("01/01/" & lcANIO), CDate("31/12/" & lcANIO), 3)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              If oRsTmp1.Fields!IdCondicionAlta = 4 Then
                 'Muerto
                 oRsAtenciones.MoveFirst
              Else
                 'Vivo
                 oRsAtenciones.MoveFirst
                 oRsAtenciones.MoveNext
              End If
              Select Case Month(oRsTmp1.Fields!FechaEgreso)
              Case 1
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!EneroSIS = oRsAtenciones.Fields!EneroSIS + 1
                   Else
                      oRsAtenciones.Fields!EneroOtros = oRsAtenciones.Fields!EneroOtros + 1
                   End If
              Case 2
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!FebreroSIS = oRsAtenciones.Fields!FebreroSIS + 1
                   Else
                      oRsAtenciones.Fields!FebreroOtros = oRsAtenciones.Fields!FebreroOtros + 1
                   End If
              Case 3
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MarzoSIS = oRsAtenciones.Fields!MarzoSIS + 1
                   Else
                      oRsAtenciones.Fields!MarzoOtros = oRsAtenciones.Fields!MarzoOtros + 1
                   End If
              Case 4
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AbrilSIS = oRsAtenciones.Fields!AbrilSIS + 1
                   Else
                      oRsAtenciones.Fields!AbrilOtros = oRsAtenciones.Fields!AbrilOtros + 1
                   End If
              Case 5
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!MayoSIS = oRsAtenciones.Fields!MayoSIS + 1
                   Else
                      oRsAtenciones.Fields!MayoOtros = oRsAtenciones.Fields!MayoOtros + 1
                   End If
              Case 6
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JunioSIS = oRsAtenciones.Fields!JunioSIS + 1
                   Else
                      oRsAtenciones.Fields!JunioOtros = oRsAtenciones.Fields!JunioOtros + 1
                   End If
              Case 7
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!JulioSIS = oRsAtenciones.Fields!JulioSIS + 1
                   Else
                      oRsAtenciones.Fields!JulioOtros = oRsAtenciones.Fields!JulioOtros + 1
                   End If
              Case 8
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!AgostoSIS = oRsAtenciones.Fields!AgostoSIS + 1
                   Else
                      oRsAtenciones.Fields!AgostoOtros = oRsAtenciones.Fields!AgostoOtros + 1
                   End If
              Case 9
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!SetiembreSIS = oRsAtenciones.Fields!SetiembreSIS + 1
                   Else
                      oRsAtenciones.Fields!SetiembreOtros = oRsAtenciones.Fields!SetiembreOtros + 1
                   End If
              Case 10
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!OctubreSIS = oRsAtenciones.Fields!OctubreSIS + 1
                   Else
                      oRsAtenciones.Fields!OctubreOtros = oRsAtenciones.Fields!OctubreOtros + 1
                   End If
              Case 11
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!NoviembreSIS = oRsAtenciones.Fields!NoviembreSIS + 1
                   Else
                      oRsAtenciones.Fields!NoviembreOtros = oRsAtenciones.Fields!NoviembreOtros + 1
                   End If
              Case 12
                   If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                      oRsAtenciones.Fields!DiciembreSIS = oRsAtenciones.Fields!DiciembreSIS + 1
                   Else
                      oRsAtenciones.Fields!DiciembreOtros = oRsAtenciones.Fields!DiciembreOtros + 1
                   End If
              End Select
              If oRsTmp1.Fields!idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                 oRsAtenciones.Fields!TotalSIS = oRsAtenciones.Fields!TotalSIS + 1
              Else
                 oRsAtenciones.Fields!TotalOtros = oRsAtenciones.Fields!TotalOtros + 1
              End If
              oRsAtenciones.Update
              oRsTmp1.MoveNext
           Loop
        Else
           MsgBox "No hay datos", vbInformation, "Reporte"
           Exit Sub
        End If
    End Select
    'Reporte
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\AtencionesTotales.ods"
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
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
        
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\AtencionesTotales.xls")
        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
        If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(0, 1).setFormula(lcTitulo)
        Call Feuille.getcellbyposition(0, 3).setFormula("(AÑO: " & lcANIO & ")")
    Else
        oWorkSheet.Cells(2, 1).Value = lcTitulo
        oWorkSheet.Cells(4, 1).Value = "(AÑO: " & lcANIO & ")"
    End If
    iFila = 8
    If ml_opcion = 4 Then
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Facellidos")
            Call Feuille.getcellbyposition(0, iFila + 0).setFormula("Vivos")
            Call Feuille.getcellbyposition(0, iFila + 1).setFormula("")
        Else
            oWorkSheet.Cells(iFila, 1).Value = "Facellidos"
            oWorkSheet.Cells(iFila + 1, 1).Value = "Vivos"
            oWorkSheet.Cells(iFila + 2, 1).Value = ""
        End If
    End If
    oRsAtenciones.MoveFirst
    Do While Not oRsAtenciones.EOF
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!EneroSIS, oRsAtenciones.Fields!EneroSisAcum))
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!EneroOtros, oRsAtenciones.Fields!EneroOtrosAcum))
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!FebreroSIS, oRsAtenciones.Fields!FebreroSisAcum))
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!FebreroOtros, oRsAtenciones.Fields!FebreroOtrosAcum))
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MarzoSIS, oRsAtenciones.Fields!MarzoSisAcum))
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MarzoOtros, oRsAtenciones.Fields!MarzoOtrosAcum))
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AbrilSIS, oRsAtenciones.Fields!AbrilSisAcum))
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AbrilOtros, oRsAtenciones.Fields!AbrilOtrosAcum))
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MayoSIS, oRsAtenciones.Fields!MayoSisAcum))
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MayoOtros, oRsAtenciones.Fields!MayoOtrosAcum))
            Call Feuille.getcellbyposition(11, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JunioSIS, oRsAtenciones.Fields!JunioSisAcum))
            Call Feuille.getcellbyposition(12, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JunioOtros, oRsAtenciones.Fields!JunioOtrosAcum))
            Call Feuille.getcellbyposition(13, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JulioSIS, oRsAtenciones.Fields!JulioSisAcum))
            Call Feuille.getcellbyposition(14, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JulioOtros, oRsAtenciones.Fields!JulioOtrosAcum))
            Call Feuille.getcellbyposition(15, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AgostoSIS, oRsAtenciones.Fields!AgostoSisAcum))
            Call Feuille.getcellbyposition(16, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AgostoOtros, oRsAtenciones.Fields!AgostoOtrosAcum))
            Call Feuille.getcellbyposition(17, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!SetiembreSIS, oRsAtenciones.Fields!SetiembreSisAcum))
            Call Feuille.getcellbyposition(18, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!SetiembreOtros, oRsAtenciones.Fields!SetiembreOtrosAcum))
            Call Feuille.getcellbyposition(19, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!OctubreSIS, oRsAtenciones.Fields!OctubreSisAcum))
            Call Feuille.getcellbyposition(20, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!OctubreOtros, oRsAtenciones.Fields!OctubreOtrosAcum))
            Call Feuille.getcellbyposition(21, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!NoviembreSIS, oRsAtenciones.Fields!NoviembreSisAcum))
            Call Feuille.getcellbyposition(22, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!NoviembreOtros, oRsAtenciones.Fields!NoviembreOtrosAcum))
            Call Feuille.getcellbyposition(23, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!DiciembreSIS, oRsAtenciones.Fields!DiciembreSisAcum))
            Call Feuille.getcellbyposition(24, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!DiciembreOtros, oRsAtenciones.Fields!DiciembreOtrosAcum))
            Call Feuille.getcellbyposition(25, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!TotalSIS, oRsAtenciones.Fields!totalSisAcum))
            Call Feuille.getcellbyposition(26, iFila - 1).setFormula(IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!TotalOtros, oRsAtenciones.Fields!TotalOtrosAcum))
        Else
            oWorkSheet.Cells(iFila, 2).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!EneroSIS, oRsAtenciones.Fields!EneroSisAcum)
            oWorkSheet.Cells(iFila, 3).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!EneroOtros, oRsAtenciones.Fields!EneroOtrosAcum)
            oWorkSheet.Cells(iFila, 4).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!FebreroSIS, oRsAtenciones.Fields!FebreroSisAcum)
            oWorkSheet.Cells(iFila, 5).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!FebreroOtros, oRsAtenciones.Fields!FebreroOtrosAcum)
            oWorkSheet.Cells(iFila, 6).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MarzoSIS, oRsAtenciones.Fields!MarzoSisAcum)
            oWorkSheet.Cells(iFila, 7).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MarzoOtros, oRsAtenciones.Fields!MarzoOtrosAcum)
            oWorkSheet.Cells(iFila, 8).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AbrilSIS, oRsAtenciones.Fields!AbrilSisAcum)
            oWorkSheet.Cells(iFila, 9).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AbrilOtros, oRsAtenciones.Fields!AbrilOtrosAcum)
            oWorkSheet.Cells(iFila, 10).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MayoSIS, oRsAtenciones.Fields!MayoSisAcum)
            oWorkSheet.Cells(iFila, 11).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!MayoOtros, oRsAtenciones.Fields!MayoOtrosAcum)
            oWorkSheet.Cells(iFila, 12).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JunioSIS, oRsAtenciones.Fields!JunioSisAcum)
            oWorkSheet.Cells(iFila, 13).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JunioOtros, oRsAtenciones.Fields!JunioOtrosAcum)
            oWorkSheet.Cells(iFila, 14).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JulioSIS, oRsAtenciones.Fields!JulioSisAcum)
            oWorkSheet.Cells(iFila, 15).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!JulioOtros, oRsAtenciones.Fields!JulioOtrosAcum)
            oWorkSheet.Cells(iFila, 16).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AgostoSIS, oRsAtenciones.Fields!AgostoSisAcum)
            oWorkSheet.Cells(iFila, 17).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!AgostoOtros, oRsAtenciones.Fields!AgostoOtrosAcum)
            oWorkSheet.Cells(iFila, 18).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!SetiembreSIS, oRsAtenciones.Fields!SetiembreSisAcum)
            oWorkSheet.Cells(iFila, 19).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!SetiembreOtros, oRsAtenciones.Fields!SetiembreOtrosAcum)
            oWorkSheet.Cells(iFila, 20).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!OctubreSIS, oRsAtenciones.Fields!OctubreSisAcum)
            oWorkSheet.Cells(iFila, 21).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!OctubreOtros, oRsAtenciones.Fields!OctubreOtrosAcum)
            oWorkSheet.Cells(iFila, 22).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!NoviembreSIS, oRsAtenciones.Fields!NoviembreSisAcum)
            oWorkSheet.Cells(iFila, 23).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!NoviembreOtros, oRsAtenciones.Fields!NoviembreOtrosAcum)
            oWorkSheet.Cells(iFila, 24).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!DiciembreSIS, oRsAtenciones.Fields!DiciembreSisAcum)
            oWorkSheet.Cells(iFila, 25).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!DiciembreOtros, oRsAtenciones.Fields!DiciembreOtrosAcum)
            oWorkSheet.Cells(iFila, 26).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!TotalSIS, oRsAtenciones.Fields!totalSisAcum)
            oWorkSheet.Cells(iFila, 27).Value = IIf(ml_opcion = 1 Or ml_opcion = 4, oRsAtenciones.Fields!TotalOtros, oRsAtenciones.Fields!TotalOtrosAcum)
        End If
       iFila = iFila + 1
       oRsAtenciones.MoveNext
    Loop
    If lbEsOpenOffice = True Then
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
    Else
        oExcel.Visible = True
        'oWorkSheet.PrintPreview
    End If
    If lbEsOpenOffice = True Then
        'Liberar Memoria
        Set Plage = Nothing
        Set Feuille = Nothing
        Set Document = Nothing
        Set Desktop = Nothing
        Set ServiceManager = Nothing
        Set Style = Nothing
        Set Border = Nothing
        Set PageStyles = Nothing
        Set Sheet = Nothing
        Set StyleFamilies = Nothing
        Set DefPage = Nothing
        Set Htext = Nothing
        Set Hcontent = Nothing
    Else
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If

End Sub



Sub TiempoEnAtenciones(lcANIO As String, ml_opcion As Integer, ml_opcionSub As Integer, lcTitulo As String, _
                       ldFechaInicial As Date, ldFechaFinal As Date)
    Dim iFila As Long, lnDiasEstancia As Double, lnMesX As Integer
    Dim lnEstancia As Integer, lnConsumoServicio As Double, lnConsumoFarmacia As Double
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oBuscaDiasPaciente As New SIGHDatos.Parametros
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oRsAtenciones As New Recordset
    Dim oRsAtencionesCE As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lcHoraEstanciaMax As String
    Dim lbSegundaVezEne As Boolean, lbSegundaVezFeb As Boolean, lbSegundaVezMar As Boolean, lbSegundaVezAbr As Boolean
    Dim lbSegundaVezMay  As Boolean, lbSegundaVezJun As Boolean, lbSegundaVezJul As Boolean, lbSegundaVezAgo As Boolean
    Dim lbSegundaVezSet  As Boolean, lbSegundaVezOct As Boolean, lbSegundaVezNov As Boolean, lbSegundaVezDic As Boolean
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    Dim lcSql As String, lcHoraEgresoUlt As String, lcConsultorio As String, ldFecha As Date
    
    Dim lbEsOpenOffice As Boolean
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
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
        Dim con As Long
        con = 0
        
        ldFechaFinal = CDate(Format(ldFechaFinal, sighentidades.DevuelveFechaSoloFormato_DMY & " 23:59:59"))
        Set oRsAtencionesCE = mo_ReglasAdmision.AtencionesTiempoCE(lcANIO, ldFechaInicial, ldFechaFinal)
        
        
        
        If oRsAtencionesCE.RecordCount > 0 Then
           oRsAtencionesCE.MoveFirst
           
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\TiempoAtencionCE.ods"
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
                mo_CabeceraReportes.CabeceraReportes Document, True
            Else

            
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\TiempoAtencionCE.xls")
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                oWorkBook.Sheets(1).Name = "ENERO"
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False


            End If
    
           Do While Not oRsAtencionesCE.EOF
              con = con + 1

              lnMesX = Month(oRsAtencionesCE.Fields!FechaIngreso)
              Select Case lnMesX
              Case 1
                  If lbSegundaVezEne = False Then
                     mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                     lbSegundaVezEne = True
                  End If
              Case 2
                  If lbSegundaVezFeb = False Then
                      If lbEsOpenOffice = True Then
                          Set Feuille = Document.getSheets().getByIndex(1)
                      Else

                          Set oWorkSheet = oWorkBook.Sheets(2)
                          oWorkBook.Sheets(2).Name = "FEBRERO"
                      End If
                      mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                      lbSegundaVezFeb = True
                  End If
              Case 3
                  If lbSegundaVezMar = False Then
                      If lbEsOpenOffice = True Then
                          Set Feuille = Document.getSheets().getByIndex(2)
                          
                      Else
                          Set oWorkSheet = oWorkBook.Sheets(3)
                          oWorkBook.Sheets(3).Name = "Marzo"
                          
                      End If
                      mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                      lbSegundaVezMar = True
                  End If
              Case 4
                  If lbSegundaVezAbr = False Then
                      If lbEsOpenOffice = True Then
                          Set Feuille = Document.getSheets().getByIndex(3)
                          
                      Else
                          Set oWorkSheet = oWorkBook.Sheets(4)
                          oWorkBook.Sheets(4).Name = "Abril"
                          
                      End If
                      mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                      lbSegundaVezAbr = True
                  End If
              Case 5
                  If lbSegundaVezMay = False Then
                      If lbEsOpenOffice = True Then
                          Set Feuille = Document.getSheets().getByIndex(4)
                          
                      Else
                          Set oWorkSheet = oWorkBook.Sheets(5)
                          oWorkBook.Sheets(5).Name = "Mayo"
                          
                      End If
                      mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                      lbSegundaVezMay = True
                  End If
              Case 6
                  If lbSegundaVezJun = False Then
                      If lbEsOpenOffice = True Then
                          Set Feuille = Document.getSheets().getByIndex(5)
                          
                      Else
                          Set oWorkSheet = oWorkBook.Sheets(6)
                          oWorkBook.Sheets(6).Name = "Junio"
                          
                      End If
                      lbSegundaVezJun = True
                      mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                  End If
             Case 7
                    If lbSegundaVezJul = False Then
                        If lbEsOpenOffice = True Then
                            Set Feuille = Document.getSheets().getByIndex(6)
                            
                        Else
                            Set oWorkSheet = oWorkBook.Sheets(7)
                            oWorkBook.Sheets(7).Name = "Julio"
                            
                        End If
                        lbSegundaVezJul = True
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                    End If
             Case 8
                    If lbSegundaVezAgo = False Then
                        If lbEsOpenOffice = True Then
                            Set Feuille = Document.getSheets().getByIndex(7)
                            
                        Else
                            Set oWorkSheet = oWorkBook.Sheets(8)
                            oWorkBook.Sheets(8).Name = "Agosto"
                            
                        End If
                        lbSegundaVezAgo = True
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                    End If
             Case 9
                    If lbSegundaVezSet = False Then
                        If lbEsOpenOffice = True Then
                            Set Feuille = Document.getSheets().getByIndex(8)
                             
                        Else
                            Set oWorkSheet = oWorkBook.Sheets(9)
                            oWorkBook.Sheets(9).Name = "Setiembre"
                            
                        End If
                        lbSegundaVezSet = True
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                    End If
            Case 10
                   If lbSegundaVezOct = False Then
                       If lbEsOpenOffice = True Then
                           Set Feuille = Document.getSheets().getByIndex(9)
                            
                       Else
                           Set oWorkSheet = oWorkBook.Sheets(10)
                           oWorkBook.Sheets(10).Name = "Octubre"
                           
                       End If
                       lbSegundaVezOct = True
                       mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                   End If
             Case 11
                   If lbSegundaVezNov = False Then
                       If lbEsOpenOffice = True Then
                           Set Feuille = Document.getSheets().getByIndex(10)
                            
                       Else
                           Set oWorkSheet = oWorkBook.Sheets(11)
                           oWorkBook.Sheets(11).Name = "Noviembre"
                           
                       End If
                       mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                       lbSegundaVezNov = True
                   End If
             Case 12
                    If lbSegundaVezDic = False Then
                        If lbEsOpenOffice = True Then
                            Set Feuille = Document.getSheets().getByIndex(11)
                            
                        Else
                            Set oWorkSheet = oWorkBook.Sheets(12)
                            oWorkBook.Sheets(12).Name = "Diciembre"
                            
                        End If
                        lbSegundaVezDic = True
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, lbEsOpenOffice
                    End If
              End Select
              If lbEsOpenOffice = True Then
                  'Call Feuille.getcellbyposition(0, 1).setFormula(lcTitulo)
                  Call Feuille.getcellbyposition(0, 3).setFormula("(AÑO: " & lcANIO & ") (Ordenado por columna: 'Hora fin')")
              Else
                  'oWorkSheet.Cells(2, 1).Value = lcTitulo
                  If Val(lcANIO) > 0 Then
                     oWorkSheet.Cells(4, 1).Value = "(AÑO: " & lcANIO & ") (Ordenado por columna: 'Hora fin')"
                  Else
                     oWorkSheet.Cells(4, 1).Value = "(Fechas: " & Format(ldFechaInicial, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                                    " al " & Format(ldFechaFinal, sighentidades.DevuelveFechaSoloFormato_DMY) & _
                                                    ") (Ordenado por columna: 'Hora fin')"
                  End If
              End If
              iFila = 8
              Do While Not oRsAtencionesCE.EOF And Month(oRsAtencionesCE.Fields!FechaIngreso) = lnMesX
                  lcHoraEgresoUlt = IIf(IsNull(oRsAtencionesCE!horaInicioAtencion), "", oRsAtencionesCE!horaInicioAtencion)
                  lcConsultorio = oRsAtencionesCE.Fields!Consultorio
                  ldFecha = oRsAtencionesCE.Fields!FechaIngreso
                  lcSql = ""
                  Do While Not oRsAtencionesCE.EOF And lcConsultorio = oRsAtencionesCE.Fields!Consultorio And ldFecha = oRsAtencionesCE.Fields!FechaIngreso
                      If lbEsOpenOffice = True Then
                          Call Feuille.getcellbyposition(8, iFila - 1).setFormula(oRsAtencionesCE.Fields!Consultorio)
                          Call Feuille.getcellbyposition(6, iFila - 1).setFormula(oRsAtencionesCE.Fields!dpto)
                          Call Feuille.getcellbyposition(7, iFila - 1).setFormula(oRsAtencionesCE.Fields!especialidad)
                          Call Feuille.getcellbyposition(1, iFila - 1).setFormula(Day(oRsAtencionesCE.Fields!FechaIngreso))
                          Call Feuille.getcellbyposition(2, iFila - 1).setFormula(oRsAtencionesCE.Fields!HoraIngreso)
                      Else
                          oWorkSheet.Cells(iFila, 9).Value = oRsAtencionesCE.Fields!Consultorio
                          oWorkSheet.Cells(iFila, 7).Value = oRsAtencionesCE.Fields!dpto
                          oWorkSheet.Cells(iFila, 8).Value = oRsAtencionesCE.Fields!especialidad
                          oWorkSheet.Cells(iFila, 2).Value = Day(oRsAtencionesCE.Fields!FechaIngreso)
                          oWorkSheet.Cells(iFila, 3).Value = oRsAtencionesCE.Fields!HoraIngreso
                      End If
                      If lcHoraEgresoUlt <> "" And lcHoraEgresoUlt <> sighentidades.HORA_VACIA_HM Then
                          If lbEsOpenOffice = True Then
                              Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcHoraEgresoUlt)
                              Call Feuille.getcellbyposition(4, iFila - 1).setFormula("'" & oRsAtencionesCE.Fields!horaEgreso)
                              Call Feuille.getcellbyposition(5, iFila - 1).setFormula(DateDiff("n", lcHoraEgresoUlt, oRsAtencionesCE.Fields!horaEgreso))
                              
                          Else
                              oWorkSheet.Cells(iFila, 4).Value = "'" & lcHoraEgresoUlt
                              oWorkSheet.Cells(iFila, 5).Value = oRsAtencionesCE.Fields!horaEgreso
                              oWorkSheet.Cells(iFila, 6).Value = DateDiff("n", lcHoraEgresoUlt, oRsAtencionesCE.Fields!horaEgreso)
                          End If
                      Else
                          If lbEsOpenOffice = True Then
                              Call Feuille.getcellbyposition(3, iFila - 1).setFormula(oRsAtencionesCE.Fields!HoraIngreso)
                              Call Feuille.getcellbyposition(4, iFila - 1).setFormula(oRsAtencionesCE.Fields!horaEgreso)
                              If lcSql = "" Then
                                 Call Feuille.getcellbyposition(5, iFila - 1).setFormula("<>")
                              End If
                          Else
                              oWorkSheet.Cells(iFila, 4).Value = oRsAtencionesCE.Fields!HoraIngreso
                              oWorkSheet.Cells(iFila, 5).Value = oRsAtencionesCE.Fields!horaEgreso
                              If lcSql = "" Then
                               '  oWorkSheet.Cells(iFila, 6).Value = "<>"
                              End If
                          End If
                          lcSql = "*"
                      End If
                      If IsNull(oRsAtencionesCE.Fields!horaEgreso) Or oRsAtencionesCE.Fields!horaEgreso = "" Then
                      Else
                         lcHoraEgresoUlt = IIf(IsNull(oRsAtencionesCE!horaInicioAtencion), oRsAtencionesCE!horaEgreso, oRsAtencionesCE!horaInicioAtencion)
                      End If
                      If oRsAtencionesCE.Fields!FechaIngreso = oRsAtencionesCE.Fields!FechaCreacion Then
                          If lbEsOpenOffice = True Then
                              Call Feuille.getcellbyposition(9, iFila - 1).setFormula("Historia Nueva")
                          Else
                              oWorkSheet.Cells(iFila, 10).Value = "Historia Nueva"
                          End If
                      End If
                      If lbEsOpenOffice = True Then
                              Call Feuille.getcellbyposition(10, iFila - 1).setFormula(IIf(IsNull(oRsAtencionesCE.Fields!etnias), "", oRsAtencionesCE.Fields!etnias))
                              Call Feuille.getcellbyposition(11, iFila - 1).setFormula(IIf(IsNull(oRsAtencionesCE.Fields!lengua), "", oRsAtencionesCE.Fields!lengua))
                      Else
                              oWorkSheet.Cells(iFila, 11).Value = IIf(IsNull(oRsAtencionesCE.Fields!etnias), "", oRsAtencionesCE.Fields!etnias)
                              oWorkSheet.Cells(iFila, 12).Value = IIf(IsNull(oRsAtencionesCE.Fields!lengua), "", oRsAtencionesCE.Fields!lengua)
                      End If
                      Set oRsTmp2 = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(oRsAtencionesCE.Fields!idAtencion, sghRegistroCitaCE, oConexion)
                      If oRsTmp2.RecordCount > 0 Then
                          If lbEsOpenOffice = True Then
                              Call Feuille.getcellbyposition(12, iFila - 1).setFormula(IIf(IsNull(oRsTmp2.Fields!nombrePc), "", oRsTmp2.Fields!nombrePc))
                              Call Feuille.getcellbyposition(13, iFila - 1).setFormula(IIf(IsNull(oRsTmp2.Fields!Usuario), "", oRsTmp2.Fields!Usuario))
                          Else
                              oWorkSheet.Cells(iFila, 13).Value = IIf(IsNull(oRsTmp2.Fields!nombrePc), "", oRsTmp2.Fields!nombrePc)
                              oWorkSheet.Cells(iFila, 14).Value = IIf(IsNull(oRsTmp2.Fields!Usuario), "", oRsTmp2.Fields!Usuario)
                          End If
                      End If
                      oRsTmp2.Close
                      iFila = iFila + 1
                      oRsAtencionesCE.MoveNext
                      If oRsAtencionesCE.EOF Then
                          Exit Do
                      End If
                  Loop
                  iFila = iFila + 1
                  If oRsAtencionesCE.EOF Then
                      Exit Do
                  End If
             Loop
             If oRsAtencionesCE.EOF Then
                 Exit Do
             End If
        Loop
        If lbEsOpenOffice = True Then
           Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
           MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
           oExcel.Visible = True
           oWorkSheet.PrintPreview
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
           Set oWorkSheet = Nothing
           Set oExcel = Nothing
           Set oConexion = Nothing
        End If
        Exit Sub
        Else
           MsgBox "No hay datos", vbInformation, "Reporte"
           Set oConexion = Nothing
           Exit Sub
        End If

End Sub
Sub AtencionesXsexoEdad(lnAnio As Integer, lnTrimestre As Integer, lnIdFuenteFinanciamiento As Integer, _
                        lnTipoServicio As Integer, lcTitulo As String, ml_TextoDelFiltro As String, _
                        lbEnExcel As Boolean, lnHwnd As Long)
    Dim oRsTmp1 As New Recordset
    Dim rsReporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim oConexion As New ADODB.Connection
    Dim lnFor As Integer, lnTotFemenino As Long, lnTotMasculino As Long, lnTotal As Long, lnPorcentaje1 As Double
    Dim lnTotalF As Long, lntotalM As Long, lnPorcentaje2 As Double

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With rsReporte1
                .Fields.Append "Edad", adVarChar, 50
                .Fields.Append "Femenino", adInteger
                .Fields.Append "Masculino", adInteger
                .Fields.Append "TotalGeneral", adInteger
                .Fields.Append "Porcentaje", adDouble
                .Fields.Append "Acumulado", adDouble
                .LockType = adLockOptimistic
                .Open
    End With
    With rsReporte
                .Fields.Append "Edad", adVarChar, 50
                .Fields.Append "EdadInicio", adInteger
                .Fields.Append "EdadFinal", adInteger
                .Fields.Append "Femenino", adInteger
                .Fields.Append "Masculino", adInteger
                .Fields.Append "TotalGeneral", adInteger
                .Fields.Append "Porcentaje", adDouble
                .Fields.Append "Acumulado", adDouble
                
                .LockType = adLockOptimistic
                .Open
                .AddNew
                .Fields!Edad = "0 a 4"
                .Fields!EdadInicio = 0
                .Fields!EdadFinal = 4
                .Update
                .AddNew
                .Fields!Edad = "5 a 9"
                .Fields!EdadInicio = 5
                .Fields!EdadFinal = 9
                .Update
                .AddNew
                .Fields!Edad = "10 a 14"
                .Fields!EdadInicio = 10
                .Fields!EdadFinal = 14
                .Update
                .AddNew
                .Fields!Edad = "15 a 19"
                .Fields!EdadInicio = 15
                .Fields!EdadFinal = 19
                .Update
                .AddNew
                .Fields!Edad = "20 a 24"
                .Fields!EdadInicio = 20
                .Fields!EdadFinal = 24
                .Update
                .AddNew
                .Fields!Edad = "25 a 29"
                .Fields!EdadInicio = 25
                .Fields!EdadFinal = 29
                .Update
                .AddNew
                .Fields!Edad = "30 a 34"
                .Fields!EdadInicio = 30
                .Fields!EdadFinal = 34
                .Update
                .AddNew
                .Fields!Edad = "35 a 39"
                .Fields!EdadInicio = 35
                .Fields!EdadFinal = 39
                .Update
                .AddNew
                .Fields!Edad = "40 a 44"
                .Fields!EdadInicio = 40
                .Fields!EdadFinal = 44
                .Update
                .AddNew
                .Fields!Edad = "45 a 49"
                .Fields!EdadInicio = 45
                .Fields!EdadFinal = 49
                .Update
                .AddNew
                .Fields!Edad = "50 a 54"
                .Fields!EdadInicio = 50
                .Fields!EdadFinal = 54
                .Update
                .AddNew
                .Fields!Edad = "55 a 59"
                .Fields!EdadInicio = 55
                .Fields!EdadFinal = 59
                .Update
                .AddNew
                .Fields!Edad = "60 a 64"
                .Fields!EdadInicio = 60
                .Fields!EdadFinal = 64
                .Update
                .AddNew
                .Fields!Edad = "65 a 69"
                .Fields!EdadInicio = 65
                .Fields!EdadFinal = 69
                .Update
                .AddNew
                .Fields!Edad = "70 a 74"
                .Fields!EdadInicio = 70
                .Fields!EdadFinal = 74
                .Update
                .AddNew
                .Fields!Edad = "75 a 79"
                .Fields!EdadInicio = 75
                .Fields!EdadFinal = 79
                .Update
                .AddNew
                .Fields!Edad = "80 a 130"
                .Fields!EdadInicio = 80
                .Fields!EdadFinal = 130
                .Update
    End With
    'totales femeninos y masculinos
    rsReporte.MoveFirst
    lnTotFemenino = 0: lnTotMasculino = 0
    Do While Not rsReporte.EOF
        '***masculino
        Set oRsTmp1 = AtencionesXtrimestreSexoRangoEdadAnios(lnAnio, lnTrimestre, rsReporte.Fields!EdadInicio, _
                                                             rsReporte.Fields!EdadFinal, 1, lnIdFuenteFinanciamiento, _
                                                             lnTipoServicio, oConexion)

        lntotalM = oRsTmp1.RecordCount
        If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Then
            Set oRsTmp1 = AtencionesXtrimestreSexoRangoEdadAnios(lnAnio, lnTrimestre, rsReporte.Fields!EdadInicio, _
                                                             rsReporte.Fields!EdadFinal, 1, sghFuenteFinanciamiento.sghFFParticularHospitalizado, _
                                                             lnTipoServicio, oConexion)
            lntotalM = lntotalM + oRsTmp1.RecordCount
        End If
        '***femenino
        Set oRsTmp1 = AtencionesXtrimestreSexoRangoEdadAnios(lnAnio, lnTrimestre, rsReporte.Fields!EdadInicio, _
                                                             rsReporte.Fields!EdadFinal, 2, lnIdFuenteFinanciamiento, _
                                                             lnTipoServicio, oConexion)
        lnTotalF = oRsTmp1.RecordCount
        If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente Then
            Set oRsTmp1 = AtencionesXtrimestreSexoRangoEdadAnios(lnAnio, lnTrimestre, rsReporte.Fields!EdadInicio, _
                                                             rsReporte.Fields!EdadFinal, 2, sghFuenteFinanciamiento.sghFFParticularHospitalizado, _
                                                             lnTipoServicio, oConexion)
            lnTotalF = lnTotalF + oRsTmp1.RecordCount
        End If
        '
        lnTotMasculino = lnTotMasculino + lntotalM
        lnTotFemenino = lnTotFemenino + lnTotalF
        rsReporte.Fields!Femenino = lnTotalF
        rsReporte.Fields!Masculino = lntotalM
        rsReporte.Fields!TotalGeneral = lnTotalF + lntotalM
        rsReporte.Update
        rsReporte.MoveNext
    Loop
    'porcentajes
    lnTotal = lnTotFemenino + lnTotMasculino
    If lnTotal > 0 Then
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
           
            lnPorcentaje1 = Round(rsReporte.Fields!TotalGeneral * 100 / lnTotal, 1)
            If rsReporte!EdadInicio = 0 Then
               lnPorcentaje2 = lnPorcentaje1
            Else
               lnPorcentaje2 = lnPorcentaje1 + lnPorcentaje2
            End If
            rsReporte.Fields!Porcentaje = lnPorcentaje1
            rsReporte.Fields!Acumulado = IIf(lnPorcentaje2 > 100, 100, lnPorcentaje2)
            rsReporte.Update
            rsReporte.MoveNext
        Loop
    End If
    'total general
    rsReporte.AddNew
    rsReporte.Fields!Edad = "Total General"
    rsReporte.Fields!EdadInicio = 0
    rsReporte.Fields!EdadFinal = 130
    rsReporte.Fields!Femenino = lnTotFemenino
    rsReporte.Fields!Masculino = lnTotMasculino
    rsReporte.Fields!TotalGeneral = lnTotal
    rsReporte.Update
    'pasar lo necesario p el reporte
    rsReporte.MoveFirst
    Do While Not rsReporte.EOF
        rsReporte1.AddNew
        rsReporte1.Fields!Edad = rsReporte!Edad
        rsReporte1.Fields!Femenino = rsReporte!Femenino
        rsReporte1.Fields!Masculino = rsReporte!Masculino
        rsReporte1.Fields!TotalGeneral = rsReporte!TotalGeneral
        rsReporte1.Fields!Porcentaje = rsReporte!Porcentaje
        rsReporte1.Fields!Acumulado = rsReporte!Acumulado
        rsReporte1.Update
        rsReporte.MoveNext
    Loop
    '
    If lbEnExcel = True Then
       mo_ReglasReportes.ExportarRecordSetAexcel rsReporte1, lcTitulo, ml_TextoDelFiltro, "", lnHwnd
    End If
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set rsReporte = Nothing
    Set mo_ReglasReportes = Nothing
    Set oConexion = Nothing
End Sub



Function AtencionesXtrimestreSexoRangoEdadAnios(lnAnio As Integer, lnTrimestre As Integer, lnEdadAnio1 As Integer, _
                                                 lnEdadAnio2 As Integer, lnIdTipoSexo As Integer, _
                                                 lnIdFuenteFinanciamiento As Integer, lnIdTipoServicio As Integer, _
                                                 oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError1
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesXtrimestreSexoRangoEdadAnios"
        Set oParameter = .CreateParameter("@anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Trimestre", adInteger, adParamInput, 0, lnTrimestre + 1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@edadAnio1", adInteger, adParamInput, 0, lnEdadAnio1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@edadAnio2", adInteger, adParamInput, 0, lnEdadAnio2): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoSexo", adInteger, adParamInput, 0, lnIdTipoSexo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesXtrimestreSexoRangoEdadAnios = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError1:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub AtencionesMorbilidad(lnAnio As Integer, lnTrimestre As Integer, lnIdFuenteFinanciamiento As Integer, _
                        lnTipoServicio As Integer, lcTitulo As String, ml_TextoDelFiltro As String, _
                        lbEnExcel As Boolean, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim rsReporte  As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim lnTotDNro As Long, ldDiagnostico As String, lnDiagnostico As Long, lcCie10 As String
    Dim mda_FechaInicio As Date, mda_FechaFin As Date, lnTotal As Long, lbNuevo As Boolean
    Dim lnPorcentaje1 As Double, lnPorcentaje2 As Double
    With mrs_Tmp
          .Fields.Append "Cie10", adVarChar, 10, adFldIsNullable
          .Fields.Append "Diagnostico", adVarChar, 50, adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Porcentaje", adDouble
          .Fields.Append "Acumulado", adDouble
          .LockType = adLockOptimistic
          .Open
    End With
    mda_FechaInicio = CDate("01/" & IIf(lnTrimestre = 0, "01/", IIf(lnTrimestre = 1, "04/", IIf(lnTrimestre = 2, "07/", "10/"))) & Trim(Str(lnAnio)))
    mda_FechaFin = CDate(IIf(lnTrimestre = 0, "31/03/", IIf(lnTrimestre = 1, "30/06/", IIf(lnTrimestre = 2, "30/09/", "31/12/"))) & Trim(Str(lnAnio)))
    lnTotal = 0
    '**** consulta externa
    If lnTipoServicio = 0 Or lnTipoServicio = sghTipoServicio.sghConsultaExterna Then
        Set rsReporte = AtencionesDiagnosticosSeleccionarCEatendidosPorFechasSexoTipodx(0, 0, mda_FechaInicio, mda_FechaFin)
        If lnIdFuenteFinanciamiento > 0 Then
           If lnIdFuenteFinanciamiento = 1 Then
               rsReporte.Filter = "idFuenteFinanciamiento=1 or idFuenteFinanciamiento=5"
           Else
               rsReporte.Filter = "idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento
           End If
        End If
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTotDNro = 0
                ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value
                lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                If Not IsNull(rsReporte!cie10) Then
                  lcCie10 = rsReporte.Fields("Cie10").Value
                Else
                  lcCie10 = ""
                End If
                Do While Not rsReporte.EOF And ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value And lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                    lnTotDNro = lnTotDNro + 1
                    lnTotal = lnTotal + 1
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                mrs_Tmp.AddNew
                mrs_Tmp.Fields("Cie10").Value = lcCie10
                mrs_Tmp.Fields("Diagnostico").Value = Left(ldDiagnostico, 50)
                mrs_Tmp.Fields("Cantidad").Value = lnTotDNro
                mrs_Tmp.Update
            Loop
        End If
    End If
    '***** hospitalizacion
    If lnTipoServicio = 0 Or lnTipoServicio = sghTipoServicio.sghHospitalizacion Then
        Set rsReporte = AtencionesDiagnosticosSeleccionarHospitalizadosEgresadosPorFechasSexoTipodx(0, 301, _
                                                                                mda_FechaInicio, mda_FechaFin)
        If lnIdFuenteFinanciamiento > 0 Then
           If lnIdFuenteFinanciamiento = 1 Then
               rsReporte.Filter = "idFuenteFinanciamiento=1 or idFuenteFinanciamiento=5"
           Else
               rsReporte.Filter = "idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento
           End If
        End If
        If rsReporte.RecordCount > 0 Then
            rsReporte.MoveFirst
            Do While Not rsReporte.EOF
                lnTotDNro = 0
                ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value
                lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                If Not IsNull(rsReporte!cie10) Then
                  lcCie10 = rsReporte.Fields("Cie10").Value
                Else
                  lcCie10 = ""
                End If
                Do While Not rsReporte.EOF And ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value And lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                    lnTotDNro = lnTotDNro + 1
                    lnTotal = lnTotal + 1
                    rsReporte.MoveNext
                    If rsReporte.EOF Then
                       Exit Do
                    End If
                Loop
                lbNuevo = True
                If lnTotal > 0 Then
                   mrs_Tmp.Find "Cie10='" & lcCie10 & "'"
                   If Not mrs_Tmp.EOF Then
                      lbNuevo = False
                   End If
                End If
                If lbNuevo = True Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields("Cie10").Value = lcCie10
                    mrs_Tmp.Fields("Diagnostico").Value = Left(ldDiagnostico, 50)
                    mrs_Tmp.Fields("Cantidad").Value = lnTotDNro
                Else
                    mrs_Tmp.Fields("Cantidad").Value = mrs_Tmp.Fields("Cantidad").Value + lnTotDNro
                End If
                mrs_Tmp.Update
            Loop
         End If
     End If
     '***** emergencia
     If lnTipoServicio = 0 Or lnTipoServicio = sghTipoServicio.sghEmergenciaConsultorios Then
        Set rsReporte = AtencionesDiagnosticosSeleccionarEmergenciaEgresadosPorFechasSexoTipodx(0, mda_FechaInicio, mda_FechaFin)
        If lnIdFuenteFinanciamiento > 0 Then
           If lnIdFuenteFinanciamiento = 1 Then
               rsReporte.Filter = "idFuenteFinanciamiento=1 or idFuenteFinanciamiento=5"
           Else
               rsReporte.Filter = "idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento
           End If
        End If
        If rsReporte.RecordCount > 0 Then
           rsReporte.MoveFirst
           Do While Not rsReporte.EOF
                lnTotDNro = 0
                ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value
                lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                If Not IsNull(rsReporte!cie10) Then
                 lcCie10 = rsReporte.Fields("Cie10").Value
                Else
                 lcCie10 = ""
                End If
                Do While Not rsReporte.EOF And ldDiagnostico = rsReporte.Fields("Ddiagnostico").Value And lnDiagnostico = rsReporte.Fields("IdDiagnostico").Value
                   lnTotDNro = lnTotDNro + 1
                   lnTotal = lnTotal + 1
                   rsReporte.MoveNext
                   If rsReporte.EOF Then
                      Exit Do
                   End If
                Loop
                lbNuevo = True
                If lnTotal > 0 Then
                   mrs_Tmp.Find "Cie10='" & lcCie10 & "'"
                   If Not mrs_Tmp.EOF Then
                      lbNuevo = False
                   End If
                End If
                If lbNuevo = True Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields("Cie10").Value = lcCie10
                    mrs_Tmp.Fields("Diagnostico").Value = Left(ldDiagnostico, 50)
                    mrs_Tmp.Fields("Cantidad").Value = lnTotDNro
                Else
                    mrs_Tmp.Fields("Cantidad").Value = mrs_Tmp.Fields("Cantidad").Value + lnTotDNro
                End If
                mrs_Tmp.Update
           Loop
        End If
     End If
     If mrs_Tmp.RecordCount > 0 Then
        mrs_Tmp.Sort = "cantidad desc,diagnostico"
        'porcentajes
        If lnTotal > 0 Then
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                lnPorcentaje1 = Round(mrs_Tmp.Fields!Cantidad * 100 / lnTotal, 2)
                If mrs_Tmp.AbsolutePosition = 1 Then
                   lnPorcentaje2 = lnPorcentaje1
                Else
                   lnPorcentaje2 = lnPorcentaje1 + lnPorcentaje2
                End If
                mrs_Tmp.Fields!Porcentaje = lnPorcentaje1
                mrs_Tmp.Fields!Acumulado = IIf(lnPorcentaje2 > 100, 100, lnPorcentaje2)
                mrs_Tmp.Update
                mrs_Tmp.MoveNext
            Loop
        End If
        'total general
        mrs_Tmp.AddNew
        mrs_Tmp.Fields!Diagnostico = "Total General"
        mrs_Tmp.Fields!Cantidad = lnTotal
        mrs_Tmp.Update
        '
        If lbEnExcel = True Then
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, ml_TextoDelFiltro, "", lnHwnd, True
        End If
     End If

     Set mrs_Tmp = Nothing
     Set rsReporte = Nothing
     Set mo_ReglasReportes = Nothing
End Sub
Function AtencionesDiagnosticosSeleccionarCEatendidosPorFechasSexoTipodx(lnTipoSexo As Integer, lnTipoDx As Integer, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set AtencionesDiagnosticosSeleccionarCEatendidosPorFechasSexoTipodx = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionarCEatendidosPorFechasSexoTipodx"
        Set oParameter = .CreateParameter("@TipoSexo", adInteger, adParamInput, 0, lnTipoSexo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoDx", adInteger, adParamInput, 0, lnTipoDx): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrIni", adDate, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDate, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDiagnosticosSeleccionarCEatendidosPorFechasSexoTipodx = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function AtencionesDiagnosticosSeleccionarHospitalizadosEgresadosPorFechasSexoTipodx(lnTipoSexo As Integer, lnTipoDx As Integer, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set AtencionesDiagnosticosSeleccionarHospitalizadosEgresadosPorFechasSexoTipodx = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionarHospitalizadosEgresadosPorFechasSexoTipodx"
        Set oParameter = .CreateParameter("@TipoSexo", adInteger, adParamInput, 0, lnTipoSexo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoDx", adInteger, adParamInput, 0, lnTipoDx): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrIni", adDate, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDate, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDiagnosticosSeleccionarHospitalizadosEgresadosPorFechasSexoTipodx = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function AtencionesDiagnosticosSeleccionarEmergenciaEgresadosPorFechasSexoTipodx(lnTipoSexo As Integer, ldFechaEgresoIni As Date, ldFechaEgresoFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set AtencionesDiagnosticosSeleccionarEmergenciaEgresadosPorFechasSexoTipodx = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesDiagnosticosSeleccionarEmergenciaEgresadosPorFechasSexoTipodx"
        Set oParameter = .CreateParameter("@TipoSexo", adInteger, adParamInput, 0, lnTipoSexo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrIni", adDate, adParamInput, 0, ldFechaEgresoIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaEgrFin", adDate, adParamInput, 0, ldFechaEgresoFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesDiagnosticosSeleccionarEmergenciaEgresadosPorFechasSexoTipodx = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub AtencionesServiciosIntermedios(lnAnio As Integer, lnTrimestre As Integer, lnIdFuenteFinanciamiento As Long, _
                                    lnTipoServicio As Integer, lcTitulo As String, ml_TextoDelFiltro As String, _
                                    lbEnExcel As Boolean, lnHwnd As Long)
                                    
    Dim mo_ReglasFacturacion As New ReglasFacturacion
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim rsReporte  As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim oConexion As New Connection
    Dim lnTotDNro As Long, ldDiagnostico As String, lnDiagnostico As Long, lcCie10 As String
    Dim mda_FechaInicio As Date, mda_FechaFin As Date, lnTotal As Long, lbNuevo As Boolean
    Dim lnPorFarmacia As Double, lnPorImagenes As Double, lnPorcLaboratorio As Double, lnPorcHoteleria As Double, lnPorcOtros As Double
    Dim lnTotFarmacia As Double, lnTotImagenes As Double, lnTotLaboratorio As Double, lnTotHoteleria As Double, lnTotOtros As Double
    With mrs_Tmp
          .Fields.Append "Concepto", adVarChar, 20, adFldIsNullable
          .Fields.Append "Total", adDouble
          .Fields.Append "Farmacia", adInteger
          .Fields.Append "Laboratorio", adDouble
          .Fields.Append "Imagenes", adDouble
          .Fields.Append "Hoteleria", adDouble
          .Fields.Append "Otros", adDouble
          .LockType = adLockOptimistic
          .Open
          .AddNew
          .Fields!Concepto = "ConsultaExterna"
          .Update
          .AddNew
          .Fields!Concepto = "Hospitalario"
          .Update
          .AddNew
          .Fields!Concepto = "Emergencia"
          .Update
          .AddNew
          .Fields!Concepto = "Otros"
          .Update
    End With
    With mrs_Tmp1
          .Fields.Append "Concepto", adVarChar, 20, adFldIsNullable
          .Fields.Append "Total", adVarChar, 10, adFldIsNullable
          .Fields.Append "Farmacia", adVarChar, 10, adFldIsNullable
          .Fields.Append "Laboratorio", adVarChar, 10, adFldIsNullable
          .Fields.Append "Imagenes", adVarChar, 10, adFldIsNullable
          .Fields.Append "Hoteleria", adVarChar, 10, adFldIsNullable
          .Fields.Append "Otros", adVarChar, 10, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    mda_FechaInicio = CDate("01/" & IIf(lnTrimestre = 0, "01/", IIf(lnTrimestre = 1, "04/", IIf(lnTrimestre = 2, "07/", "10/"))) & Trim(Str(lnAnio)))
    mda_FechaFin = CDate(IIf(lnTrimestre = 0, "31/03/", IIf(lnTrimestre = 1, "30/06/", IIf(lnTrimestre = 2, "30/09/", "31/12/"))) & Trim(Str(lnAnio)))
    If lnIdFuenteFinanciamiento = 1 Then
        Set rsReporte = AtencionesSeleccionarXfechaEgresoMedico(mda_FechaInicio, mda_FechaFin, 0)
        rsReporte.Filter = "idFuenteFinanciamiento=1 or idFuenteFinanciamiento=5"
    Else
        Set rsReporte = AtencionesSeleccionarXfechaEgresoMedico(mda_FechaInicio, mda_FechaFin, lnIdFuenteFinanciamiento)
    End If
    If rsReporte.RecordCount = 0 Then
       MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighentidades.CadenaConexion
       rsReporte.MoveFirst
    
       lnTotFarmacia = 0: lnTotImagenes = 0: lnTotLaboratorio = 0: lnTotHoteleria = 0: lnTotOtros = 0
       Do While Not rsReporte.EOF
          Set oRsTmp1 = mo_ReglasFacturacion.FacturacionCuentasAtencionPtosXcuenta(rsReporte!idCuentaAtencion, oConexion)
          If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF
                mrs_Tmp.MoveFirst
                Select Case rsReporte!idTipoServicio
                Case sghTipoServicio.sghConsultaExterna
                Case sghTipoServicio.sghEmergenciaConsultorios
                     mrs_Tmp.Move 2
                Case sghTipoServicio.sghHospitalizacion
                     mrs_Tmp.Move 1
                Case Else
                     mrs_Tmp.Move 3
                End Select
                Select Case oRsTmp1!IdPuntoCarga
                Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, _
                          sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, _
                          sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica
                     mrs_Tmp.Fields!laboratorio = mrs_Tmp.Fields!laboratorio + oRsTmp1!totalConsumos
                     lnTotLaboratorio = lnTotLaboratorio + oRsTmp1!totalConsumos
                Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, _
                         sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia
                     mrs_Tmp.Fields!Imagenes = mrs_Tmp.Fields!Imagenes + oRsTmp1!totalConsumos
                     lnTotImagenes = lnTotImagenes + oRsTmp1!totalConsumos
                Case sghPuntosCargaBasicos.sghPtoCargaFarmacia
                     mrs_Tmp.Fields!Farmacia = mrs_Tmp.Fields!Farmacia + oRsTmp1!totalConsumos
                     lnTotFarmacia = lnTotFarmacia + oRsTmp1!totalConsumos
                Case sghPuntosCargaBasicos.sghPtoCargaAdmisionHospitalizacion
                     mrs_Tmp.Fields!Hoteleria = mrs_Tmp.Fields!Hoteleria + oRsTmp1!totalConsumos
                     lnTotHoteleria = lnTotHoteleria + oRsTmp1!totalConsumos
                Case Else
                     mrs_Tmp.Fields!Otros = mrs_Tmp.Fields!Otros + oRsTmp1!totalConsumos
                     lnTotOtros = lnTotOtros + oRsTmp1!totalConsumos
                End Select
                mrs_Tmp.Fields!Total = mrs_Tmp.Fields!Total + oRsTmp1!totalConsumos
                mrs_Tmp.Update
                oRsTmp1.MoveNext
             Loop
          End If
          rsReporte.MoveNext
       Loop
       'Calcular %
       lnTotal = lnTotLaboratorio + lnTotImagenes + lnTotFarmacia + lnTotHoteleria + lnTotOtros
       If lnTotal = 0 Then
            MsgBox "No existe información con esos datos", vbInformation, "Reporte"
       Else
            mrs_Tmp1.AddNew
            mrs_Tmp1.Fields!Concepto = "Total"
            mrs_Tmp1.Fields!Total = Format(lnTotal, "#,###,###.#0")
            mrs_Tmp1.Fields!Farmacia = Format(lnTotFarmacia, "#,###,###.#0")
            mrs_Tmp1.Fields!laboratorio = Format(lnTotLaboratorio, "#,###,###.#0")
            mrs_Tmp1.Fields!Imagenes = Format(lnTotImagenes, "#,###,###.#0")
            mrs_Tmp1.Fields!Hoteleria = Format(lnTotHoteleria, "#,###,###.#0")
            mrs_Tmp1.Fields!Otros = Format(lnTotOtros, "#,###,###.#0")
            mrs_Tmp1.Update
            mrs_Tmp1.AddNew
            mrs_Tmp1.Fields!Concepto = "%"
            mrs_Tmp1.Fields!Total = "100"
            mrs_Tmp1.Fields!Farmacia = Trim(Str(Round(lnTotFarmacia * 100 / lnTotal, 1))) & " %"
            mrs_Tmp1.Fields!laboratorio = Trim(Str(Round(lnTotLaboratorio * 100 / lnTotal, 1))) & " %"
            mrs_Tmp1.Fields!Imagenes = Trim(Str(Round(lnTotImagenes * 100 / lnTotal, 1))) & " %"
            mrs_Tmp1.Fields!Hoteleria = Trim(Str(Round(lnTotHoteleria * 100 / lnTotal, 1))) & " %"
            mrs_Tmp1.Fields!Otros = Trim(Str(Round(lnTotOtros * 100 / lnTotal, 1))) & " %"
            mrs_Tmp1.Update
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               mrs_Tmp1.AddNew
               mrs_Tmp1.Fields!Concepto = mrs_Tmp.Fields!Concepto
               mrs_Tmp1.Fields!Total = Format(mrs_Tmp.Fields!Total, "#,###,###.#0")
               If mrs_Tmp.Fields!Total = 0 Then
                    mrs_Tmp1.Fields!Farmacia = "0 %"
                    mrs_Tmp1.Fields!laboratorio = "0 %"
                    mrs_Tmp1.Fields!Imagenes = "0 %"
                    mrs_Tmp1.Fields!Hoteleria = "0 %"
                    mrs_Tmp1.Fields!Otros = "0 %"
               Else
                    mrs_Tmp1.Fields!Farmacia = Trim(Str(Round(mrs_Tmp.Fields!Farmacia * 100 / mrs_Tmp.Fields!Total, 1))) & " %"
                    mrs_Tmp1.Fields!laboratorio = Trim(Str(Round(mrs_Tmp.Fields!laboratorio * 100 / mrs_Tmp.Fields!Total, 1))) & " %"
                    mrs_Tmp1.Fields!Imagenes = Trim(Str(Round(mrs_Tmp.Fields!Imagenes * 100 / mrs_Tmp.Fields!Total, 1))) & " %"
                    mrs_Tmp1.Fields!Hoteleria = Trim(Str(Round(mrs_Tmp.Fields!Hoteleria * 100 / mrs_Tmp.Fields!Total, 1))) & " %"
                    mrs_Tmp1.Fields!Otros = Trim(Str(Round(mrs_Tmp.Fields!Otros * 100 / mrs_Tmp.Fields!Total, 1))) & " %"
               End If
               mrs_Tmp1.Update
               mrs_Tmp.MoveNext
            Loop
            oConexion.Close
            If lbEnExcel = True Then
                mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp1, lcTitulo, ml_TextoDelFiltro, "", lnHwnd, True
            End If
       End If
    End If
    Set mo_ReglasFacturacion = Nothing
    Set mrs_Tmp = Nothing
    Set mrs_Tmp1 = Nothing
    Set rsReporte = Nothing
    Set oRsTmp1 = Nothing
    Set mo_ReglasReportes = Nothing
    Set oConexion = Nothing
    
End Sub

Function AtencionesSeleccionarXfechaEgresoMedico(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarXfechaEgresoMedico = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarXfechaEgresoMedico"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarXfechaEgresoMedico = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub AtencionesRecetas(lnAnio As Integer, lnTrimestre As Integer, lnIdFuenteFinanciamiento As Long, _
                                    lnTipoServicio As Integer, lcTitulo As String, ml_TextoDelFiltro As String, _
                                    lbEnExcel As Boolean, lnHwnd As Long)
    Dim rsReporte As New Recordset
    Dim mrs_Tmp As New Recordset
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim mda_FechaInicio As Date, mda_FechaFin As Date, lcFilter As String
    Dim lnIdServicio As Long, lnTotPrescrito As Double, lnTotEntregado As Double, lnTotal As Double
    Dim lnTotPresXServicio As Double, lnTotEntXservicio As Double, lcServicio As String
    mda_FechaInicio = CDate("01/" & IIf(lnTrimestre = 0, "01/", IIf(lnTrimestre = 1, "04/", IIf(lnTrimestre = 2, "07/", "10/"))) & Trim(Str(lnAnio)))
    mda_FechaFin = CDate(IIf(lnTrimestre = 0, "31/03/", IIf(lnTrimestre = 1, "30/06/", IIf(lnTrimestre = 2, "30/09/", "31/12/"))) & Trim(Str(lnAnio)))
    lcFilter = ""
    If lnIdFuenteFinanciamiento = 1 Then
        Set rsReporte = RecetasPrescritosYentregadas(mda_FechaInicio, mda_FechaFin, 0)
        lcFilter = "(idFuenteFinanciamiento=1 or idFuenteFinanciamiento=5)"
    Else
        Set rsReporte = RecetasPrescritosYentregadas(mda_FechaInicio, mda_FechaFin, lnIdFuenteFinanciamiento)
    End If
    If lnTipoServicio > 0 Then
       If lcFilter = "" Then
          lcFilter = "idTipoServicio=" & lnTipoServicio
       Else
          lcFilter = " and idTipoServicio=" & lnTipoServicio
       End If
    End If
    If lcFilter <> "" Then
        rsReporte.Filter = lcFilter
    End If
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        With mrs_Tmp
                 .Fields.Append "Servicio", adVarChar, 50, adFldIsNullable
                 .Fields.Append "Prescrito", adDouble
                 .Fields.Append "Entregado", adDouble
                 .Fields.Append "Deficit", adDouble
                 .Fields.Append "Porc", adDouble
                 .LockType = adLockOptimistic
                 .Open
       End With
       lnTotPrescrito = 0: lnTotEntregado = 0
       rsReporte.MoveFirst
       Do While Not rsReporte.EOF
          lnTotPresXServicio = 0: lnTotEntXservicio = 0
          lcServicio = Left(rsReporte!Servicio, 43) & IIf(lnTipoServicio <> 0, "", " (" & Left(rsReporte!TipoServicio, 3) & ")")
          lnIdServicio = rsReporte!idServicio
          Do While Not rsReporte.EOF And lnIdServicio = rsReporte!idServicio
             lnTotPresXServicio = lnTotPresXServicio + Round(rsReporte!cantidadPedida * rsReporte!precio, 2)
             If Not IsNull(rsReporte!cantidadDespachada) Then
                lnTotEntXservicio = lnTotEntXservicio + Round(rsReporte!cantidadDespachada * rsReporte!precio, 2)
             End If
             rsReporte.MoveNext
             If rsReporte.EOF Then
                Exit Do
             End If
          Loop
          mrs_Tmp.AddNew
          mrs_Tmp.Fields!Servicio = lcServicio
          mrs_Tmp.Fields!Prescrito = lnTotPresXServicio
          mrs_Tmp.Fields!Entregado = lnTotEntXservicio
          mrs_Tmp.Fields!deficit = lnTotPresXServicio - lnTotEntXservicio
          mrs_Tmp.Fields!Porc = Round((lnTotPresXServicio - lnTotEntXservicio) * 100 / lnTotPresXServicio, 1)
          mrs_Tmp.Update
          lnTotPrescrito = lnTotPrescrito + lnTotPresXServicio
          lnTotEntregado = lnTotEntregado + lnTotEntXservicio
       Loop
       '%
       mrs_Tmp.AddNew
       mrs_Tmp.Fields!Servicio = " Total General: "
       mrs_Tmp.Fields!Prescrito = lnTotPrescrito
       mrs_Tmp.Fields!Entregado = lnTotEntregado
       mrs_Tmp.Fields!deficit = lnTotPrescrito - lnTotEntregado
       mrs_Tmp.Fields!Porc = Round((lnTotPrescrito - lnTotEntregado) * 100 / lnTotPrescrito, 1)
       mrs_Tmp.Update
       mrs_Tmp.Sort = "servicio"
       If lbEnExcel = True Then
                mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, ml_TextoDelFiltro, "", lnHwnd, True
       End If
    End If
    Set rsReporte = Nothing
    Set mrs_Tmp = Nothing
    Set mo_ReglasReportes = Nothing
End Sub

Function RecetasPrescritosYentregadas(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set RecetasPrescritosYentregadas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "RecetasPrescritosYentregadas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RecetasPrescritosYentregadas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function




'debb-09/08/2016
Sub ProcesaMuerteHospitalizados(lcTitulo As String, ml_TextoDelFiltro As String, lcFechaInicial As String, _
                                 lcFechaFinal As String, ml_IdTipoEspecialidad As Integer, lbEnExcel As Boolean, _
                                 lnHwnd As Long)
    Dim rsReporte As New Recordset
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
    Dim mo_AdminAdmision As New ReglasAdmision
    Dim mrs_Tmp As New Recordset
    Dim rsServicio As New Recordset
    Dim oRsPie As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mda_FechaInicio As Date, mda_FechaFin As Date, lcFilter As String, lcServEgreso As String
    Dim lcHoraEstanciaMax As String, lnDiasEstancia As Integer, lnHorasEstancia As Integer
    Dim lnEstMenor24 As Integer, lnEstEntre24y48 As Integer, lnEstMayor48 As Integer, lcNroEmergencia As String
    Dim lcDxbasico As String, lcDxInterm As String, lcMedicoEgreso As String, lcRenaes As String
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    mda_FechaInicio = CDate(lcFechaInicial)
    mda_FechaFin = CDate(lcFechaFinal)
    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        
        With mrs_Tmp
            .Fields.Append "Historia", adVarChar, 10
            .Fields.Append "Paciente", adVarChar, 50
            .Fields.Append "Sexo", adVarChar, 15
            .Fields.Append "Edad", adVarChar, 3
            .Fields.Append "EdadTipo", adVarChar, 3
            .Fields.Append "Procedencia", adVarChar, 20
            .Fields.Append "Fingreso", adVarChar, 20
            .Fields.Append "Fegreso", adVarChar, 20
            .Fields.Append "Estancia", adInteger
            .Fields.Append "EstMenorVcuatro", adInteger
            .Fields.Append "EstVcuatroCocho", adInteger
            .Fields.Append "EstMayorCocho", adInteger
            .Fields.Append "Servicio", adVarChar, 20
            .Fields.Append "CausaBasica", adVarChar, 50
            .Fields.Append "CausaInterm", adVarChar, 50
            .Fields.Append "ResponsableAten", adVarChar, 50
            .LockType = adLockOptimistic
            .Open
       End With
       Set oRsPie = sighentidades.CopyRecordset(mrs_Tmp, "")
       lnEstMenor24 = 0: lnEstEntre24y48 = 0: lnEstMayor48 = 0
       rsReporte.MoveFirst
       Do While Not rsReporte.EOF
          If UCase(Left(rsReporte!CondicionAlta, 5)) = "FALLE" Then
                lcNroEmergencia = ""
                Set rsServicio = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro("dbo.AtencionesDatosAdicionales.idAtencion=" & rsReporte.Fields!idAtencion, oConexion)
                If rsServicio.RecordCount > 0 Then
                   If Not IsNull(rsServicio.Fields!emergenciaCorrelativo) Then
                      lcNroEmergencia = Mid(rsServicio.Fields!emergenciaCorrelativo, 5, 20)
                   End If
                End If
                rsServicio.Close
                '
                Set rsServicio = mo_AdminServComunes.ServiciosSeleccionarXidentificador(rsReporte.Fields("idServicioEgreso").Value)
                lcServEgreso = rsServicio.Fields("nombre").Value
                rsServicio.Close
                '
                Set rsServicio = mo_AdminServComunes.AtencionesSeleccionarMedicoEgresoPorCuenta(rsReporte!idCuentaAtencion)
                lcMedicoEgreso = rsServicio!dmedico
                rsServicio.Close
                '
                Set rsServicio = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
                rsServicio.Filter = "idSubClasificacionDx=303"
                lcDxbasico = IIf(rsServicio.RecordCount = 0, "", rsServicio!CodigoDx & " " & rsServicio!DescripcionDx)
                rsServicio.Filter = "idSubClasificacionDx<>303"
                lcDxInterm = IIf(rsServicio.RecordCount = 0, "", rsServicio!CodigoDx & " " & rsServicio!DescripcionDx)
                rsServicio.Close
                '
                If Val(lcRenaes) = 6918 Then  'hosp tarapoto
                    lnDiasEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax, True)
                Else
                    lnDiasEstancia = lcBuscaParametro.DiasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso, lcHoraEstanciaMax)
                End If
                lnHorasEstancia = lcBuscaParametro.HorasDelPacienteEnHospitalizacionEmergencia(rsReporte!FechaIngreso, rsReporte!HoraIngreso, rsReporte!FechaEgreso, rsReporte!horaEgreso)
                lnEstMenor24 = lnEstMenor24 + IIf(lnHorasEstancia < 24, 1, 0)
                lnEstEntre24y48 = lnEstEntre24y48 + IIf(lnHorasEstancia >= 24 And lnHorasEstancia <= 48, 1, 0)
                lnEstMayor48 = lnEstMayor48 + IIf(lnHorasEstancia > 48, 1, 0)
                mrs_Tmp.AddNew
                If Val(lcRenaes) = 6918 Then  'hosp tarapoto
                   If ml_IdTipoEspecialidad = 3 Then
                      mrs_Tmp.Fields!historia = IIf(IsNull(rsReporte!NroHistoriaClinica), "", Right("000000" & Trim(Str(rsReporte!NroHistoriaClinica)), 6))
                   Else
                      mrs_Tmp.Fields!historia = lcNroEmergencia
                   End If
                Else
                   mrs_Tmp.Fields!historia = IIf(IsNull(rsReporte!NroHistoriaClinica), "", rsReporte!NroHistoriaClinica)
                End If
                mrs_Tmp.Fields!Paciente = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & _
                                          Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre), 50)
                mrs_Tmp.Fields!Sexo = rsReporte!Sexo
                mrs_Tmp.Fields!Edad = rsReporte!Edad
                mrs_Tmp.Fields!EdadTipo = rsReporte!TipoEdad
                mrs_Tmp.Fields!Procedencia = IIf(IsNull(rsReporte!distrito), "", Left(rsReporte!distrito, 20))
                mrs_Tmp.Fields!Fingreso = Format(rsReporte!FechaIngreso & " " & rsReporte!HoraIngreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                mrs_Tmp.Fields!Fegreso = Format(rsReporte!FechaEgreso & " " & rsReporte!horaEgreso, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
                mrs_Tmp.Fields!Estancia = lnDiasEstancia
                mrs_Tmp.Fields!EstMenorVcuatro = IIf(lnHorasEstancia < 24, 1, 0)
                mrs_Tmp.Fields!EstVcuatroCocho = IIf(lnHorasEstancia >= 24 And lnHorasEstancia <= 48, 1, 0)
                mrs_Tmp.Fields!EstMayorCocho = IIf(lnHorasEstancia > 48, 1, 0)
                mrs_Tmp.Fields!Servicio = Left(lcServEgreso, 20)
                mrs_Tmp.Fields!CausaBasica = Left(lcDxbasico, 50)
                mrs_Tmp.Fields!CausaInterm = Left(lcDxInterm, 50)
                mrs_Tmp.Fields!ResponsableAten = lcMedicoEgreso
                mrs_Tmp.Update
          End If
          rsReporte.MoveNext
       Loop
       '
       oRsPie.AddNew
       oRsPie.Fields!historia = ""
       oRsPie.Fields!Paciente = "N° total: " & Trim(Str(mrs_Tmp.RecordCount))
       oRsPie.Fields!Sexo = ""
       oRsPie.Fields!Edad = 0
       oRsPie.Fields!EdadTipo = ""
       oRsPie.Fields!Procedencia = ""
       oRsPie.Fields!Fingreso = ""
       oRsPie.Fields!Fegreso = ""
       oRsPie.Fields!Estancia = 0
       oRsPie.Fields!EstMenorVcuatro = lnEstMenor24
       oRsPie.Fields!EstVcuatroCocho = lnEstEntre24y48
       oRsPie.Fields!EstMayorCocho = lnEstMayor48
       oRsPie.Fields!Servicio = ""
       oRsPie.Fields!CausaBasica = ""
       oRsPie.Fields!CausaInterm = ""
       oRsPie.Fields!ResponsableAten = ""
       oRsPie.Update
       '
       mrs_Tmp.Sort = "fingreso"
       If lbEnExcel = True Then
           mo_AdminReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, ml_TextoDelFiltro, "", lnHwnd, True, _
                                                    True, oRsPie
       End If
    End If
    oConexion.Close
    Set rsReporte = Nothing
    Set mo_AdminReportes = Nothing
    Set mo_AdminServComunes = Nothing
    Set mrs_Tmp = Nothing
    Set rsServicio = Nothing
    Set oRsPie = Nothing
    Set lcBuscaParametro = Nothing
    Set oConexion = Nothing
    Set mo_AdminAdmision = Nothing
End Sub

'debb-09/08/2016
Sub ProcesaMuerteHospitalizadosMenoresEdad(lcTitulo As String, ml_TextoDelFiltro As String, lcFechaInicial As String, _
                                 lcFechaFinal As String, ml_IdTipoEspecialidad As Integer, lbEnExcel As Boolean, _
                                 lnEdad As Integer, _
                                 lnHwnd As Long)
    Dim rsReporte As New Recordset
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_AdminServComunes As New SIGHNegocios.ReglasComunes
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mrs_Tmp As New Recordset
    Dim rsServicio As New Recordset
    Dim oRsPie As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_Pacientes As New DOPaciente
    Dim mo_distritos As New DODistrito
    Dim oEdad As Edad
    Dim mda_FechaInicio As Date, mda_FechaFin As Date, lcFilter As String, lcServEgreso As String
    Dim lcHoraEstanciaMax As String, lnDiasEstancia As Integer, lnHorasEstancia As Integer
    Dim lnEstMenor24 As Integer, lnEstEntre24y48 As Integer, lnEstMayor48 As Integer, lcRenaes As String
    Dim lcDxbasico As String, lcDxInterm As String, lcMedicoEgreso As String, lbContinuar As Boolean
    Dim lcMama As String, lcMamaDNI As String, lcMamaEdad As Integer, lcMamaProcedencia As String
    lcRenaes = Right("00000000" & lcBuscaParametro.SeleccionaFilaParametro(208), 10)
    lcHoraEstanciaMax = lcBuscaParametro.SeleccionaFilaParametro(201)
    mda_FechaInicio = CDate(lcFechaInicial)
    mda_FechaFin = CDate(lcFechaFinal)
    Set rsReporte = mo_AdminReportes.ReporteEgresosHospitalarios(0, 0, 0, mda_FechaInicio, mda_FechaFin, ml_IdTipoEspecialidad)
    
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos datos", vbInformation, "Reporte"
    Else
        
        With mrs_Tmp
            .Fields.Append "Historia", adVarChar, 10
            .Fields.Append "Paciente", adVarChar, 50
            .Fields.Append "Sexo", adVarChar, 15
            .Fields.Append "Edad", adVarChar, 3
            .Fields.Append "EdadTipo", adVarChar, 3
            .Fields.Append "mama", adVarChar, 50
            .Fields.Append "mamaEdad", adInteger
            .Fields.Append "mamaProcedencia", adVarChar, 20
            .Fields.Append "CausaBasica", adVarChar, 50
            .Fields.Append "CausaInterm", adVarChar, 50
            .Fields.Append "ResponsableAten", adVarChar, 50
            .LockType = adLockOptimistic
            .Open
       End With
       Set oRsPie = sighentidades.CopyRecordset(mrs_Tmp, "")
       lnEstMenor24 = 0: lnEstEntre24y48 = 0: lnEstMayor48 = 0
       rsReporte.MoveFirst
       Do While Not rsReporte.EOF
          lbContinuar = True
          If rsReporte!TipoEdad = "A" Then
             If rsReporte!Edad > lnEdad Then
                lbContinuar = False
             End If
          End If
          If UCase(Left(rsReporte!CondicionAlta, 5)) = "FALLE" And lbContinuar = True Then
                lcMama = "": lcMamaDNI = "": lcMamaEdad = 0: lcMamaProcedencia = ""
                Set rsServicio = mo_ReglasAdmision.PacientesSeleccionarPorIdentificador(rsReporte!idPaciente)
                If rsServicio.RecordCount > 0 Then
                   lcMama = IIf(IsNull(rsServicio!madreApellidoPaterno), "", rsServicio!madreApellidoPaterno) & " " & _
                          IIf(IsNull(rsServicio!madreApellidoMaterno), "", rsServicio!madreApellidoMaterno) & " " & _
                          IIf(IsNull(rsServicio!madrePrimerNombre), "", rsServicio!madrePrimerNombre)
                   lcMamaDNI = IIf(IsNull(rsServicio!madreDocumento), "", rsServicio!madreDocumento)
                   If Len(lcMamaDNI) = 8 Then
                        Set mo_Pacientes = mo_ReglasAdmision.PacientesSeleccionarPorDNI(lcMamaDNI)
                        If mo_Pacientes.idPaciente > 0 Then
                           oEdad = CalcularEdad(mo_Pacientes.FechaNacimiento, _
                                    CDate(Format(rsReporte!FechaIngreso, "dd/mm/yyyy") & " " & rsReporte!HoraIngreso))
                           lcMamaEdad = oEdad.Edad
                           Set mo_distritos = mo_AdminServComunes.DistritoSeleccionarPorId(mo_Pacientes.IdDistritoDomicilio)
                           lcMamaProcedencia = mo_distritos.nombre
                        End If
                   End If
                End If
                rsServicio.Close
                '
                Set rsServicio = mo_AdminServComunes.AtencionesSeleccionarMedicoEgresoPorCuenta(rsReporte!idCuentaAtencion)
                lcMedicoEgreso = rsServicio!dmedico
                rsServicio.Close
                '
                Set rsServicio = mo_AdminReportes.ReporteAtencionesDiagnosticosDeEgreso(rsReporte!idAtencion)
                rsServicio.Filter = "idSubClasificacionDx=303"
                lcDxbasico = IIf(rsServicio.RecordCount = 0, "", rsServicio!CodigoDx & " " & rsServicio!DescripcionDx)
                rsServicio.Filter = "idSubClasificacionDx<>303"
                lcDxInterm = IIf(rsServicio.RecordCount = 0, "", rsServicio!CodigoDx & " " & rsServicio!DescripcionDx)
                rsServicio.Close
                '
                mrs_Tmp.AddNew
                If Val(lcRenaes) = 6918 Then  'hosp tarapoto
                   mrs_Tmp.Fields!historia = IIf(IsNull(rsReporte!NroHistoriaClinica), "", Right("000000" & Trim(Str(rsReporte!NroHistoriaClinica)), 6))
                Else
                   mrs_Tmp.Fields!historia = IIf(IsNull(rsReporte!NroHistoriaClinica), "", rsReporte!NroHistoriaClinica)
                End If
                mrs_Tmp.Fields!Paciente = Left(Trim(rsReporte.Fields!ApellidoPaterno) & " " & _
                                          Trim(rsReporte.Fields!ApellidoMaterno) & " " & Trim(rsReporte.Fields!PrimerNombre), 50)
                mrs_Tmp.Fields!Sexo = rsReporte!Sexo
                mrs_Tmp.Fields!Edad = rsReporte!Edad
                mrs_Tmp.Fields!EdadTipo = rsReporte!TipoEdad
                mrs_Tmp.Fields!mama = lcMama
                mrs_Tmp.Fields!mamaEdad = lcMamaEdad
                mrs_Tmp.Fields!mamaProcedencia = Left(lcMamaProcedencia, 20)
                mrs_Tmp.Fields!CausaBasica = Left(lcDxbasico, 50)
                mrs_Tmp.Fields!CausaInterm = Left(lcDxInterm, 50)
                mrs_Tmp.Fields!ResponsableAten = lcMedicoEgreso
                mrs_Tmp.Update
          End If
          rsReporte.MoveNext
       Loop
       '
       oRsPie.AddNew
       oRsPie.Fields!historia = ""
       oRsPie.Fields!Paciente = "N° total: " & Trim(Str(mrs_Tmp.RecordCount))
       oRsPie.Fields!Sexo = ""
       oRsPie.Fields!Edad = 0
       oRsPie.Fields!EdadTipo = ""
       oRsPie.Fields!mama = ""
       oRsPie.Fields!mamaEdad = 0
       oRsPie.Fields!mamaProcedencia = ""
       oRsPie.Fields!CausaBasica = ""
       oRsPie.Fields!CausaInterm = ""
       oRsPie.Fields!ResponsableAten = ""
       oRsPie.Update
       '
       mrs_Tmp.Sort = "paciente"
       If lbEnExcel = True Then
           mo_AdminReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, ml_TextoDelFiltro, "", lnHwnd, True, _
                                                    True, oRsPie
       End If
    End If
    Set rsReporte = Nothing
    Set mo_AdminReportes = Nothing
    Set mo_AdminServComunes = Nothing
    Set mrs_Tmp = Nothing
    Set rsServicio = Nothing
    Set oRsPie = Nothing
    Set lcBuscaParametro = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_Pacientes = Nothing
    Set mo_distritos = Nothing
End Sub


