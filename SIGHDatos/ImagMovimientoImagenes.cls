VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagMovimientoImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ImagMovimientoImagenes
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla ImagMovimientoImagenes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOImagMovimientoImagenes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoImagenesAgregar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CorrelativoAnual", adInteger, adParamInput, 0, IIf(oTabla.CorrelativoAnual = 0, Null, oTabla.CorrelativoAnual))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPersonaTomaImagen", adInteger, adParamInput, 0, IIf(oTabla.idPersonaTomaImagen = 0, Null, oTabla.idPersonaTomaImagen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPersonaRecoge", adInteger, adParamInput, 0, IIf(oTabla.idPersonaRecoge = 0, Null, oTabla.idPersonaRecoge))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@zonaRayosX", adVarChar, adParamInput, 50, IIf(oTabla.zonaRayosX = "", Null, oTabla.zonaRayosX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PorcInformeRadiolog", adCurrency, adParamInput, 0, oTabla.PorcInformeRadiolog)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFinal", adVarChar, adParamInput, 3000, IIf(oTabla.ResultadoFinal = "", Null, oTabla.ResultadoFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsContraste", adInteger, adParamInput, 0, IIf(oTabla.EsContraste = 0, Null, oTabla.EsContraste))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsContrasteIonico", adInteger, adParamInput, 0, IIf(oTabla.EsContrasteIonico = 0, Null, oTabla.EsContrasteIonico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDiagnosticoDefinitivo", adInteger, adParamInput, 0, IIf(oTabla.EsDiagnosticoDefinitivo = 0, Null, oTabla.EsDiagnosticoDefinitivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_FUM", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Eo_FUM = 0, Null, oTabla.Eo_FUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_Gestantes", adVarChar, adParamInput, 2, IIf(oTabla.Eo_Gestantes = "", Null, oTabla.Eo_Gestantes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_Partos", adVarChar, adParamInput, 20, IIf(oTabla.Eo_Partos = "", Null, oTabla.Eo_Partos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_EG", adInteger, adParamInput, 0, IIf(oTabla.Eo_EG = 0, Null, oTabla.Eo_EG))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, Left(oTabla.Paciente, 100)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HcRx", adVarChar, adParamInput, 10, IIf(oTabla.HcRX = "", Null, oTabla.HcRX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla ImagMovimientoImagenes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOImagMovimientoImagenes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoImagenesModificar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, IIf(oTabla.IdOrden = 0, Null, oTabla.IdOrden))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CorrelativoAnual", adInteger, adParamInput, 0, IIf(oTabla.CorrelativoAnual = 0, Null, oTabla.CorrelativoAnual))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPersonaTomaImagen", adInteger, adParamInput, 0, IIf(oTabla.idPersonaTomaImagen = 0, Null, oTabla.idPersonaTomaImagen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPersonaRecoge", adInteger, adParamInput, 0, IIf(oTabla.idPersonaRecoge = 0, Null, oTabla.idPersonaRecoge))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@zonaRayosX", adVarChar, adParamInput, 50, IIf(oTabla.zonaRayosX = "", Null, oTabla.zonaRayosX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PorcInformeRadiolog", adCurrency, adParamInput, 0, oTabla.PorcInformeRadiolog)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFinal", adVarChar, adParamInput, 3000, IIf(oTabla.ResultadoFinal = "", Null, oTabla.ResultadoFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsContraste", adInteger, adParamInput, 0, IIf(oTabla.EsContraste = 0, Null, oTabla.EsContraste))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsContrasteIonico", adInteger, adParamInput, 0, IIf(oTabla.EsContrasteIonico = 0, Null, oTabla.EsContrasteIonico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDiagnosticoDefinitivo", adInteger, adParamInput, 0, IIf(oTabla.EsDiagnosticoDefinitivo = 0, Null, oTabla.EsDiagnosticoDefinitivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_FUM", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Eo_FUM = 0, Null, oTabla.Eo_FUM))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_Gestantes", adVarChar, adParamInput, 2, IIf(oTabla.Eo_Gestantes = "", Null, oTabla.Eo_Gestantes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_Partos", adVarChar, adParamInput, 20, IIf(oTabla.Eo_Partos = "", Null, oTabla.Eo_Partos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Eo_EG", adInteger, adParamInput, 0, IIf(oTabla.Eo_EG = 0, Null, oTabla.Eo_EG))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, IIf(oTabla.Paciente = "", Null, Left(oTabla.Paciente, 100)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@HcRx", adVarChar, adParamInput, 10, IIf(oTabla.HcRX = "", Null, oTabla.HcRX))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla ImagMovimientoImagenes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOImagMovimientoImagenes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoImagenesEliminar"
   Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla ImagMovimientoImagenes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOImagMovimientoImagenes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ImagMovimientoImagenesSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
       oTabla.IdOrden = IIf(IsNull(oRecordset!IdOrden), 0, oRecordset!IdOrden)
       oTabla.CorrelativoAnual = IIf(IsNull(oRecordset!CorrelativoAnual), 0, oRecordset!CorrelativoAnual)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.idPersonaTomaImagen = IIf(IsNull(oRecordset!idPersonaTomaImagen), 0, oRecordset!idPersonaTomaImagen)
       oTabla.idPersonaRecoge = IIf(IsNull(oRecordset!idPersonaRecoge), "", oRecordset!idPersonaRecoge)
       oTabla.zonaRayosX = IIf(IsNull(oRecordset!zonaRayosX), "", oRecordset!zonaRayosX)
       oTabla.PorcInformeRadiolog = oRecordset!PorcInformeRadiolog
       oTabla.ResultadoFinal = IIf(IsNull(oRecordset!ResultadoFinal), "", oRecordset!ResultadoFinal)
       oTabla.EsContraste = IIf(IsNull(oRecordset!EsContraste), "", oRecordset!EsContraste)
       oTabla.EsContrasteIonico = IIf(IsNull(oRecordset!EsContrasteIonico), "", oRecordset!EsContrasteIonico)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), "", oRecordset!IdDiagnostico)
       oTabla.EsDiagnosticoDefinitivo = IIf(IsNull(oRecordset!EsDiagnosticoDefinitivo), "", oRecordset!EsDiagnosticoDefinitivo)
       oTabla.Eo_FUM = IIf(IsNull(oRecordset!Eo_FUM), 0, oRecordset!Eo_FUM)
       oTabla.Eo_Gestantes = IIf(IsNull(oRecordset!Eo_Gestantes), "", oRecordset!Eo_Gestantes)
       oTabla.Eo_Partos = IIf(IsNull(oRecordset!Eo_Partos), "", oRecordset!Eo_Partos)
       oTabla.Eo_EG = IIf(IsNull(oRecordset!Eo_EG), 0, oRecordset!Eo_EG)
       oTabla.Paciente = IIf(IsNull(oRecordset!Paciente), 0, oRecordset!Paciente)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!idSexo), 0, oRecordset!idSexo)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.HcRX = IIf(IsNull(oRecordset!HcRX), "", oRecordset!HcRX)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOImagMovimientoImagenes)
   oTabla.IdMovimiento = 0
   oTabla.IdOrden = 0
   oTabla.CorrelativoAnual = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.IdComprobantePago = 0
   oTabla.idPersonaTomaImagen = 0
   oTabla.idPersonaRecoge = 0
   oTabla.zonaRayosX = ""
   oTabla.PorcInformeRadiolog = 0
   oTabla.ResultadoFinal = ""
   oTabla.EsContraste = 0
   oTabla.EsContrasteIonico = 0
   oTabla.IdDiagnostico = 0
   oTabla.EsDiagnosticoDefinitivo = 0
   oTabla.Eo_FUM = Now
   oTabla.Eo_Gestantes = ""
   oTabla.Eo_Partos = ""
   oTabla.Eo_EG = 0
   oTabla.Paciente = ""
   oTabla.IdTipoSexo = 0
   oTabla.FechaNacimiento = 0
End Sub
