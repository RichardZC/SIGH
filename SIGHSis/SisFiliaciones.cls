VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFiliaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el Mantenimiento de la tabla SisFiliaciones
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
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
'   Descripción:    Inserta un registro a la tabla SisFiliaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

'Actualizado 13102014
Function Insertar(ByVal oTabla As DoSisFiliaciones, lbUsaModificarSiYaExisteKEY As Boolean) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFiliacionesAgregar"
   Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, IIf(oTabla.idSiasis = 0, Null, oTabla.idSiasis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.codigo = "", "", oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionDisa", adVarChar, adParamInput, 3, IIf(oTabla.AfiliacionDisa = "", "", oTabla.AfiliacionDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionTipoFormato = "", "", oTabla.AfiliacionTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroFormato", adVarChar, adParamInput, 10, IIf(oTabla.AfiliacionNroFormato = "", "", oTabla.AfiliacionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroIntegrante", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionNroIntegrante = "", Null, oTabla.AfiliacionNroIntegrante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoTipo", adVarChar, adParamInput, 1, IIf(oTabla.DocumentoTipo = "", Null, oTabla.DocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablAdscripcion", adVarChar, adParamInput, 10, IIf(oTabla.CodigoEstablAdscripcion = "", Null, oTabla.CodigoEstablAdscripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.AfiliacionFecha = 0, Null, oTabla.AfiliacionFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paterno", adVarChar, adParamInput, 40, IIf(oTabla.Paterno = "", sighentidades.DevuelveSinApellido, mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Paterno))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Materno", adVarChar, adParamInput, 40, IIf(oTabla.Materno = "", sighentidades.DevuelveSinApellido, mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Materno))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pnombre", adVarChar, adParamInput, 70, IIf(oTabla.Pnombre = "", "", mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Pnombre))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Onombres", adVarChar, adParamInput, 70, IIf(oTabla.Onombres = "", "", mo_Teclado.DevuelveTextoSINtildes(mo_Teclado.CapitalizarNombres(oTabla.Onombres))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Genero", adVarChar, adParamInput, 1, IIf(oTabla.Genero = "", Null, oTabla.Genero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fnacimiento = 0, Null, oTabla.Fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDistritoDomicilio", adVarChar, adParamInput, 6, IIf(oTabla.IdDistritoDomicilio = "", Null, oTabla.IdDistritoDomicilio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Estado", adVarChar, adParamInput, 1, IIf(oTabla.Estado = "", Null, oTabla.Estado))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Fbaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fbaja = 0, Null, oTabla.Fbaja))
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fbaja", adVarChar, adParamInput, 10, oTabla.Fbaja)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 10, IIf(oTabla.DocumentoNumero = "", Null, oTabla.DocumentoNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoBaja", adVarChar, adParamInput, 70, IIf(oTabla.MotivoBaja = "", Null, oTabla.MotivoBaja))
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
   If Err.Number = -2147217873 And lbUsaModificarSiYaExisteKEY = True Then
      If Modificar(oTabla) = True Then
         Insertar = True
      End If
   End If
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla SisFiliaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSisFiliaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFiliacionesModificar"
   Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, IIf(oTabla.idSiasis = 0, Null, oTabla.idSiasis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionDisa", adVarChar, adParamInput, 3, IIf(oTabla.AfiliacionDisa = "", Null, oTabla.AfiliacionDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionTipoFormato", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionTipoFormato = "", Null, oTabla.AfiliacionTipoFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroFormato", adVarChar, adParamInput, 10, IIf(oTabla.AfiliacionNroFormato = "", Null, oTabla.AfiliacionNroFormato))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionNroIntegrante", adVarChar, adParamInput, 2, IIf(oTabla.AfiliacionNroIntegrante = "", Null, oTabla.AfiliacionNroIntegrante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoTipo", adVarChar, adParamInput, 1, IIf(oTabla.DocumentoTipo = "", Null, oTabla.DocumentoTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablAdscripcion", adVarChar, adParamInput, 10, IIf(oTabla.CodigoEstablAdscripcion = "", Null, oTabla.CodigoEstablAdscripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AfiliacionFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.AfiliacionFecha = 0, Null, oTabla.AfiliacionFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Paterno", adVarChar, adParamInput, 40, IIf(oTabla.Paterno = "", sighentidades.DevuelveSinApellido, mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Paterno))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Materno", adVarChar, adParamInput, 40, IIf(oTabla.Materno = "", sighentidades.DevuelveSinApellido, mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Materno))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Pnombre", adVarChar, adParamInput, 70, IIf(oTabla.Pnombre = "", Null, mo_Teclado.DevuelveTextoSINtildes(UCase(oTabla.Pnombre))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Onombres", adVarChar, adParamInput, 70, IIf(oTabla.Onombres = "", Null, mo_Teclado.DevuelveTextoSINtildes(mo_Teclado.CapitalizarNombres(oTabla.Onombres))))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Genero", adVarChar, adParamInput, 1, IIf(oTabla.Genero = "", Null, oTabla.Genero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fnacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fnacimiento = 0, Null, oTabla.Fnacimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDistritoDomicilio", adVarChar, adParamInput, 6, IIf(oTabla.IdDistritoDomicilio = "", Null, oTabla.IdDistritoDomicilio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Estado", adVarChar, adParamInput, 1, IIf(oTabla.Estado = "", Null, oTabla.Estado))
   .Parameters.Append oParameter
'   Set oParameter = .CreateParameter("@Fbaja", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fbaja = 0, Null, oTabla.Fbaja))
'   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fbaja", adVarChar, adParamInput, 10, oTabla.Fbaja)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 10, IIf(oTabla.DocumentoNumero = "", Null, oTabla.DocumentoNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MotivoBaja", adVarChar, adParamInput, 70, IIf(oTabla.MotivoBaja = "", Null, oTabla.MotivoBaja))
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
'   Descripción:    elimia o borra un registro a la tabla SisFiliaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSisFiliaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFiliacionesEliminar"
   Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, oTabla.idSiasis)
   .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, oTabla.codigo)
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
'   Descripción:    Selecciona un unico registro de la tabla SisFiliaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSisFiliaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFiliacionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, oTabla.idSiasis)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 2, oTabla.codigo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idSiasis = IIf(IsNull(oRecordset!idSiasis), 0, oRecordset!idSiasis)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.AfiliacionDisa = IIf(IsNull(oRecordset!AfiliacionDisa), "", oRecordset!AfiliacionDisa)
       oTabla.AfiliacionTipoFormato = IIf(IsNull(oRecordset!AfiliacionTipoFormato), "", oRecordset!AfiliacionTipoFormato)
       oTabla.AfiliacionNroFormato = IIf(IsNull(oRecordset!AfiliacionNroFormato), "", oRecordset!AfiliacionNroFormato)
       oTabla.AfiliacionNroIntegrante = IIf(IsNull(oRecordset!AfiliacionNroIntegrante), "", oRecordset!AfiliacionNroIntegrante)
       oTabla.DocumentoTipo = IIf(IsNull(oRecordset!DocumentoTipo), "", oRecordset!DocumentoTipo)
       oTabla.CodigoEstablAdscripcion = IIf(IsNull(oRecordset!CodigoEstablAdscripcion), "", oRecordset!CodigoEstablAdscripcion)
       oTabla.AfiliacionFecha = IIf(IsNull(oRecordset!AfiliacionFecha), 0, oRecordset!AfiliacionFecha)
       oTabla.Paterno = IIf(IsNull(oRecordset!Paterno), "", oRecordset!Paterno)
       oTabla.Materno = IIf(IsNull(oRecordset!Materno), "", oRecordset!Materno)
       oTabla.Pnombre = IIf(IsNull(oRecordset!Pnombre), "", oRecordset!Pnombre)
       oTabla.Onombres = IIf(IsNull(oRecordset!Onombres), "", oRecordset!Onombres)
       oTabla.Genero = IIf(IsNull(oRecordset!Genero), "", oRecordset!Genero)
       oTabla.Fnacimiento = IIf(IsNull(oRecordset!Fnacimiento), 0, oRecordset!Fnacimiento)
       oTabla.IdDistritoDomicilio = IIf(IsNull(oRecordset!IdDistritoDomicilio), "", oRecordset!IdDistritoDomicilio)
       oTabla.Estado = IIf(IsNull(oRecordset!Estado), "", oRecordset!Estado)
       oTabla.Fbaja = IIf(IsNull(oRecordset!Fbaja), "", oRecordset!Fbaja)
       oTabla.DocumentoNumero = IIf(IsNull(oRecordset!DocumentoNumero), "", oRecordset!DocumentoNumero)
       oTabla.MotivoBaja = IIf(IsNull(oRecordset!MotivoBaja), "", oRecordset!MotivoBaja)
       SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSisFiliaciones)
   oTabla.idSiasis = 0
   oTabla.codigo = ""
   oTabla.AfiliacionDisa = ""
   oTabla.AfiliacionTipoFormato = ""
   oTabla.AfiliacionNroFormato = ""
   oTabla.AfiliacionNroIntegrante = ""
   oTabla.DocumentoTipo = ""
   oTabla.CodigoEstablAdscripcion = ""
   oTabla.AfiliacionFecha = Now
   oTabla.Paterno = ""
   oTabla.Materno = ""
   oTabla.Pnombre = ""
   oTabla.Onombres = ""
   oTabla.Genero = ""
   oTabla.Fnacimiento = Now
   oTabla.IdDistritoDomicilio = ""
   oTabla.Estado = ""
   oTabla.Fbaja = "" 'Actualizado 15102014
   oTabla.DocumentoNumero = ""
   oTabla.MotivoBaja = ""
End Sub
