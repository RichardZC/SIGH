VERSION 1.0 CLASS
BEGIN
   MultiUse = -1  'True
   Persistable = 0  'NotPersistable
   DataBindingBehavior = 0  'vbNone
   DataSourceBehavior = 0   'vbNone
   MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "LabResultado"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: LabResultado
'        Autor: Adams BONILLA MAGALLANES
'        Fecha: 25/06/2009 11:24:46 a.m.
'        Auspiciado por: Hospital Regional Ayacucho
'        Control De Cambios: 
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String

Property Set Conexion(oValue as ADODB.Connection)
  set mo_Conexion = oValue
End Property

Property Get Conexion() as ADODB.Connection
  set Conexion = mo_Conexion
End Property

Property Let MensajeError(sValue as String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() as String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla LabResultado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOLabResultado) as Boolean
  On Error goto ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Insertar = False 
  With oCommand
    .CommandType = adCmdStoredProc
    set .ActiveConnection = mo_Conexion
    .CommandText = "LabResultadoAgregar"
    Set oParameter = .CreateParameter("@idAnalisis",adInteger,adParamInput,0,IIf(oTabla.idAnalisis=0, Null,oTabla.idAnalisis))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idOrden",adInteger,adParamInput,0,IIf(oTabla.idOrden=0, Null,oTabla.idOrden))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@resultadoAnalisis",adLongVarChar,adParamInput,2147483647,oTabla.resultadoAnalisis)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@observacionResultado",adLongVarChar,adParamInput,2147483647,oTabla.observacionResultado)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idUsuario",adInteger,adParamInput,0,IIf(oTabla.idUsuario=0, Null,oTabla.idUsuario))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdUsuarioAuditoria",adInteger,adParamInput,0,oTabla.IdUsuarioAuditoria)
   . Parameters.Append oParameter
    .Execute
  End With
  Insertar = True 
  ms_MensajeError=""
  Exit Function

ManejadorDeError:
  ms_MensajeError=Err.Number & " " + Err.Description
  Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    Modifica un registro de la tabla LabResultado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Modificar(ByVal oTabla As DOLabResultado) as Boolean
  On Error goto ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Modificar = False 
  With oCommand
    .CommandType = adCmdStoredProc
    set .ActiveConnection = mo_Conexion
    .CommandText = "LabResultadoModificar"
    Set oParameter = .CreateParameter("@idAnalisis",adInteger,adParamInput,0,IIf(oTabla.idAnalisis=0, Null,oTabla.idAnalisis))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idOrden",adInteger,adParamInput,0,IIf(oTabla.idOrden=0, Null,oTabla.idOrden))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@resultadoAnalisis",adLongVarChar,adParamInput,2147483647,oTabla.resultadoAnalisis)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@observacionResultado",adLongVarChar,adParamInput,2147483647,oTabla.observacionResultado)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idUsuario",adInteger,adParamInput,0,IIf(oTabla.idUsuario=0, Null,oTabla.idUsuario))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdUsuarioAuditoria",adInteger,adParamInput,0,oTabla.IdUsuarioAuditoria)
    .Parameters.Append oParameter
    .Execute
  End With
  Modificar = True 
  ms_MensajeError=""
Exit Function

ManejadorDeError:
  ms_MensajeError=Err.Number & " " + Err.Description
  Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    Elimina o borra un registro de la tabla LabResultado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Eliminar(ByVal oTabla As DOLabResultado) as Boolean
  On Error goto ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Eliminar = False 
  With oCommand
    .CommandType = adCmdStoredProc
    set .ActiveConnection = mo_Conexion
    .CommandText = "LabResultadoEliminar"
    Set oParameter = .CreateParameter("@idAnalisis",adInteger,adParamInput,0,oTabla.idAnalisis)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdUsuarioAuditoria",adInteger,adParamInput,0,oTabla.IdUsuarioAuditoria)
    .Parameters.Append oParameter
    .Execute
  End With
  Eliminar = True 
  ms_MensajeError=""
  Exit Function

ManejadorDeError:
  ms_MensajeError=Err.Number & " " + Err.Description
  Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un único registro de la tabla LabResultado
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DOLabResultado) as Boolean
  On Error goto ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  SeleccionarPorId = False 
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabResultadoSeleccionarPorId"
    Set oParameter = .CreateParameter("@idAnalisis",adInteger,adParamInput,0,oTabla.idAnalisis)
    .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  If Not (oRecordset.EOF And oRecordset.BOF) Then
    oTabla.idAnalisis = IIf(isnull(oRecordset!idAnalisis),0,oRecordset!idAnalisis)
    oTabla.idOrden = IIf(isnull(oRecordset!idOrden),0,oRecordset!idOrden)
    oTabla.resultadoAnalisis=oRecordset!resultadoAnalisis
    oTabla.observacionResultado=oRecordset!observacionResultado
    oTabla.idUsuario = IIf(isnull(oRecordset!idUsuario),0,oRecordset!idUsuario)
    SeleccionarPorId = True 
    ms_MensajeError=""
  End If
  oRecordset.Close
  Exit Function

ManejadorDeError:
  ms_MensajeError=Err.Number & " " + Err.Description
  Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOLabResultado)
  oTabla.idAnalisis = 0
  oTabla.idOrden = 0
  oTabla.resultadoAnalisis = ""
  oTabla.observacionResultado = ""
  oTabla.idUsuario = 0
End Sub

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: LabResultado
'        Autor:   Adams BONILLA MAGALLANES  
'------------------------------------------------------------------------------------
 
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: LabResultado
'        Autor:   Adams BONILLA MAGALLANES 
'        Fecha: 25/06/2009 11:24:47 a.m.
'------------------------------------------------------------------------------------
 
