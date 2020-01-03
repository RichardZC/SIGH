VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterLabDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterLabDetalle
'        Autor: Daniel Enrique Barrantes
'        Fecha: 12/05/2018 09:32:48 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla InterLabDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInterLabDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabDetalleAgregar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, IIf(oTabla.identificador = 0, Null, oTabla.identificador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, IIf(oTabla.cptCodigo = "", Null, oTabla.cptCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 250, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraResultado = 0, Null, oTabla.FechaHoraResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Grupo", adVarChar, adParamInput, 100, IIf(oTabla.Grupo = "", Null, oTabla.Grupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@item", adVarChar, adParamInput, 100, IIf(oTabla.item = "", Null, oTabla.item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemResultado", adVarChar, adParamInput, 500, IIf(oTabla.itemResultado = "", Null, oTabla.itemResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemReferencia", adVarChar, adParamInput, 100, IIf(oTabla.itemReferencia = "", Null, oTabla.itemReferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemMetodo", adVarChar, adParamInput, 50, IIf(oTabla.itemMetodo = "", Null, oTabla.itemMetodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.idEstado = 0, Null, oTabla.idEstado))
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
'   Descripción:    modifica un registro a la tabla InterLabDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInterLabDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabDetalleModificar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, IIf(oTabla.identificador = 0, Null, oTabla.identificador))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cptCodigo", adVarChar, adParamInput, 20, IIf(oTabla.cptCodigo = "", Null, oTabla.cptCodigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@cpt", adVarChar, adParamInput, 250, IIf(oTabla.cpt = "", Null, oTabla.cpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaHoraResultado", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaHoraResultado = 0, Null, oTabla.FechaHoraResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Grupo", adVarChar, adParamInput, 100, IIf(oTabla.Grupo = "", Null, oTabla.Grupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@item", adVarChar, adParamInput, 100, IIf(oTabla.item = "", Null, oTabla.item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemResultado", adVarChar, adParamInput, 500, IIf(oTabla.itemResultado = "", Null, oTabla.itemResultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemReferencia", adVarChar, adParamInput, 100, IIf(oTabla.itemReferencia = "", Null, oTabla.itemReferencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@itemMetodo", adVarChar, adParamInput, 50, IIf(oTabla.itemMetodo = "", Null, oTabla.itemMetodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, IIf(oTabla.idEstado = 0, Null, oTabla.idEstado))
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
'   Descripción:    elimia o borra un registro a la tabla InterLabDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInterLabDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InterLabDetalleEliminar"
   Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, oTabla.identificador)
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
'   Descripción:    Selecciona un unico registro de la tabla InterLabDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInterLabDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InterLabDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@identificador", adInteger, adParamInput, 0, oTabla.identificador)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.identificador = IIf(IsNull(oRecordset!identificador), 0, oRecordset!identificador)
       oTabla.cptCodigo = IIf(IsNull(oRecordset!cptCodigo), "", oRecordset!cptCodigo)
       oTabla.cpt = IIf(IsNull(oRecordset!cpt), "", oRecordset!cpt)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.FechaHoraResultado = IIf(IsNull(oRecordset!FechaHoraResultado), 0, oRecordset!FechaHoraResultado)
       oTabla.Grupo = IIf(IsNull(oRecordset!Grupo), "", oRecordset!Grupo)
       oTabla.item = IIf(IsNull(oRecordset!item), "", oRecordset!item)
       oTabla.itemResultado = IIf(IsNull(oRecordset!itemResultado), "", oRecordset!itemResultado)
       oTabla.itemReferencia = IIf(IsNull(oRecordset!itemReferencia), "", oRecordset!itemReferencia)
       oTabla.itemMetodo = IIf(IsNull(oRecordset!itemMetodo), "", oRecordset!itemMetodo)
       oTabla.idEstado = IIf(IsNull(oRecordset!idEstado), 0, oRecordset!idEstado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOInterLabDetalle)
   oTabla.identificador = 0
   oTabla.cptCodigo = ""
   oTabla.cpt = ""
   oTabla.DNI = ""
   oTabla.FechaHoraResultado = Now
   oTabla.Grupo = ""
   oTabla.item = ""
   oTabla.itemResultado = ""
   oTabla.itemReferencia = ""
   oTabla.itemMetodo = ""
   oTabla.idEstado = 0
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: InterLabDetalle
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: InterLabDetalle
'        Autor:   DAniel Barrantes B.
'        Fecha: 12/05/2018 09:32:48 a.m.
'------------------------------------------------------------------------------------
 

