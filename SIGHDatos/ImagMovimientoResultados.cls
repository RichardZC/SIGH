VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagMovimientoResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: ImagMovimientoResultados
'        Autor: Daniel Enrique Barrantes
'        Fecha: 23/10/2018 06:15:54 p.m.
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
'   Descripción:    Inserta un registro a la tabla ImagMovimientoResultados
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOImagMovimientoResultados) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoResultadosAgregar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCPT = 0, Null, oTabla.idProductoCPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Resultado", adVarChar, adParamInput, 3000, IIf(oTabla.Resultado = "", Null, oTabla.Resultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ResultadoFecha = 0, Null, oTabla.ResultadoFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoDNI", adVarChar, adParamInput, 8, IIf(oTabla.ResultadoDNI = "", Null, oTabla.ResultadoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoRuta", adVarChar, adParamInput, 200, IIf(oTabla.EquipoRuta = "", Null, oTabla.EquipoRuta))
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
'   Descripción:    modifica un registro a la tabla ImagMovimientoResultados
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOImagMovimientoResultados) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoResultadosModificar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.idMovimiento = 0, Null, oTabla.idMovimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCPT = 0, Null, oTabla.idProductoCPT))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Resultado", adVarChar, adParamInput, 3000, IIf(oTabla.Resultado = "", Null, oTabla.Resultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ResultadoFecha = 0, Null, oTabla.ResultadoFecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ResultadoDNI", adVarChar, adParamInput, 8, IIf(oTabla.ResultadoDNI = "", Null, oTabla.ResultadoDNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EquipoRuta", adVarChar, adParamInput, 200, IIf(oTabla.EquipoRuta = "", Null, oTabla.EquipoRuta))
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
'   Descripción:    elimia o borra un registro a la tabla ImagMovimientoResultados
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOImagMovimientoResultados) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagMovimientoResultadosEliminar"
   Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla ImagMovimientoResultados
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOImagMovimientoResultados) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ImagMovimientoResultadosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.idMovimiento)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idMovimiento = IIf(IsNull(oRecordset!idMovimiento), 0, oRecordset!idMovimiento)
       oTabla.idProductoCPT = IIf(IsNull(oRecordset!idProductoCPT), 0, oRecordset!idProductoCPT)
       oTabla.Resultado = IIf(IsNull(oRecordset!Resultado), "", oRecordset!Resultado)
       oTabla.ResultadoFecha = IIf(IsNull(oRecordset!ResultadoFecha), 0, oRecordset!ResultadoFecha)
       oTabla.ResultadoDNI = IIf(IsNull(oRecordset!ResultadoDNI), "", oRecordset!ResultadoDNI)
       oTabla.EquipoRuta = IIf(IsNull(oRecordset!EquipoRuta), "", oRecordset!EquipoRuta)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOImagMovimientoResultados)
   oTabla.idMovimiento = 0
   oTabla.idProductoCPT = 0
   oTabla.Resultado = ""
   oTabla.ResultadoFecha = Now
   oTabla.ResultadoDNI = ""
   oTabla.EquipoRuta = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: ImagMovimientoResultados
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: ImagMovimientoResultados
'        Autor:   DAniel Barrantes B.
'        Fecha: 23/10/2018 06:15:55 p.m.
'------------------------------------------------------------------------------------
 

