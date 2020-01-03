VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmAlmacen
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_idUsuario
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmAlmacen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmAlmacen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmAlmacenAgregar"
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, oTabla.idAlmacen)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoLocales", adVarChar, adParamInput, 1, IIf(oTabla.idTipoLocales = "", Null, oTabla.idTipoLocales))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSuministro", adVarChar, adParamInput, 2, IIf(oTabla.idTipoSuministro = "", Null, oTabla.idTipoSuministro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, oTabla.IdEstado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@codigoSISMED", adVarChar, adParamInput, 11, oTabla.codigoSISMED)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarDias", adVarChar, adParamInput, 7, oTabla.regenerarDias)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarHora", adVarChar, adParamInput, 5, oTabla.regenerarHora)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarEstado", adVarChar, adParamInput, 7, IIf(Trim(oTabla.regenerarEstado) = "", Null, oTabla.regenerarEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@esUnidosis", adInteger, adParamInput, 0, oTabla.esUnidosis)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idAlmacen = .Parameters("@idAlmacen")
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
'   Descripción:    modifica un registro a la tabla farmAlmacen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Modificar(ByVal oTabla As DOfarmAlmacen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmAlmacenModificar"
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, IIf(oTabla.idAlmacen = 0, Null, oTabla.idAlmacen))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoLocales", adVarChar, adParamInput, 1, IIf(oTabla.idTipoLocales = "", Null, oTabla.idTipoLocales))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSuministro", adVarChar, adParamInput, 2, IIf(oTabla.idTipoSuministro = "", Null, oTabla.idTipoSuministro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, oTabla.IdEstado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@codigoSISMED", adVarChar, adParamInput, 11, oTabla.codigoSISMED)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarDias", adVarChar, adParamInput, 7, oTabla.regenerarDias)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarHora", adVarChar, adParamInput, 5, oTabla.regenerarHora)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@regenerarEstado", adVarChar, adParamInput, 7, IIf(Trim(oTabla.regenerarEstado) = "", Null, oTabla.regenerarEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@esUnidosis", adInteger, adParamInput, 0, oTabla.esUnidosis)
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
'   Descripción:    elimia o borra un registro a la tabla farmAlmacen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmAlmacen) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmAlmacenEliminar"
   Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, oTabla.idAlmacen)
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
'   Descripción:    Selecciona un unico registro de la tabla farmAlmacen
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmAlmacen) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmAlmacenSeleccionarPorId"
     Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, oTabla.idAlmacen)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAlmacen = IIf(IsNull(oRecordset!idAlmacen), 0, oRecordset!idAlmacen)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.idTipoLocales = IIf(IsNull(oRecordset!idTipoLocales), "", oRecordset!idTipoLocales)
       oTabla.idTipoSuministro = IIf(IsNull(oRecordset!idTipoSuministro), "", oRecordset!idTipoSuministro)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.codigoSISMED = IIf(IsNull(oRecordset!codigoSISMED), 0, oRecordset!codigoSISMED)
       oTabla.regenerarDias = IIf(IsNull(oRecordset!regenerarDias), "", oRecordset!regenerarDias)
       oTabla.regenerarEstado = IIf(IsNull(oRecordset!regenerarEstado), "", oRecordset!regenerarEstado)
       oTabla.regenerarHora = IIf(IsNull(oRecordset!regenerarHora), "", oRecordset!regenerarHora)
       oTabla.esUnidosis = IIf(IsNull(oRecordset!esUnidosis), 0, oRecordset!esUnidosis)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmAlmacen)
   oTabla.idAlmacen = 0
   oTabla.Descripcion = ""
   oTabla.idTipoLocales = ""
   oTabla.idTipoSuministro = ""
   oTabla.IdEstado = 0
End Sub


Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FarmAlmacenSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarTodosMenosExternos() As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "farmAlmacenSeleccionarTodosMenosExternos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodosMenosExternos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarSegunFiltro(lcFiltro As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FarmAlmacenFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarSegunFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function CargaUltimoCorrelativoIdAlmacen() As Long

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FarmAlmacenUltimoAlmacen"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount > 0 Then
       CargaUltimoCorrelativoIdAlmacen = oRecordset.Fields!idAlmacen + 1
   End If
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

