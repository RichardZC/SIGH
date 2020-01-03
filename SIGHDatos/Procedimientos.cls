VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Procedimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Procedimientos
'        Programado por: Garay M
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProcedimientosAgregar"
           Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMinDias = 0, Null, oTabla.EdadMinDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DescripcionOPCS", adChar, adParamInput, 10, IIf(oTabla.DescripcionOPCS = "", Null, oTabla.DescripcionOPCS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoOPCS", adChar, adParamInput, 10, IIf(oTabla.CodigoOPCS = "", Null, oTabla.CodigoOPCS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMaxDias = 0, Null, oTabla.EdadMaxDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Restriccion", adBoolean, adParamInput, 0, IIf(oTabla.Restriccion = 0, Null, oTabla.Restriccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCPT2004", adChar, adParamInput, 7, IIf(oTabla.CodigoCPT2004 = "", Null, oTabla.CodigoCPT2004)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCPT99", adChar, adParamInput, 7, IIf(oTabla.CodigoCPT99 = "", Null, oTabla.CodigoCPT99)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdProcedimiento = .Parameters("@IdProcedimiento")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Procedimientos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProcedimientosModificar"
           Set oParameter = .CreateParameter("@EdadMinDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMinDias = 0, Null, oTabla.EdadMinDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DescripcionOPCS", adChar, adParamInput, 10, IIf(oTabla.DescripcionOPCS = "", Null, oTabla.DescripcionOPCS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoOPCS", adChar, adParamInput, 10, IIf(oTabla.CodigoOPCS = "", Null, oTabla.CodigoOPCS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadMaxDias", adInteger, adParamInput, 0, IIf(oTabla.EdadMaxDias = 0, Null, oTabla.EdadMaxDias)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Restriccion", adBoolean, adParamInput, 0, IIf(oTabla.Restriccion = 0, Null, oTabla.Restriccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCPT2004", adChar, adParamInput, 7, IIf(oTabla.CodigoCPT2004 = "", Null, oTabla.CodigoCPT2004)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoCPT99", adChar, adParamInput, 7, IIf(oTabla.CodigoCPT99 = "", Null, oTabla.CodigoCPT99)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "ProcedimientosEliminar"
           Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, IIf(oTabla.IdProcedimiento = 0, Null, oTabla.IdProcedimiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub ActualizaTabla(oTabla As DOProcedimiento, oRecordset As Recordset)
       oTabla.EdadMinDias = IIf(IsNull(oRecordset!EdadMinDias), 0, oRecordset!EdadMinDias)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
       oTabla.DescripcionOPCS = IIf(IsNull(oRecordset!DescripcionOPCS), "", oRecordset!DescripcionOPCS)
       oTabla.CodigoOPCS = IIf(IsNull(oRecordset!CodigoOPCS), "", oRecordset!CodigoOPCS)
       oTabla.EdadMaxDias = IIf(IsNull(oRecordset!EdadMaxDias), 0, oRecordset!EdadMaxDias)
       oTabla.Restriccion = IIf(IsNull(oRecordset!Restriccion), 0, oRecordset!Restriccion)
       oTabla.CodigoCPT2004 = IIf(IsNull(oRecordset!CodigoCPT2004), "", oRecordset!CodigoCPT2004)
       oTabla.CodigoCPT99 = IIf(IsNull(oRecordset!CodigoCPT99), "", oRecordset!CodigoCPT99)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.IdProcedimiento = IIf(IsNull(oRecordset!IdProcedimiento), 0, oRecordset!IdProcedimiento)

End Sub


'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Procedimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOProcedimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProcedimientosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdProcedimiento", adInteger, adParamInput, 0, oTabla.IdProcedimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       ActualizaTabla oTabla, oRecordset
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProcedimientosSeleccionarTodos"
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorCodigoCPT(ByVal oTabla As DOProcedimiento) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ProcedimientosXCodigo"
        Set oParameter = .CreateParameter("@CodigoCPT2004", adVarChar, adParamInput, 7, oTabla.CodigoCPT2004): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
       SeleccionarPorCodigoCPT = True
   Else
        SeleccionarPorCodigoCPT = False
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function Filtrar(oDOProcedimiento As DOProcedimiento) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If oDOProcedimiento.CodigoCPT2004 <> "" Then
         sWhere = sWhere + " and CodigoCPT2004 like '" + oDOProcedimiento.CodigoCPT2004 + "%'"
    End If
    If oDOProcedimiento.Descripcion <> "" Then
         sWhere = sWhere + " and Descripcion like '" + oDOProcedimiento.Descripcion + "%' "
    End If
    If sWhere <> "" Then
         sSql = sSql + " where " + Mid(sWhere, 4)
    End If
    sSql = sSql + " order by CodigoCPT2004, Descripcion "
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ProcedimientosFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

