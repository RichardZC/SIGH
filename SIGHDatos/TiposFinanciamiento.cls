VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiposFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TiposFinanciamiento
'        Programado por: Castro W
'        Fecha: Marzo 2006
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
'   Descripción:    Inserta un registro a la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOTiposFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposFinanciamientoAgregar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oTabla.idTipoFinanciamiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esOficina", adBoolean, adParamInput, 0, oTabla.esOficina): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esSalida", adBoolean, adParamInput, 0, oTabla.esSalida): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SeIngresPrecios", adBoolean, adParamInput, 0, oTabla.SeIngresPrecios): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsFarmacia", adBoolean, adParamInput, 0, oTabla.esFArmacia): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCajaTiposComprobante", adInteger, adParamInput, 0, IIf(oTabla.idCajaTiposComprobante = 0, Null, oTabla.idCajaTiposComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@tipoVenta", adVarChar, adParamInput, 1, IIf(oTabla.tipoVenta = "", Null, oTabla.tipoVenta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SeImprimeComprobante", adBoolean, adParamInput, 0, IIf(oTabla.SeImprimeComprobante = 0, Null, oTabla.SeImprimeComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esFuenteFinanciamiento", adBoolean, adParamInput, 0, oTabla.esFuenteFinanciamiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@GeneraPago", adInteger, adParamInput, 0, oTabla.GeneraPago): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, IIf(oTabla.idTipoConcepto = 0, Null, oTabla.idTipoConcepto)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.idTipoFinanciamiento = .Parameters("@IdTipoFinanciamiento")
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
'   Descripción:    modifica un registro a la tabla TiposFinanciamiento
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTiposFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposFinanciamientoModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esOficina", adBoolean, adParamInput, 0, oTabla.esOficina): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esSalida", adBoolean, adParamInput, 0, oTabla.esSalida): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SeIngresPrecios", adBoolean, adParamInput, 0, oTabla.SeIngresPrecios): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsFarmacia", adBoolean, adParamInput, 0, oTabla.esFArmacia): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCajaTiposComprobante", adInteger, adParamInput, 0, IIf(oTabla.idCajaTiposComprobante = 0, Null, oTabla.idCajaTiposComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@tipoVenta", adVarChar, adParamInput, 1, IIf(oTabla.tipoVenta = "", Null, oTabla.tipoVenta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SeImprimeComprobante", adBoolean, adParamInput, 0, IIf(oTabla.SeImprimeComprobante = 0, Null, oTabla.SeImprimeComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@esFuenteFinanciamiento", adBoolean, adParamInput, 0, oTabla.esFuenteFinanciamiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@GeneraPago", adInteger, adParamInput, 0, oTabla.GeneraPago): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoConcepto", adInteger, adParamInput, 0, IIf(oTabla.idTipoConcepto = 0, Null, oTabla.idTipoConcepto)): .Parameters.Append oParameter
       
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
'   Descripción:    elimia o borra un registro a la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTiposFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "TiposFinanciamientoEliminar"
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
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

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla TiposFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOTiposFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TiposFinanciamientoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oTabla.idTipoFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.esOficina = IIf(IsNull(oRecordset!esOficina), 0, oRecordset!esOficina)
       oTabla.esSalida = IIf(IsNull(oRecordset!esSalida), 0, oRecordset!esSalida)
       oTabla.SeIngresPrecios = IIf(IsNull(oRecordset!SeIngresPrecios), 0, oRecordset!SeIngresPrecios)
       oTabla.esFArmacia = IIf(IsNull(oRecordset!esFArmacia), 0, oRecordset!esFArmacia)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.tipoVenta = IIf(IsNull(oRecordset!tipoVenta), "", oRecordset!tipoVenta)
       oTabla.SeImprimeComprobante = IIf(IsNull(oRecordset!SeImprimeComprobante), 0, oRecordset!SeImprimeComprobante)
       oTabla.esFuenteFinanciamiento = IIf(IsNull(oRecordset!esFuenteFinanciamiento), 0, oRecordset!esFuenteFinanciamiento)
       oTabla.GeneraPago = IIf(IsNull(oRecordset!GeneraPago), 0, oRecordset!GeneraPago)
       oTabla.idTipoConcepto = IIf(IsNull(oRecordset!idTipoConcepto), 0, oRecordset!idTipoConcepto)
       oTabla.idCajaTiposComprobante = IIf(IsNull(oRecordset!idCajaTiposComprobante), 0, oRecordset!idCajaTiposComprobante)
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
     .CommandText = "TiposFinanciamientoSeleccionarTodos"
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


Function SeleccionarParaCaja() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarParaCaja = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TiposFinanciamientoSeleccionarParaCaja"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarParaCaja = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SeleccionarPorIdentificador(LnIdTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   Set SeleccionarPorIdentificador = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TiposFinanciamientoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, LnIdTipoFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
     'Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdentificador = oRecordset
   ms_MensajeError = ""
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function TiposFinanciamientoSegunFiltro(lcFiltro As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientoSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set TiposFinanciamientoSegunFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
