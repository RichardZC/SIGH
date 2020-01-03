VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuentesFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FuentesFinanciamiento
'        Programado por: Barrantes D
'        Fecha: Julio 2009
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
'   Descripción:    Inserta un registro a la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFuenteFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FuentesFinanciamientoAgregar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oTabla.idFuenteFinanciamiento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoConceptoFarmacia", adInteger, adParamInput, 0, oTabla.idTipoConceptoFarmacia): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UtilizadoEn", adInteger, adParamInput, 0, oTabla.UtilizadoEn): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoFuenteFinanciamientoSEM", adVarChar, adParamInput, 2, oTabla.CodigoFuenteFinanciamientoSEM): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idAreaTramitaSeguros", adInteger, adParamInput, 0, oTabla.idAreaTramitaSeguros): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsUsadoEnCaja", adBoolean, adParamInput, 0, oTabla.EsUsadoEnCaja): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 2, oTabla.CodigoHIS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciador", adInteger, adParamInput, 0, oTabla.idTipoFinanciador): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 11, oTabla.codigo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneEPS", adInteger, adParamInput, 0, oTabla.TieneEPS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@usadoEnFUnidosis", adInteger, adParamInput, 0, oTabla.usadoEnFUnidosis): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CuentaParaUnidosis", adInteger, adParamInput, 0, oTabla.CuentaParaUnidosis): .Parameters.Append oParameter
            .Execute
           oTabla.idFuenteFinanciamiento = .Parameters("@IdFuenteFinanciamiento")
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
'   Descripción:    modifica un registro a la tabla FuentesFinanciamiento
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFuenteFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FuentesFinanciamientoModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 50, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoConceptoFarmacia", adInteger, adParamInput, 0, oTabla.idTipoConceptoFarmacia): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UtilizadoEn", adInteger, adParamInput, 0, oTabla.UtilizadoEn): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoFuenteFinanciamientoSEM", adVarChar, adParamInput, 2, oTabla.CodigoFuenteFinanciamientoSEM): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idAreaTramitaSeguros", adInteger, adParamInput, 0, oTabla.idAreaTramitaSeguros): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsUsadoEnCaja", adBoolean, adParamInput, 0, oTabla.EsUsadoEnCaja): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CodigoHIS", adVarChar, adParamInput, 2, oTabla.CodigoHIS): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdTipoFinanciador", adInteger, adParamInput, 0, oTabla.idTipoFinanciador): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 11, oTabla.codigo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TieneEPS", adInteger, adParamInput, 0, oTabla.TieneEPS): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@usadoEnFUnidosis", adInteger, adParamInput, 0, oTabla.usadoEnFUnidosis): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CuentaParaUnidosis", adInteger, adParamInput, 0, oTabla.CuentaParaUnidosis): .Parameters.Append oParameter
       
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
'   Descripción:    elimia o borra un registro a la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFuenteFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FuentesFinanciamientoEliminar"
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FuentesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFuenteFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FuentesFinanciamientoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oTabla.idFuenteFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.idTipoConceptoFarmacia = IIf(IsNull(oRecordset!idTipoConceptoFarmacia), 0, oRecordset!idTipoConceptoFarmacia)
       oTabla.UtilizadoEn = IIf(IsNull(oRecordset!UtilizadoEn), 0, oRecordset!UtilizadoEn)
       oTabla.CodigoFuenteFinanciamientoSEM = IIf(IsNull(oRecordset!CodigoFuenteFinanciamientoSEM), "", oRecordset!CodigoFuenteFinanciamientoSEM)
       oTabla.idAreaTramitaSeguros = IIf(IsNull(oRecordset!idAreaTramitaSeguros), 0, oRecordset!idAreaTramitaSeguros)
       oTabla.EsUsadoEnCaja = IIf(IsNull(oRecordset!EsUsadoEnCaja), False, oRecordset!EsUsadoEnCaja)
       oTabla.CodigoHIS = IIf(IsNull(oRecordset!CodigoHIS), "", oRecordset!CodigoHIS)
       oTabla.idTipoFinanciador = IIf(IsNull(oRecordset!idTipoFinanciador), 0, oRecordset!idTipoFinanciador)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.TieneEPS = IIf(IsNull(oRecordset!TieneEPS), 0, oRecordset!TieneEPS)
       oTabla.usadoEnFUnidosis = IIf(IsNull(oRecordset!usadoEnFUnidosis), 0, oRecordset!usadoEnFUnidosis)
       oTabla.CuentaParaUnidosis = IIf(IsNull(oRecordset!CuentaParaUnidosis), 0, oRecordset!CuentaParaUnidosis)
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
Function SeleccionarPorTipoFinanciamiento(lIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoFinanciamiento = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FuentesFinanciamientoSeleccionarPorTipoFinanciamiento"
       Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, lIdFuenteFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarPorTipoFinanciamiento = oRecordset
 
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
     .CommandText = "FuentesFinanciamientoSeleccionarTodos"
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

Function SeleccionarPorIdentificador(LnIdFuenteFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   Set SeleccionarPorIdentificador = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FuentesFinanciamientoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, LnIdFuenteFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
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


Function FuentesFinanciamientoSegunFiltro(lcFiltro As String) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FuentesFinanciamientoSegunFiltro"
       Set oParameter = .CreateParameter("@lcfiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoSegunFiltro = oRecordset
   ms_MensajeError = ""
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorCodigo(sCodigo As String) As Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "FuentesFinanciamientoSeleccionarPorCodigo"
        Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 11, sCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionarPorCodigo = oRecordset
    ms_MensajeError = ""
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
