VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 02/06/2005 [Todo el archivo]
'WCG01
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaCaja
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:57
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla CajaCaja
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
    '**** Programa: se agrego @ImpresoraDefault y @Impresora2
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014

Function Insertar(ByVal oTabla As DOCajaCaja) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajaAgregar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoginPc", adVarChar, adParamInput, 50, IIf(oTabla.loginPc = "", Null, oTabla.loginPc)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImpresoraDefault", adVarChar, adParamInput, 50, IIf(oTabla.ImpresoraDefault = "", Null, oTabla.ImpresoraDefault)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Impresora2", adVarChar, adParamInput, 50, IIf(oTabla.Impresora2 = "", Null, oTabla.Impresora2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SerieImpresoraDefault", adVarChar, adParamInput, 20, IIf(oTabla.SerieImpresoraDefault = "", Null, oTabla.SerieImpresoraDefault)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SerieImpresora2", adVarChar, adParamInput, 20, IIf(oTabla.SerieImpresora2 = "", Null, oTabla.SerieImpresora2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormatoImp2Cinta", adInteger, adParamInput, 0, IIf(oTabla.FormatoImp2Cinta = 0, Null, oTabla.FormatoImp2Cinta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormatoImpDefaultCinta", adInteger, adParamInput, 0, IIf(oTabla.FormatoImpDefaultCinta = 0, Null, oTabla.FormatoImpDefaultCinta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante2", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante2 = 0, Null, oTabla.IdTipoComprobante2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
       oTabla.IdCaja = .Parameters("@IdCaja")
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
'   Descripción:    modifica un registro a la tabla CajaCaja
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

    '**** Programa: se agrego @ImpresoraDefault y @Impresora2
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014
Function Modificar(ByVal oTabla As DOCajaCaja) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajaModificar"
           Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LoginPc", adVarChar, adParamInput, 50, IIf(oTabla.loginPc = "", Null, oTabla.loginPc)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ImpresoraDefault", adVarChar, adParamInput, 50, IIf(oTabla.ImpresoraDefault = "", Null, oTabla.ImpresoraDefault)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Impresora2", adVarChar, adParamInput, 50, IIf(oTabla.Impresora2 = "", Null, oTabla.Impresora2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SerieImpresoraDefault", adVarChar, adParamInput, 50, IIf(oTabla.SerieImpresoraDefault = "", Null, oTabla.SerieImpresoraDefault)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SerieImpresora2", adVarChar, adParamInput, 50, IIf(oTabla.SerieImpresora2 = "", Null, oTabla.SerieImpresora2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormatoImp2Cinta", adInteger, adParamInput, 0, IIf(oTabla.FormatoImp2Cinta = 0, Null, oTabla.FormatoImp2Cinta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormatoImpDefaultCinta", adInteger, adParamInput, 0, IIf(oTabla.FormatoImpDefaultCinta = 0, Null, oTabla.FormatoImpDefaultCinta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoComprobante2", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante2 = 0, Null, oTabla.IdTipoComprobante2)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaCaja
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaCaja) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaCajaEliminar"
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaCaja
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
    '**** Programa: se agrego ImpresoraDefault y Impresora2
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014
Function SeleccionarPorId(ByVal oTabla As DOCajaCaja) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaCajaSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, oTabla.IdCaja): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.Descripcion = Trim(IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion))
       oTabla.codigo = Trim(IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo))
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.loginPc = Trim(IIf(IsNull(oRecordset!loginPc), "", oRecordset!loginPc))
       oTabla.ImpresoraDefault = Trim(IIf(IsNull(oRecordset!ImpresoraDefault), "", oRecordset!ImpresoraDefault))
       oTabla.Impresora2 = Trim(IIf(IsNull(oRecordset!Impresora2), "", oRecordset!Impresora2))
       oTabla.SerieImpresoraDefault = Trim(IIf(IsNull(oRecordset!SerieImpresoraDefault), "", oRecordset!SerieImpresoraDefault))
       oTabla.SerieImpresora2 = Trim(IIf(IsNull(oRecordset!SerieImpresora2), "", oRecordset!SerieImpresora2))
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.IdTipoComprobante2 = IIf(IsNull(oRecordset!IdTipoComprobante2), 0, oRecordset!IdTipoComprobante2)
       oTabla.FormatoImp2Cinta = IIf(IsNull(oRecordset!FormatoImp2Cinta), 0, oRecordset!FormatoImp2Cinta)
       oTabla.FormatoImpDefaultCinta = IIf(IsNull(oRecordset!FormatoImpDefaultCinta), 0, oRecordset!FormatoImpDefaultCinta)
       oTabla.IdPartida = IIf(IsNull(oRecordset!IdPartida), 0, oRecordset!IdPartida)
       oTabla.IdCentroCosto = IIf(IsNull(oRecordset!IdCentroCosto), 0, oRecordset!IdCentroCosto)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""

Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function ObtenerSiguienteNumeroDocumento(oTabla As DOCajaNroDocumento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lNuevoNumero As Long
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaNroDocumentoXidCajaTipoComprobante"
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, oTabla.IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, oTabla.IdTipoComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerSiguienteNumeroDocumento = False
   If Not (oRecordset.EOF And oRecordset.BOF) Then
         oTabla.NroSerie = Trim(IIf(IsNull(oRecordset!NroSerie), "0", oRecordset!NroSerie))
         oTabla.NroDocumento = Trim(IIf(IsNull(oRecordset!NroDocumento), "0", oRecordset!NroDocumento))
         lNuevoNumero = Val(oTabla.NroDocumento) + 1
         oTabla.NroDocumento = Right("0000000000" & lNuevoNumero, IIf(oTabla.NroDocumento = "", 6, Len(oTabla.NroDocumento)))
         ObtenerSiguienteNumeroDocumento = True
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerSiguienteNumeroDocumentoYgrabarlo(ByVal oTabla As DOCajaNroDocumento) As Boolean
    Dim oCajaNroDocumento As New SIGHDatos.CajaNroDocumento
    Dim lnItem As Integer
    On Error GoTo ManejaError
    ObtenerSiguienteNumeroDocumentoYgrabarlo = False
    Set oCajaNroDocumento.Conexion = mo_Conexion
    If oCajaNroDocumento.SeleccionarPorId(oTabla) Then
        lnItem = Len(Trim(oTabla.NroDocumento))
        oTabla.NroDocumento = Right("0000000000" + Trim(Str(Val(oTabla.NroDocumento) + 1)), lnItem)
        If oCajaNroDocumento.Modificar(oTabla) Then
           ObtenerSiguienteNumeroDocumentoYgrabarlo = True
        End If
    End If
    Exit Function
ManejaError:
    ms_MensajeError = Err.Number & " " + Err.Description
End Function


Function RealizarFiltro(oCaja As DOCajaCaja) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

   SQL = ""
   If oCaja.codigo <> "" Then
        SQL = SQL + " c.codigo like '%" & oCaja.codigo & "%' and"
   End If
   If oCaja.Descripcion <> "" Then
        SQL = SQL + " c.descripcion like '%" & oCaja.Descripcion & "%' and"
   End If
    If SQL <> "" Then
       SQL = Left(SQL, Len(SQL) - 4)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaCajaSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, SQL): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set RealizarFiltro = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function SeleccionarTodosParaLista() As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CajaCajaSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodosParaLista = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



