VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_Lotes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_Lotes
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
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
'   Descripción:    Inserta un registro a la tabla HIS_Lotes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_Lotes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_LotesAgregar"
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adChar, adParamInput, 4, IIf(oTabla.Lote = "", Null, oTabla.Lote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHojas", adInteger, adParamInput, 0, IIf(oTabla.NroHojas = 0, 0, oTabla.NroHojas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, IIf(oTabla.Anio = 0, Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoLote", adInteger, adParamInput, 0, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cerrado", adInteger, adParamInput, 0, IIf(oTabla.cerrado = 0, 0, oTabla.cerrado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DobleDigitacion", adInteger, adParamInput, 0, IIf(oTabla.DobleDigitacion = 0, 0, oTabla.DobleDigitacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisLote = .Parameters("@IdHisLote")
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
'   Descripción:    modifica un registro a la tabla HIS_Lotes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_Lotes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_LotesModificar"
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, IIf(oTabla.IdHisLote = 0, Null, oTabla.IdHisLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adChar, adParamInput, 4, oTabla.Lote)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHojas", adInteger, adParamInput, 0, IIf(oTabla.NroHojas = 0, Null, oTabla.NroHojas))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, IIf(oTabla.Anio = 0, Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoLote", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoLote = 0, 0, oTabla.IdEstadoLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cerrado", adInteger, adParamInput, 0, IIf(oTabla.cerrado = 0, 0, oTabla.cerrado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DobleDigitacion", adInteger, adParamInput, 0, IIf(oTabla.DobleDigitacion = 0, 0, oTabla.DobleDigitacion))
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

Function ModificarEstadoLote(ByVal oTabla As DOHIS_Lotes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarEstadoLote = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_ModificarEstadoLote"
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, IIf(oTabla.IdHisLote = 0, Null, oTabla.IdHisLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoLote", adInteger, adParamInput, 0, oTabla.IdEstadoLote)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarEstadoLote = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla HIS_Lotes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_Lotes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_LotesEliminar"
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, oTabla.IdHisLote)
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

Sub ActualizaTabla(ByRef oTabla As DOHIS_Lotes, oRecordset As Recordset)
       oTabla.IdHisLote = IIf(IsNull(oRecordset!IdHisLote), 0, oRecordset!IdHisLote)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.Lote = IIf(IsNull(oRecordset!Lote), "", oRecordset!Lote)
       oTabla.NroHojas = IIf(IsNull(oRecordset!NroHojas), 0, oRecordset!NroHojas)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), 0, oRecordset!Mes)
       oTabla.Anio = IIf(IsNull(oRecordset!Anio), 0, oRecordset!Anio)
       oTabla.IdEstadoLote = IIf(IsNull(oRecordset!IdEstadoLote), 0, oRecordset!IdEstadoLote)
       oTabla.DobleDigitacion = IIf(IsNull(oRecordset!DobleDigitacion), 0, oRecordset!DobleDigitacion)
       oTabla.cerrado = IIf(IsNull(oRecordset!cerrado), 0, oRecordset!cerrado)
    
End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_Lotes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_Lotes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_LotesSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, oTabla.IdHisLote)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ActualizaTabla oTabla, oRecordset
'       oTabla.IdHisLote = IIf(IsNull(oRecordSet!IdHisLote), 0, oRecordSet!IdHisLote)
'       oTabla.IdEstablecimiento = IIf(IsNull(oRecordSet!IdEstablecimiento), 0, oRecordSet!IdEstablecimiento)
'       oTabla.Lote = IIf(IsNull(oRecordSet!Lote), "", oRecordSet!Lote)
'       oTabla.NroHojas = IIf(IsNull(oRecordSet!NroHojas), 0, oRecordSet!NroHojas)
'       oTabla.Mes = IIf(IsNull(oRecordSet!Mes), 0, oRecordSet!Mes)
'       oTabla.Anio = IIf(IsNull(oRecordSet!Anio), 0, oRecordSet!Anio)
'       oTabla.cerrado = IIf(IsNull(oRecordSet!cerrado), 0, oRecordSet!cerrado)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_Lotes)
   oTabla.IdHisLote = 0
   oTabla.IdEstablecimiento = 0
   oTabla.Lote = ""
   oTabla.NroHojas = 0
   oTabla.Mes = 0
   oTabla.Anio = 0
   oTabla.IdEstadoLote = 0
   oTabla.cerrado = 0
End Sub

Function ConsultarRegistrosLotesPorIdEstablec(IdEstablecimiento As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
'Dim oRecordSet As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSql As String
'
'sSql = "SELECT [IdHisLote], [IdEstablecimiento], [Lote], [NroHojas], [Mes], [Anio], [Cerrado] FROM [HIS_Lotes]"
'sSql = sSql & " WHERE IdEstablecimiento = " & IdEstablecimiento
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSql
'    Set oRecordSet = .Execute
'    Set oRecordSet.ActiveConnection = Nothing
'End With
'
'Set ConsultarRegistrosLotesPorIdEstablec = oRecordSet
'ms_MensajeError = ""
'Exit Function
'ManejadorDeError:
'       ms_MensajeError = Err.Number & " " + Err.Description
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesConsultarRegistrosLotesPorIdEstablec"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ConsultarRegistrosLotesPorIdEstablec = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ValidarLoteHIS_LoteExiste(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesValidarLoteHIS_LoteExiste"
        Set oParameter = .CreateParameter("IdEstablecimiento", adInteger, adParamInput, 0, oTablaDOHIS_Lote.IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 4, oTablaDOHIS_Lote.Lote): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroHojas", adInteger, adParamInput, 0, oTablaDOHIS_Lote.NroHojas): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, oTablaDOHIS_Lote.Mes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, oTablaDOHIS_Lote.Anio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        oRecordset.MoveFirst
        If CInt(oRecordset.Fields(0)) = 0 Then
            ValidarLoteHIS_LoteExiste = False
        Else
            ValidarLoteHIS_LoteExiste = True
        End If
   Else
        ValidarLoteHIS_LoteExiste = False
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ValidarLoteHIS_UltimoLoteNoCerrado(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesValidarLoteHIS_UltimoLoteNoCerrado"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, oTablaDOHIS_Lote.IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, oTablaDOHIS_Lote.Anio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, oTablaDOHIS_Lote.Mes): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        oRecordset.MoveFirst
        Do While Not oRecordset.EOF
            cerrado = CInt(oRecordset!cerrado)
            Exit Do
        Loop
        If cerrado > 0 Then
            ValidarLoteHIS_UltimoLoteNoCerrado = True
        Else
            ValidarLoteHIS_UltimoLoteNoCerrado = False
        End If
   Else
        ValidarLoteHIS_UltimoLoteNoCerrado = False
   End If

   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ValidarLoteHIS_UltimoLoteCerrado(ml_IdEstablecimientoActual As Long) As Boolean
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesValidarLoteHIS_UltimoLoteCerrado"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, ml_IdEstablecimientoActual): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        Do While Not oRecordset.EOF
            If CInt(oRecordset!cerrado) = 0 Then
                ValidarLoteHIS_UltimoLoteCerrado = False
            Else
                ValidarLoteHIS_UltimoLoteCerrado = True
            End If
        Loop
   Else
        ValidarLoteHIS_UltimoLoteCerrado = False
   End If

   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ValidarLoteHIS_PasoMaximoPaginasPermitido(ml_IdEstablecimientoActual As Long) As Boolean
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim MaxHojas As Integer
    Dim NumHojasIngresar As Integer
    Dim IdLote As Long
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesPasoMaximoPaginasPermitido"
        Set oParameter = .CreateParameter("@ml_IdEstablecimientoActual", adInteger, adParamInput, 0, ml_IdEstablecimientoActual): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        oRecordset.MoveFirst
        IdLote = CLng(oRecordset.Fields(0))
        MaxHojas = CInt(oRecordset.Fields(1))
    
        Set oCommand = Nothing
        Set oParameter = Nothing
        oRecordset.Close
        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = mo_Conexion
             .CommandTimeout = 150
             .CommandText = "His_cabeceraXlote"
             Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IdLote): .Parameters.Append oParameter
             Set oRecordset = .Execute
             Set oRecordset.ActiveConnection = Nothing
        End With
    
        NumHojasIngresar = CInt(oRecordset.Fields(0))
    
        If NumHojasIngresar > MaxHojas Then
            ValidarLoteHIS_PasoMaximoPaginasPermitido = True
        Else
            ValidarLoteHIS_PasoMaximoPaginasPermitido = False
        End If
   Else
        'caso en que comience a iniciar un Nuevo Lote
        ValidarLoteHIS_PasoMaximoPaginasPermitido = False
   End If
   
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerDatosLotePorEstablecimiento(ml_IdEstablecimientoActual As Long) As DOHIS_Lotes
    Stop
    On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim mo_DOHIS_Lotes As DOHIS_Lotes
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "HIS_LotesXIdEstablecimiento"
            Set oParameter = .CreateParameter("@ml_IdEstablecimientoActual", adInteger, adParamInput, 0, ml_IdEstablecimientoActual): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       If oRecordset.RecordCount > 0 Then
          ActualizaTabla mo_DOHIS_Lotes, oRecordset
       End If
       Set ObtenerDatosLotePorEstablecimiento = mo_DOHIS_Lotes
       Set oRecordset = Nothing
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function

Function ObtenerDatosLoteNroHojaLibre(ml_IdLote As Long) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim mo_DOHIS_Lotes As DOHIS_Lotes
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_CabeceraObtenerDatosLoteNroHojaLibre"
        Set oParameter = .CreateParameter("@ml_IdLote", adInteger, adParamInput, 0, ml_IdLote): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerDatosLoteNroHojaLibre = 0
   If oRecordset.RecordCount > 0 Then
      ObtenerDatosLoteNroHojaLibre = CInt(oRecordset!NroHojaHis)
   End If
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ConsultarRegistroFiltroLotes(IdEstablecimiento As Long, Anio As Integer, IdMes As Integer, Lote As String, Id As Boolean) As Recordset
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If Anio <> 0 Then
        sSql = sSql & " AND LT.Anio = " & Anio
    End If
    'mes
    If IdMes <> 0 Then
        sSql = sSql & " AND LT.Mes = " & IdMes
    End If
    'codigo de lote
    If Lote <> "" Then
        sSql = sSql & " AND LT.Lote = '" & Lote & "'"
    End If
    sSql = sSql & " and he.idestado in (0,1)"

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesSegunFiltro"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstablecimiento", adVarChar, adParamInput, 10, Trim(Str(IdEstablecimiento))): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ConsultarRegistroFiltroLotes = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function His_LotesConsultarFiltro(IdEstablecimiento As Long, IdEstado As Integer) As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "His_LotesConsultarFiltro"
            Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IdEstado): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set His_LotesConsultarFiltro = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function


Function ObtenerDatosNumeroHojasUtilizadas(IdHisLote As Long) As Integer
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
        .CommandText = "HIS_CabeceraNumeroHojasUtilizadas"
        Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, IdHisLote): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerDatosNumeroHojasUtilizadas = 0
   If oRecordset.RecordCount > 0 Then
      ObtenerDatosNumeroHojasUtilizadas = CInt(oRecordset!CantidadUtilizada)
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ModificarRegistroLoteHISAdicionarHoja(ml_IdLote As Long, NumHoja As Integer) As Boolean
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
        .CommandText = "HIS_LotesActualizarXidLoteYnroHoja"
        Set oParameter = .CreateParameter("@NumHoja", adInteger, adParamInput, 0, NumHoja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_IdLote", adInteger, adParamInput, 0, ml_IdLote): .Parameters.Append oParameter
         .Execute
        
   End With
   ModificarRegistroLoteHISAdicionarHoja = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function DisminuirNroHojaHIS(ml_IdLote As Long) As Boolean
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
        .CommandText = "HIS_LotesDisminuirNroHojaHIS"
        Set oParameter = .CreateParameter("@ml_IdLote", adInteger, adParamInput, 0, ml_IdLote): .Parameters.Append oParameter
         .Execute
   End With
   DisminuirNroHojaHIS = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function His_ConsultarHojasRegistradas(IdEstablecimiento As Long, IdLote As Long) As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "His_ConsultarHojasRegistradas"
            Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IdLote): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set His_ConsultarHojasRegistradas = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function

Function His_ConsultarTotalRegistrosLote(IdEstablecimiento As Long, IdLote As Long) As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "His_ConsultarTotalRegistrosLote"
            Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IdLote): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set His_ConsultarTotalRegistrosLote = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function

Function HIS_ConsultarRegMuestraLotes(IdLote As Long) As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "HIS_ConsultarRegMuestraLotes"
            Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IdLote): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set HIS_ConsultarRegMuestraLotes = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function



Function HisActualizarNroRegistroHisDetalle(ml_IdHisDetalle As Long, ml_NumReg As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "His_ActualizarNroRegistroHisDetalle"
        Set oParameter = .CreateParameter("@HisDetalle", adInteger, adParamInput, 0, ml_IdHisDetalle): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NumReg", adInteger, adParamInput, 0, ml_NumReg): .Parameters.Append oParameter
         .Execute
   End With
   HisActualizarNroRegistroHisDetalle = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function HIS_ConsultarRegistrosTotalesLotes(IdLote As Long) As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "HIS_ConsultarRegistrosTotalesLotes"
            Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IdLote): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set HIS_ConsultarRegistrosTotalesLotes = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function

Function HIS_ConsultarEstadosLote() As Recordset
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandTimeout = 150
            .CommandText = "HIS_ConsultarEstadosLote"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set HIS_ConsultarEstadosLote = oRecordset
       Set oCommand = Nothing
       Exit Function
ManejadorDeError:
        MsgBox Err.Description
End Function

