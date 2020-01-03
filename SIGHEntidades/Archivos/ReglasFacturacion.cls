VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Facturación
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim ms_MensajeError As String
Dim ms_Texto As String
Property Let Texto(sValue As String)
   ms_Texto = sValue
End Property
Property Get Texto() As String
   Texto = ms_Texto
End Property

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'-------------------------------------------------------------------------------------
'                                       Tipos de Financiandor
'-------------------------------------------------------------------------------------
'mgaray201503
Public Function TipoFinanciadorSeleccionarPorId(lcIdTipoFinanciador As Long) As DoTipoFinanciador
    Dim oConexion As New ADODB.Connection
    Dim oTipoFinanciador As New TipoFinanciador
    Dim oDOTipoFinanciador As New DoTipoFinanciador
    
    Set TipoFinanciadorSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTipoFinanciador.Conexion = oConexion
    oDOTipoFinanciador.idTipoFinanciador = lcIdTipoFinanciador
    If oTipoFinanciador.SeleccionarPorId(oDOTipoFinanciador) Then
        Set TipoFinanciadorSeleccionarPorId = oDOTipoFinanciador
    End If
    oConexion.Close

    ms_MensajeError = oTipoFinanciador.MensajeError
    Set oConexion = Nothing
    Set oTipoFinanciador = Nothing
End Function

Function TipoFinanciadorSeleccionarTodos() As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New TipoFinanciador
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TipoFinanciadorSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       Fuentes de Financiamiento
'-------------------------------------------------------------------------------------

Function FuentesFinanciamientoAgregar(oDOFuenteFinanciamiento As DOFuenteFinanciamiento, oRsTarifa As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDplan As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oFuentesFinanciamiento As New FuentesFinanciamiento
Dim oRsFuenteFinanciamientoTarifas    As New Recordset
    FuentesFinanciamientoAgregar = False

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oFuentesFinanciamiento.Conexion = oConexion
    If oFuentesFinanciamiento.Insertar(oDOFuenteFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFuenteFinanciamiento.IdUsuarioAuditoria, "A", oDOFuenteFinanciamiento.idFuenteFinanciamiento, "FuentesFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDplan)              'ListBarItems.idListItem
        'Graba Tarifas
        oRsTarifa.MoveFirst
        Do While Not oRsTarifa.EOF

            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FuentesFinanciamientoTarifasAgregar"
                Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oDOFuenteFinanciamiento.idFuenteFinanciamiento): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oRsTarifa.Fields!IdTipoFinanciamiento): .Parameters.Append oParameter
                .Execute
            End With
            Set oParameter = Nothing
            Set oCommand = Nothing
           oRsTarifa.MoveNext
        Loop
        ms_MensajeError = ""
        FuentesFinanciamientoAgregar = True
    Else
        ms_MensajeError = oFuentesFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    'Set oCommand = Nothing
    Set oFuentesFinanciamiento = Nothing
    Set oRsTarifa = Nothing
    
End Function

Function FuentesFinanciamientoModificar(oDOFuenteFinanciamiento As DOFuenteFinanciamiento, oRsTarifa As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDplan As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oFuentesFinanciamiento As New FuentesFinanciamiento
Dim oRsFuenteFinanciamientoTarifas    As New Recordset
    FuentesFinanciamientoModificar = False
    
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexion
  
    Set oFuentesFinanciamiento.Conexion = oConexion
    If oFuentesFinanciamiento.modificar(oDOFuenteFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFuenteFinanciamiento.IdUsuarioAuditoria, "M", oDOFuenteFinanciamiento.idFuenteFinanciamiento, "FuentesFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDplan)              'ListBarItems.idListItem
        FuentesFinanciamientoModificar = True
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFuenteFinanciamiento.IdUsuarioAuditoria, "A", oDOFuenteFinanciamiento.idFuenteFinanciamiento, "FuentesFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDplan)              'ListBarItems.idListItem
        'Graba Tarifas

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FuentesFinanciamientoTarifasEliminaPorIdFuenteFinanciamiento"
            Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oDOFuenteFinanciamiento.idFuenteFinanciamiento): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        
        oRsTarifa.MoveFirst
        Do While Not oRsTarifa.EOF
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FuentesFinanciamientoTarifasAgregarPorIdFuenteFinanciamiento"
                Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oDOFuenteFinanciamiento.idFuenteFinanciamiento): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oRsTarifa.Fields!IdTipoFinanciamiento): .Parameters.Append oParameter
                .Execute
            End With
        
           oRsTarifa.MoveNext
        Loop
        ms_MensajeError = ""
    Else
        ms_MensajeError = oFuentesFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFuentesFinanciamiento = Nothing
    Set oRsFuenteFinanciamientoTarifas = Nothing
    
End Function

Function FuentesFinanciamientoEliminar(oDOFuenteFinanciamiento As DOFuenteFinanciamiento, oRsTarifa As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDplan As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFuentesFinanciamiento As New FuentesFinanciamiento
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsFuenteFinanciamientoTarifas    As New Recordset
    On Error GoTo ErrElimFF
    FuentesFinanciamientoEliminar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oFuentesFinanciamiento.Conexion = oConexion
    'Elimina Tarifas
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FuentesFinanciamientoTarifasEliminaPorIdFuenteFinanciamiento"
            Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oDOFuenteFinanciamiento.idFuenteFinanciamiento): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
    
    If oFuentesFinanciamiento.Eliminar(oDOFuenteFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFuenteFinanciamiento.IdUsuarioAuditoria, "E", oDOFuenteFinanciamiento.idFuenteFinanciamiento, "FuentesFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDplan)              'ListBarItems.idListItem
        FuentesFinanciamientoEliminar = True
    Else
        ms_MensajeError = oFuentesFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFuentesFinanciamiento = Nothing
    Set oRsFuenteFinanciamientoTarifas = Nothing
    Exit Function
ErrElimFF:
    ms_MensajeError = Err.Description
End Function

Function FuentesFinanciamientoSeleccionarPorId(IdFuentesFinanciamiento As Integer) As DOFuenteFinanciamiento
Dim oConexion As New ADODB.Connection
Dim oFuentesFinanciamiento As New FuentesFinanciamiento
Dim oDOFuenteFinanciamiento As New DOFuenteFinanciamiento
    
    Set FuentesFinanciamientoSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFuentesFinanciamiento.Conexion = oConexion
    oDOFuenteFinanciamiento.idFuenteFinanciamiento = IdFuentesFinanciamiento
    If oFuentesFinanciamiento.SeleccionarPorId(oDOFuenteFinanciamiento) Then
        Set FuentesFinanciamientoSeleccionarPorId = oDOFuenteFinanciamiento
    End If
    oConexion.Close

    ms_MensajeError = oFuentesFinanciamiento.MensajeError
    Set oConexion = Nothing
    Set oFuentesFinanciamiento = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FuentesFinanciamientoSeleccionarPorTipo(lIdFuenteFinanciamiento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FuentesFinanciamiento
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FuentesFinanciamientoSeleccionarPorTipo = oTabla.SeleccionarPorTipoFinanciamiento(lIdFuenteFinanciamiento)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FuentesFinanciamientoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FuentesFinanciamiento
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FuentesFinanciamientoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'mgaray201503
Function FuentesFinanciamientoSeleccionarPorCodigo(sCodigo As String) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New FuentesFinanciamiento
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FuentesFinanciamientoSeleccionarPorCodigo = oTabla.SeleccionarPorCodigo(sCodigo)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


'-------------------------------------------------------------------------------------
'                                       Tipos de Financiamiento
'-------------------------------------------------------------------------------------

Function TiposFinanciamientoAgregar(oDoTiposFinanciamiento As DOTiposFinanciamiento, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDfinanciamiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposFinanciamiento As New TiposFinanciamiento
    
    TiposFinanciamientoAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTiposFinanciamiento.Conexion = oConexion
    If oTiposFinanciamiento.Insertar(oDoTiposFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoTiposFinanciamiento.IdUsuarioAuditoria, "A", oDoTiposFinanciamiento.IdTipoFinanciamiento, "TiposFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDfinanciamiento)           'ListBarItems.idListItem
        TiposFinanciamientoAgregar = True
    Else
        ms_MensajeError = oTiposFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposFinanciamiento = Nothing
    
End Function

Function TiposFinanciamientoModificar(oDoTiposFinanciamiento As DOTiposFinanciamiento, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDfinanciamiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposFinanciamiento As New TiposFinanciamiento
    
    TiposFinanciamientoModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTiposFinanciamiento.Conexion = oConexion
    If oTiposFinanciamiento.modificar(oDoTiposFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoTiposFinanciamiento.IdUsuarioAuditoria, "M", oDoTiposFinanciamiento.IdTipoFinanciamiento, "TiposFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDfinanciamiento)           'ListBarItems.idListItem
        TiposFinanciamientoModificar = True
    Else
        ms_MensajeError = oTiposFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposFinanciamiento = Nothing
    
End Function

Function TiposFinanciamientoEliminar(oDoTiposFinanciamiento As DOTiposFinanciamiento, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcDfinanciamiento As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTiposFinanciamiento As New TiposFinanciamiento
    
    TiposFinanciamientoEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTiposFinanciamiento.Conexion = oConexion
    If oTiposFinanciamiento.Eliminar(oDoTiposFinanciamiento) Then
        '
        Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoTiposFinanciamiento.IdUsuarioAuditoria, "E", oDoTiposFinanciamiento.IdTipoFinanciamiento, "TiposFinanciamiento", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcDfinanciamiento)           'ListBarItems.idListItem
        TiposFinanciamientoEliminar = True
    Else
        ms_MensajeError = oTiposFinanciamiento.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oTiposFinanciamiento = Nothing
    
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function TiposFinanciamientoSeleccionarPorId(IdTiposFinanciamiento As Integer) As DOTiposFinanciamiento
Dim oConexion As New ADODB.Connection
Dim oTiposFinanciamiento As New TiposFinanciamiento
Dim oDoTiposFinanciamiento As New DOTiposFinanciamiento
    
    Set TiposFinanciamientoSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTiposFinanciamiento.Conexion = oConexion
    oDoTiposFinanciamiento.IdTipoFinanciamiento = IdTiposFinanciamiento
    If oTiposFinanciamiento.SeleccionarPorId(oDoTiposFinanciamiento) Then
        Set TiposFinanciamientoSeleccionarPorId = oDoTiposFinanciamiento
    End If
    oConexion.Close

    ms_MensajeError = oTiposFinanciamiento.MensajeError
    Set oConexion = Nothing
    Set oTiposFinanciamiento = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function TiposFinanciamientoSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposFinanciamiento
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposFinanciamientoSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposFinanciamientoGeneraReciboPago(lnIdtipoFinanciamiento As Long, oConexion As Connection) As Boolean
Dim oTabla As New TiposFinanciamiento
Dim oRsTmpReglaFacturacion As New Recordset
    Set oTabla.Conexion = oConexion
    Set oRsTmpReglaFacturacion = oTabla.SeleccionarPorIdentificador(lnIdtipoFinanciamiento)
    If oRsTmpReglaFacturacion.Fields!GeneraPago = 1 Then
       TiposFinanciamientoGeneraReciboPago = True
    Else
       TiposFinanciamientoGeneraReciboPago = False
    End If
    oRsTmpReglaFacturacion.Close
    Set oRsTmpReglaFacturacion = Nothing
    Set oTabla = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function TiposFinanciamientoSeleccionarParaCaja() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposFinanciamiento
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set TiposFinanciamientoSeleccionarParaCaja = oTabla.SeleccionarParaCaja
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function TiposFinanciamientoDevuelveIdTipoConceptoFarmacia(lnIdtipoFinanciamiento As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposFinanciamiento
Dim oRsTmpReglaFacturacion As New Recordset
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set oRsTmpReglaFacturacion = oTabla.SeleccionarPorIdentificador(lnIdtipoFinanciamiento)
    TiposFinanciamientoDevuelveIdTipoConceptoFarmacia = oRsTmpReglaFacturacion.Fields!idTipoConceptoFarmacia
    oRsTmpReglaFacturacion.Close
    oConexion.Close
    Set oRsTmpReglaFacturacion = Nothing
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia(lnIdFuenteFinanciamiento As Long, oConexion As Connection) As Long
Dim oTabla As New FuentesFinanciamiento
Dim oRsTmpReglaFacturacion As New Recordset
    Set oTabla.Conexion = oConexion
    Set oRsTmpReglaFacturacion = oTabla.SeleccionarPorIdentificador(lnIdFuenteFinanciamiento)
    If lnIdFuenteFinanciamiento > 0 Then
       FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia = oRsTmpReglaFacturacion.Fields!idTipoConceptoFarmacia
    Else
       FuentesFinanciamientoDevuelveIdTipoConceptoFarmacia = 0
    End If
    oRsTmpReglaFacturacion.Close
    Set oRsTmpReglaFacturacion = Nothing
    Set oTabla = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FuentesFinanciamientoDevuelveIdTipoFinanciamiento(lnIdFuenteFinanciamiento As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New FuentesFinanciamiento
Dim oRsTmpReglaFacturacion As New Recordset
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set oRsTmpReglaFacturacion = oTabla.SeleccionarPorIdentificador(lnIdFuenteFinanciamiento)
    FuentesFinanciamientoDevuelveIdTipoFinanciamiento = oRsTmpReglaFacturacion.Fields!IdTipoFinanciamiento
    oRsTmpReglaFacturacion.Close
    oConexion.Close
    Set oRsTmpReglaFacturacion = Nothing
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
Function FuentesFinanciamientoDevuelveNombrePlan(lnIdFuenteFinanciamiento As Long) As String
Dim oConexion As New ADODB.Connection
Dim oTabla As New FuentesFinanciamiento
Dim oRsTmp As New Recordset
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oTabla.Conexion = oConexion
    Set oRsTmp = oTabla.SeleccionarPorIdentificador(lnIdFuenteFinanciamiento)
    FuentesFinanciamientoDevuelveNombrePlan = oRsTmp.Fields!Descripcion
    oRsTmp.Close
    oConexion.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
'-------------------------------------------------------------------------------------
'                                       Cuentas de Atencion
'-------------------------------------------------------------------------------------

Function CuentasAtencionAgregar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    CuentasAtencionAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentasAtencion.Conexion = oConexion
    If oCuentasAtencion.Insertar(oDOCuentaAtencion) Then
        CuentasAtencionAgregar = True
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    
End Function

Function CuentasAtencionModificar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    CuentasAtencionModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentasAtencion.Conexion = oConexion
    If oCuentasAtencion.modificar(oDOCuentaAtencion) Then
        CuentasAtencionModificar = True
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    
End Function

Function CuentasAtencionEliminar(oDOCuentaAtencion As DOCuentaAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    CuentasAtencionEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentasAtencion.Conexion = oConexion
    If oCuentasAtencion.Eliminar(oDOCuentaAtencion) Then
        CuentasAtencionEliminar = True
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    
End Function

Function CuentasAtencionSeleccionarPorId(IdCuentasAtencion As Long, oConexion As Connection) As SIGHcomun.DOCuentaAtencion
    Dim oCuentasAtencion As New CuentasAtencion
    Dim oDOCuentaAtencion As New DOCuentaAtencion
    Set CuentasAtencionSeleccionarPorId = Nothing
    Set oCuentasAtencion.Conexion = oConexion
    oDOCuentaAtencion.IdCuentaAtencion = IdCuentasAtencion
    If oCuentasAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
        Set CuentasAtencionSeleccionarPorId = oDOCuentaAtencion
    End If
    ms_MensajeError = oCuentasAtencion.MensajeError
    Set oCuentasAtencion = Nothing
    Set oDOCuentaAtencion = Nothing
End Function
Function CuentasAtencionSeleccionarIdPorIdAtencion(IdCuentaAtencion As Long) As Long
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    CuentasAtencionSeleccionarIdPorIdAtencion = 0
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentasAtencion.Conexion = oConexion
    
    CuentasAtencionSeleccionarIdPorIdAtencion = oCuentasAtencion.SeleccionarIdPorIdAtencion(IdCuentaAtencion)
        
    ms_MensajeError = oCuentasAtencion.MensajeError
    
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CuentasAtencionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CuentasAtencionSeleccionarConsultaExterna() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionSeleccionarConsultaExterna = oTabla.SeleccionarConsultaExterna
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CuentasAtencionSeleccionarTopicoEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionSeleccionarTopicoEmergencia = oTabla.SeleccionarTopicoEmergencia
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CuentasAtencionSeleccionarObservacionEmergencia() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionSeleccionarObservacionEmergencia = oTabla.SeleccionarObservacionEmergencia
    ms_MensajeError = oTabla.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CuentasAtencionSeleccionarHospitalizacion() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set CuentasAtencionSeleccionarHospitalizacion = oTabla.SeleccionarHospitalizacion
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CuentasAtencionAbrir(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
    
    CuentasAtencionAbrir = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCuentasAtencion.Conexion = oConexion
    If oCuentasAtencion.Abrir(oDOCuentaAtencion) Then
        CuentasAtencionAbrir = True
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "M", oDOCuentaAtencion.IdCuentaAtencion, "Abrió Cta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)           'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    
End Function

Function CuentasAtencionCerrar(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCuentasAtencion As New CuentasAtencion
Dim oRsTmp As New ADODB.Recordset
Dim sSql As String
    On Error GoTo ErrCierre
    CuentasAtencionCerrar = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oCuentasAtencion.Conexion = oConexion
    oConexion.BeginTrans
    If oCuentasAtencion.Cerrar(oDOCuentaAtencion) Then
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasActualizarXIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDOCuentaAtencion.idPaciente): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "M", oDOCuentaAtencion.IdCuentaAtencion, "Cerró Cta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)           'ListBarItems.idListItem
    oConexion.CommitTrans
    oConexion.Close
    CuentasAtencionCerrar = True
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
    Exit Function
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
End Function

Function PendientePagoSeguro(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCuentasAtencion As New CuentasAtencion
Dim oRsTmp As New ADODB.Recordset
Dim sSql As String
    On Error GoTo ErrCierre
    PendientePagoSeguro = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oCuentasAtencion.Conexion = oConexion
    oConexion.BeginTrans
    If oCuentasAtencion.PendientePagoSeguro(oDOCuentaAtencion) Then
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    
    With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "CamasActualizarXIdPaciente"
    Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDOCuentaAtencion.idPaciente): .Parameters.Append oParameter
    .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "A", oDOCuentaAtencion.IdCuentaAtencion, "Pasó a PtePagSeg", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)           'ListBarItems.idListItem
    oConexion.CommitTrans
    oConexion.Close
    PendientePagoSeguro = True
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
    Exit Function
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionOrdenesServicioFiltrar(ByVal oDOFactOrdenServicio As DOFactOrdenServicio, oDoPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesServicio
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set AtencionOrdenesServicioFiltrar = oTabla.Filtrar(oDOFactOrdenServicio, oDoPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionOrdenesServicioFiltrarDEBB(ByVal oDOFactOrdenServicio As DOFactOrdenServicio, oDoPaciente As doPaciente, lcFecha As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesServicio
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set AtencionOrdenesServicioFiltrarDEBB = oTabla.FiltrarDEBB(oDOFactOrdenServicio, oDoPaciente, lcFecha)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionOrdenesBienInsumoFiltrar(ByVal oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, oDoPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesBienesInsumo
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set AtencionOrdenesBienInsumoFiltrar = oTabla.Filtrar(oDOFactOrdenBienInsumo, oDoPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionOrdenesBienInsumoFiltrarDEBB(ByVal oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, oDoPaciente As doPaciente, ldFechaInicial As Date, ldFechaFinal As Date, lnFarmacia As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesBienesInsumo
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set AtencionOrdenesBienInsumoFiltrarDEBB = oTabla.FiltrarDEBB(oDOFactOrdenBienInsumo, oDoPaciente, ldFechaInicial, ldFechaFinal, lnFarmacia)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionSeleccionarTiposConsultaPorEspecialidad(lIdEspecialidad As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    Set oTabla.Conexion = oConexion
    Set FacturacionSeleccionarTiposConsultaPorEspecialidad = oTabla.SeleccionarTiposConsultaPorEspecialidad(lIdEspecialidad)
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       Facturacion Apoyo Diagnostico
'-------------------------------------------------------------------------------------
'Function AtencionApoyoDxAgregar(oDOAtencionApoyoDiagnostico As DOAtencionApoyoDiagnostico, oProcedimientoDetalle As Collection) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDx As New AtencionApoyoDiagnostico
'Dim oDOCuentaAtencion As New DOCuentaAtencion
'Dim oCuentaAtencion As New CuentasAtencion
'Dim bProcesoOK As Boolean
'Dim oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle
'
'    AtencionApoyoDxAgregar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'
'    bProcesoOK = True
'    oConexion.BeginTrans
'
'    Set oAtencionApoyoDx.Conexion = oConexion
'    Set oCuentaAtencion.Conexion = oConexion
'
'    'Agrega la cabecera
'    If Not oAtencionApoyoDx.Insertar(oDOAtencionApoyoDiagnostico) Then
'        bProcesoOK = False: GoTo Terminar
'    End If
'
'    oDOCuentaAtencion.IdCuentaAtencion = oDOAtencionApoyoDiagnostico.IdCuentaAtencion
'    If Not oCuentaAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
'        bProcesoOK = False: GoTo Terminar
'    End If
'
'    'Agrega el detalle
'    For Each oDOAtencionApoyoDiagDetalle In oProcedimientoDetalle
'        'Servicios a Facturar
'        If Not ApoyoDxAgregarServicioYProcedimiento(oConexion, oDOAtencionApoyoDiagDetalle.IdProcedimiento, oDOCuentaAtencion, oDOAtencionApoyoDiagnostico, oDOAtencionApoyoDiagDetalle) Then
'            bProcesoOK = False: GoTo Terminar
'        End If
'    Next
'
'Terminar:
'    If bProcesoOK Then
'        AtencionApoyoDxAgregar = True
'        oConexion.CommitTrans
'    Else
'        AtencionApoyoDxAgregar = False
'        ms_MensajeError = oAtencionApoyoDx.MensajeError
'        oConexion.RollbackTrans
'    End If
'
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oAtencionApoyoDx = Nothing
'    Set oCuentaAtencion = Nothing
'    Set oDOCuentaAtencion = Nothing
'End Function
''WCG-05 Ini
'Function ApoyoDxAgregarServicioYProcedimiento(oConexion As Connection, lIdProcedimiento As Long, oDOCuentaAtencion As DOCuentaAtencion, oDOAtencionApoyoDiagnostico As DOAtencionApoyoDiagnostico, oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle) As Boolean
'Dim oDOFacturacionServicio As New DOFacturacionServicios
'Dim oDOCatalogoServicio As New DOCatalogoServicio
'Dim oProcedimiento As New Procedimientos
'Dim oDOProcedimiento As New doProcedimiento
'Dim oCatalogoServicios As New CatalogoServicios
'Dim oFacturacionServicios As New FacturacionServicios
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'
'    ApoyoDxAgregarServicioYProcedimiento = False
'
'    Set oProcedimiento.Conexion = oConexion
'    Set oFacturacionServicios.Conexion = oConexion
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    Set oCatalogoServicios.Conexion = oConexion
'
'    oDOProcedimiento.IdProcedimiento = lIdProcedimiento
'    If oProcedimiento.SeleccionarPorId(oDOProcedimiento) Then
'        If oDOProcedimiento.idProducto <> 0 Then
'            '<<<<<<<<<<<
'            With oDOFacturacionServicio
'                .idAtencion = 0
'                .idFacturacionServicio = 0
'                .idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente
'                .IdTipoFinanciamiento = sghPacienteNormal
'                .cantidad = 1
'                .idProducto = oDOProcedimiento.idProducto
'                .IdUsuarioAuditoria = oDOAtencionApoyoDiagnostico.IdUsuarioAuditoria
'                .IdEstadoFacturacion = sghEstadoFacturacion.sghPendientePago
'                .FechaAutorizaPendiente = 0
'                .FechaAutorizaSeguro = 0
'                .IdCentroCosto = 0
'                .IdUsuarioAutorizaPendiente = 0
'                .IdUsuarioAutorizaSeguro = 0
'                .PrecioUnitario = 0
'                .totalPorPagar = 0
'            End With
'            '>>>>>>>>>>>
'
'            oDOCatalogoServicio.idProducto = oDOProcedimiento.idProducto
'            If oCatalogoServicios.SeleccionarPorId(oDOCatalogoServicio) Then
'                oDOFacturacionServicio.IdCentroCosto = oDOCatalogoServicio.IdCentroCosto
'                'oDOFacturacionServicio.IdClasificacionServicio = oDOCatalogoServicio.IdClasificacionServicio
'            End If
'
'            'Obtener Precio de acuerdo al plan principal
'            oDOFacturacionServicio.PrecioUnitario = 0   'ObtenerPrecioUnitario(oDOFacturacionServicio.IdProducto, sghPacienteNormal)
'            If (oDOFacturacionServicio.PrecioUnitario >= 0) Then
'                oDOFacturacionServicio.totalPorPagar = oDOFacturacionServicio.cantidad * oDOFacturacionServicio.PrecioUnitario
'                If oFacturacionServicios.Insertar(oDOFacturacionServicio) Then
'                    'Agrega a la tabla procedimiento
'                    oDOAtencionApoyoDiagDetalle.IdAtencionApoyoDx = oDOAtencionApoyoDiagnostico.IdAtencionApoyoDx
'                    oDOAtencionApoyoDiagDetalle.idFacturacionServicio = oDOFacturacionServicio.idFacturacionServicio
'                    If Not oAtencionApoyoDxDetalle.Insertar(oDOAtencionApoyoDiagDetalle) Then
'                        Exit Function
'                    End If
'                Else
'                    Exit Function
'                End If
'
'            End If
'        End If
'    End If
'
'    ApoyoDxAgregarServicioYProcedimiento = True
'
'End Function
''WCG-05 Fin
'
'Function AtencionApoyoDxModificar(oDOAtencionApoyoDiagnostico As DOAtencionApoyoDiagnostico, oProcedimientoDetalle As Collection, rsProcedimientosEliminados As Recordset) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDx As New AtencionApoyoDiagnostico
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'Dim bProcesoOK As Boolean
'Dim oCuentaAtencion As New CuentasAtencion
'Dim oProcedimiento As New Procedimientos
'Dim oFacturacionServicios As New FacturacionServicios
'Dim oDOFacturacionServicio As New DOFacturacionServicios
'Dim oDOProcedimiento As New doProcedimiento
'Dim oDOCuentaAtencion As New DOCuentaAtencion
'Dim oDOCatalogoServicio As New DOCatalogoServicio
'
'    AtencionApoyoDxModificar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    bProcesoOK = True
'    oConexion.BeginTrans
'
'    Set oAtencionApoyoDx.Conexion = oConexion
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    Set oCuentaAtencion.Conexion = oConexion
'    Set oFacturacionServicios.Conexion = oConexion
'
'    'Modifica la cabecera
'    If Not oAtencionApoyoDx.Modificar(oDOAtencionApoyoDiagnostico) Then
'        bProcesoOK = False: GoTo Terminar
'    End If
'
'    oDOCuentaAtencion.IdCuentaAtencion = oDOAtencionApoyoDiagnostico.IdCuentaAtencion
'    If Not oCuentaAtencion.SeleccionarPorId(oDOCuentaAtencion) Then
'        bProcesoOK = False: GoTo Terminar
'    End If
'
'    'Agrega el detalle
'    Dim oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle
'    For Each oDOAtencionApoyoDiagDetalle In oProcedimientoDetalle
'        If oDOAtencionApoyoDiagDetalle.EstadoRegistro = "A" Then
'            '****************************************
'            If Not ApoyoDxAgregarServicioYProcedimiento(oConexion, oDOAtencionApoyoDiagDetalle.IdProcedimiento, oDOCuentaAtencion, oDOAtencionApoyoDiagnostico, oDOAtencionApoyoDiagDetalle) Then
'                bProcesoOK = False: GoTo Terminar
'            End If
'            '****************************************
'        End If
'        If oDOAtencionApoyoDiagDetalle.EstadoRegistro = "M" Then
'        End If
'    Next
'
'
'    Dim oDOAtencionApoyoDiagDetalleEliminado As New DOAtencionApoyoDiagDetalle
'
'    If Not (rsProcedimientosEliminados.EOF And rsProcedimientosEliminados.BOF) Then
'        rsProcedimientosEliminados.MoveFirst
'        Do While Not rsProcedimientosEliminados.EOF
'
'            oDOAtencionApoyoDiagDetalleEliminado.IdAtencionApoyoDetalle = rsProcedimientosEliminados!IdAtencionProcDetalle
'            If Not oAtencionApoyoDxDetalle.Eliminar(oDOAtencionApoyoDiagDetalleEliminado) Then
'                bProcesoOK = False: GoTo Terminar
'            End If
'
'            oDOFacturacionServicio.idFacturacionServicio = rsProcedimientosEliminados!idFacturacionServicio
'            If Not oFacturacionServicios.Eliminar(oDOFacturacionServicio) Then
'                bProcesoOK = False: GoTo Terminar
'            End If
'
'            rsProcedimientosEliminados.MoveNext
'        Loop
'    End If
'
'Terminar:
'    If bProcesoOK Then
'        AtencionApoyoDxModificar = True
'        oConexion.CommitTrans
'    Else
'        AtencionApoyoDxModificar = False
'        ms_MensajeError = oAtencionApoyoDx.MensajeError + Chr(13) + oAtencionApoyoDxDetalle.MensajeError
'        oConexion.RollbackTrans
'    End If
'
'    oConexion.Close
'End Function
'
'
'Function AtencionApoyoDxSeleccionarPorId(IdAtencionApoyoDx As Long) As DOAtencionApoyoDiagnostico
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDx As New AtencionApoyoDiagnostico
'Dim oDOAtencionApoyoDiagnostico As New DOAtencionApoyoDiagnostico
'
'    Set AtencionApoyoDxSeleccionarPorId = Nothing
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    Set oAtencionApoyoDx.Conexion = oConexion
'    oDOAtencionApoyoDiagnostico.IdAtencionApoyoDx = IdAtencionApoyoDx
'    If oAtencionApoyoDx.SeleccionarPorId(oDOAtencionApoyoDiagnostico) Then
'        Set AtencionApoyoDxSeleccionarPorId = oDOAtencionApoyoDiagnostico
'    End If
'    oConexion.Close
'
'    ms_MensajeError = oAtencionApoyoDx.MensajeError
'
'End Function
'Function AtencionApoyoDxFiltrar(ByVal oDOAtencionApoyoDiagnostico As DOAtencionApoyoDiagnostico, oDOPAciente As doPaciente, lDepartamento As Long) As ADODB.Recordset
'Dim oConexion As New ADODB.Connection
'Dim oTabla As New AtencionApoyoDiagnostico
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    Set oTabla.Conexion = oConexion
'    Set AtencionApoyoDxFiltrar = oTabla.Filtrar(oDOAtencionApoyoDiagnostico, oDOPAciente, lDepartamento)
'    oConexion.Close
'
'    ms_MensajeError = oTabla.MensajeError
'    Set oConexion = Nothing
'    Set oTabla = Nothing
'
'End Function
'
''-------------------------------------------------------------------------------------
''                                       Facturacion Procedimientos Detalle
''-------------------------------------------------------------------------------------
'
'Function AtencionApoyoDxDetalleAgregar(oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'
'    AtencionApoyoDxDetalleAgregar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    If oAtencionApoyoDxDetalle.Insertar(oDOAtencionApoyoDiagDetalle) Then
'        AtencionApoyoDxDetalleAgregar = True
'    Else
'        ms_MensajeError = oAtencionApoyoDxDetalle.MensajeError
'    End If
'    oConexion.Close
'
'End Function
'
'Function AtencionApoyoDxDetalleModificar(oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'
'    AtencionApoyoDxDetalleModificar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    If oAtencionApoyoDxDetalle.Modificar(oDOAtencionApoyoDiagDetalle) Then
'        AtencionApoyoDxDetalleModificar = True
'    Else
'        ms_MensajeError = oAtencionApoyoDxDetalle.MensajeError
'    End If
'    oConexion.Close
'
'End Function
'
'Function AtencionApoyoDxDetalleEliminar(oDOAtencionApoyoDiagDetalle As DOAtencionApoyoDiagDetalle) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'
'    AtencionApoyoDxDetalleEliminar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    If oAtencionApoyoDxDetalle.Eliminar(oDOAtencionApoyoDiagDetalle) Then
'        AtencionApoyoDxDetalleEliminar = True
'    Else
'        ms_MensajeError = oAtencionApoyoDxDetalle.MensajeError
'    End If
'    oConexion.Close
'
'End Function
'
'Function AtencionApoyoDxDetalleSeleccionarPorId(IdAtencionApoyoDxDetalle As Integer) As DOAtencionApoyoDiagDetalle
'Dim oConexion As New ADODB.Connection
'Dim oAtencionApoyoDxDetalle As New AtencionApoyoDiagDetalle
'Dim oDOAtencionApoyoDiagDetalle As New DOAtencionApoyoDiagDetalle
'
'    Set AtencionApoyoDxDetalleSeleccionarPorId = Nothing
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    Set oAtencionApoyoDxDetalle.Conexion = oConexion
'    oDOAtencionApoyoDiagDetalle.IdAtencionApoyoDetalle = IdAtencionApoyoDxDetalle
'    If oAtencionApoyoDxDetalle.SeleccionarPorId(oDOAtencionApoyoDiagDetalle) Then
'        Set AtencionApoyoDxDetalleSeleccionarPorId = oDOAtencionApoyoDiagDetalle
'    End If
'    oConexion.Close
'
'    ms_MensajeError = oAtencionApoyoDxDetalle.MensajeError
'
'End Function
'
'
'Function AtencionApoyoDxDetalleSeleccionarPorIdAtencionApoyoDx(lIdAtencionApoyoDx As Long) As ADODB.Recordset
'Dim oConexion As New ADODB.Connection
'Dim oTabla As New AtencionApoyoDiagDetalle
'
'    oConexion.Open sighentidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'
'    Set oTabla.Conexion = oConexion
'    Set AtencionApoyoDxDetalleSeleccionarPorIdAtencionApoyoDx = oTabla.SeleccionarPorIdAtencionApoyoDx(lIdAtencionApoyoDx)
'    oConexion.Close
'
'    ms_MensajeError = oTabla.MensajeError
'    Set oConexion = Nothing
'    Set oTabla = Nothing
'
'End Function




'-------------------------------------------------------------------------------------
'                   ATENCION
'-------------------------------------------------------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EstadosAtencionObtenerTodos() As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosAtencion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EstadosAtencionObtenerTodos = oTabla.SeleccionarTodos()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'-------------------------------------------------------------------------------------
'                   PREFACTURACION EXONERACIONES
'-------------------------------------------------------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EstadosFacturacionObtenerTodos() As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosFacturacion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EstadosFacturacionObtenerTodos = oTabla.SeleccionarTodos()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EstadosFacturacionObtenerTodosExceptoPagado() As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosFacturacion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EstadosFacturacionObtenerTodosExceptoPagado = oTabla.SeleccionarTodosExceptoPagado()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EstadosFacturacionObtenerDescripcionPorId(IdEstadoFacturacion As Long) As String
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosFacturacion
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    EstadosFacturacionObtenerDescripcionPorId = oTabla.ObtenerDescripcionPorId(IdEstadoFacturacion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaExoneracionServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaExoneracionServicio = oTabla.SeleccionarParaExoneracionServicio(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaExoneracionBienInsumo(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaExoneracionBienInsumo = oTabla.SeleccionarParaExoneracionBienesInsumos(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                   PREFACTURACION PENDIENTES DE PAGO
'-------------------------------------------------------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServiciosObtenerParaPendientePago(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionServiciosObtenerParaPendientePago = oTabla.SeleccionarParaPendientePago(lIdCuentaAtencion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesInsumosObtenerParaPendientePago(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesInsumosObtenerParaPendientePago = oTabla.SeleccionarParaPendientePago(lIdCuentaAtencion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaPendientePagoServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaPendientePagoServicio = oTabla.SeleccionarParaPendientePagoServicio(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaPendientePagoBienInsumo(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaPendientePagoBienInsumo = oTabla.SeleccionarParaPendientePagoBienesInsumos(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                   PREFACTURACION ESTADO DE CUENTA
'-------------------------------------------------------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServiciosObtenerParaEstadoCuenta(lIdCuentaAtencion As Long, TipoServicioARecuperar As sghTipoFacturacionServicio) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionServiciosObtenerParaEstadoCuenta = oTabla.SeleccionarParaEstadoCuenta(lIdCuentaAtencion, TipoServicioARecuperar)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesInsumosObtenerParaEstadoCuenta(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesInsumosObtenerParaEstadoCuenta = oTabla.SeleccionarParaEstadoCuenta(lIdCuentaAtencion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaModificacionCuentaServicio(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaModificacionCuentaServicio = oTabla.SeleccionarParaModificacionDeCuentaServicio(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function EmpleadosSeleccionarParaModificacionCuentaBienInsumo(lIdCuentaAtencion As Long, lIdEmpleadoActual As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Empleados
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set EmpleadosSeleccionarParaModificacionCuentaBienInsumo = oTabla.SeleccionarParaModificacionDeCuentaBienesInsumos(lIdCuentaAtencion, lIdEmpleadoActual)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ObtenerCodigoDeNuevoCarnet() As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    ObtenerCodigoDeNuevoCarnet = oTabla.ObtenerCodigoDeNuevoCarnet()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ObtenerTieneNuevoCarnet(lIdAtencion As Long) As Recordset
    Set ObtenerTieneNuevoCarnet = FactOrdenServicioPagosPorIdAtencionIdProducto(lIdAtencion, ObtenerCodigoDeNuevoCarnet())
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ObtenerCodigoDeDuplicadoCarnet() As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    ObtenerCodigoDeDuplicadoCarnet = oTabla.ObtenerCodigoDeDuplicadoCarnet()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ObtenerDuplicadoCarnet(lIdAtencion As Long) As Recordset
    Set ObtenerDuplicadoCarnet = FactOrdenServicioPagosPorIdAtencionIdProducto(lIdAtencion, ObtenerCodigoDeDuplicadoCarnet())
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ObtenerCodigoDeFolder() As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    ObtenerCodigoDeFolder = oTabla.ObtenerCodigoDeFolder()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ObtenerFolder(lIdAtencion As Long) As Recordset
    Set ObtenerFolder = FactOrdenServicioPagosPorIdAtencionIdProducto(lIdAtencion, ObtenerCodigoDeFolder())
End Function

Function ObtenerTipoDeConsulta(lIdAtencion As Long, lIdTipoConsulta As Long) As Recordset
    Set ObtenerTipoDeConsulta = FactOrdenServicioPagosPorIdAtencionIdProducto(lIdAtencion, lIdTipoConsulta)
End Function

'-------------------------------------------------------------------------------------
'                   CAJA COMPROBANTE
'-------------------------------------------------------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServiciosObtenerParaCaja(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionServiciosObtenerParaCaja = oTabla.SeleccionarParaCaja(lIdCuentaAtencion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesInsumosObtenerParaCaja(lIdCuentaAtencion As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesInsumosObtenerParaCaja = oTabla.SeleccionarParaCaja(lIdCuentaAtencion)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesInsumosSeleccionarPorId(IdFacturacionBienes As Long) As DOFacturacionBienesInsumos
Dim oConexion As New ADODB.Connection
Dim oFacturacionBienesInsumos As New FacturacionBienesInsumos
Dim oDOFacturacionBienesInsumos As New DOFacturacionBienesInsumos
    
    Set FacturacionBienesInsumosSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFacturacionBienesInsumos.Conexion = oConexion
    oDOFacturacionBienesInsumos.IdFacturacionBienes = IdFacturacionBienes
    If oFacturacionBienesInsumos.SeleccionarPorId(oDOFacturacionBienesInsumos) Then
        Set FacturacionBienesInsumosSeleccionarPorId = oDOFacturacionBienesInsumos
    End If
    oConexion.Close

    ms_MensajeError = oFacturacionBienesInsumos.MensajeError
    Set oConexion = Nothing
    Set oFacturacionBienesInsumos = Nothing

End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionSegurosSeleccionarPorId(IdFacturacionSeguro As Long) As DOFacturacionSeguros
Dim oConexion As New ADODB.Connection
Dim oFacturacionSeguro As New FacturacionSeguros
Dim oDOFacturacionSeguro As New DOFacturacionSeguros
    
    Set FacturacionSegurosSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFacturacionSeguro.Conexion = oConexion
    oDOFacturacionSeguro.IdFacturacionSeguro = IdFacturacionSeguro
    If oFacturacionSeguro.SeleccionarPorId(oDOFacturacionSeguro) Then
        Set FacturacionSegurosSeleccionarPorId = oDOFacturacionSeguro
    End If
    oConexion.Close

    ms_MensajeError = oFacturacionSeguro.MensajeError
    Set oConexion = Nothing
    Set oFacturacionSeguro = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ActualizarExoneracionItemsFacturacion(coFacturacionServicios As Collection, coFacturacionBienes As Collection) As Boolean
Dim oConexion As New ADODB.Connection

Dim oDOFacturacionServicios As New DOFacturacionServicios
Dim oDOFacturacionBienesInsumos As New DOFacturacionBienesInsumos
Dim oFacturacionServicios As New FacturacionServicios
Dim oFacturacionBienesInsumos As New FacturacionBienesInsumos

    ActualizarExoneracionItemsFacturacion = False
        
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFacturacionServicios.Conexion = oConexion
    Set oFacturacionBienesInsumos.Conexion = oConexion
    
    For Each oDOFacturacionServicios In coFacturacionServicios
        If Not oFacturacionServicios.ActualizarExoneracion(oDOFacturacionServicios) Then
            Me.MensajeError = oFacturacionServicios.MensajeError
            Exit Function
        End If
    Next
    
    oConexion.Close
    ms_MensajeError = ""
    ActualizarExoneracionItemsFacturacion = True
    Set oConexion = Nothing
    Set oDOFacturacionServicios = Nothing
    Set oDOFacturacionBienesInsumos = Nothing
    Set oFacturacionServicios = Nothing
    Set oFacturacionBienesInsumos = Nothing
End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ActualizarServiciosYBienes(coFacturacionServicios As Collection, coFacturacionBienes As Collection, coFacServiciosBorrar As Collection, coFacBienesBorrar As Collection) As Boolean
Dim oConexion As New ADODB.Connection

Dim oDOFacturacionServicios As New DOFacturacionServicios
Dim oDOFacturacionBienesInsumos As New DOFacturacionBienesInsumos

Dim oFacturacionServicios As New FacturacionServicios
Dim oFacturacionBienesInsumos As New FacturacionBienesInsumos

Dim ok As Boolean
    ActualizarServiciosYBienes = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    oConexion.BeginTrans
    Set oFacturacionServicios.Conexion = oConexion
    Set oFacturacionBienesInsumos.Conexion = oConexion
        
     
     For Each oDOFacturacionServicios In coFacServiciosBorrar
        If oDOFacturacionServicios.idFacturacionServicio > 0 Then
            ok = oFacturacionServicios.Eliminar(oDOFacturacionServicios)
        End If
        
        If Not ok Then
            oConexion.RollbackTrans
            oConexion.Close
            Me.MensajeError = oFacturacionServicios.MensajeError
            Exit Function
        End If
    Next
    
    For Each oDOFacturacionServicios In coFacturacionServicios
        If oDOFacturacionServicios.idFacturacionServicio <= 0 Then
            ok = oFacturacionServicios.Insertar(oDOFacturacionServicios)
        Else
            ok = oFacturacionServicios.modificar(oDOFacturacionServicios)
        End If
        
        If Not ok Then
            oConexion.RollbackTrans
            oConexion.Close
            Me.MensajeError = oFacturacionServicios.MensajeError
            Exit Function
        End If
    Next
    
    For Each oDOFacturacionBienesInsumos In coFacBienesBorrar
        If oDOFacturacionBienesInsumos.IdFacturacionBienes > 0 Then
            ok = oFacturacionBienesInsumos.Eliminar(oDOFacturacionBienesInsumos)
        End If
        
        If Not ok Then
            oConexion.RollbackTrans
            oConexion.Close
            Me.MensajeError = oFacturacionBienesInsumos.MensajeError
            Exit Function
        End If
    Next
    
    For Each oDOFacturacionBienesInsumos In coFacturacionBienes
        If oDOFacturacionBienesInsumos.IdFacturacionBienes <= 0 Then
            ok = oFacturacionBienesInsumos.Insertar(oDOFacturacionBienesInsumos)
        Else
            ok = oFacturacionBienesInsumos.modificar(oDOFacturacionBienesInsumos)
        End If
        
        If Not ok Then
            oConexion.RollbackTrans
            oConexion.Close
            Me.MensajeError = oFacturacionBienesInsumos.MensajeError
            Exit Function
        End If
    Next
        
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    ActualizarServiciosYBienes = True
    Set oConexion = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionSeleccionarTiposConsultaInterconsulta() As ADODB.Recordset
'WCG20060313
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    Set oTabla.Conexion = oConexion
    Set FacturacionSeleccionarTiposConsultaInterconsulta = oTabla.SeleccionarTiposConsultaInterconsulta()
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ObtenerCodigoDeConsultaDeEmergencia() As Long
Dim oConexion As New ADODB.Connection
Dim oTabla As New Parametros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    ObtenerCodigoDeConsultaDeEmergencia = oTabla.ObtenerCodigoDeConsultaDeEmergencia()
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionSeleccionarCatalogo(iTipoCatalogo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    Set oTabla.Conexion = oConexion
    
    Select Case iTipoCatalogo
    Case 0
        Set FacturacionSeleccionarCatalogo = oTabla.SeleccionarCatalogoBase()
    Case Else
        Set FacturacionSeleccionarCatalogo = oTabla.SeleccionarCatalogo(iTipoCatalogo)
    End Select
    
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CatalogoServiciosSeleccionarPorTipoCatalogo(oDOCatalogoServicio As DOCatalogoServicio, iTipoCatalogo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    Set oTabla.Conexion = oConexion
    Set CatalogoServiciosSeleccionarPorTipoCatalogo = oTabla.SeleccionarPorTipoCatalogo(oDOCatalogoServicio, iTipoCatalogo)
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function

Function CatalogoServiciosSeleccionarPorId(idCatalogoServicio As Long) As DOCatalogoServicio
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
Dim doTabla As New DOCatalogoServicio

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    doTabla.idProducto = idCatalogoServicio

    Set oTabla.Conexion = oConexion
    If (oTabla.SeleccionarPorId(doTabla)) Then
        Set CatalogoServiciosSeleccionarPorId = doTabla
    Else
        Set CatalogoServiciosSeleccionarPorId = Nothing
    End If
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function


Function CatalogoServiciosHospSeleccionarPorId(idCatalogoServicio As Long, IdTipoFinanciamiento As Long, Optional oConexion1 As Connection) As DOFinanciamientoCatalogoServ
Dim oConexion As New ADODB.Connection
Dim oTabla As New FinanciamientoCatalogoServ
Dim doTabla As New DOFinanciamientoCatalogoServ
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    End If
    doTabla.idProducto = idCatalogoServicio
    doTabla.IdTipoFinanciamiento = IdTipoFinanciamiento

    If oConexion1 Is Nothing Then
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    If (oTabla.SeleccionarPorProductoYTipoFinanciamiento(doTabla)) Then
        Set CatalogoServiciosHospSeleccionarPorId = doTabla
    Else
        Set CatalogoServiciosHospSeleccionarPorId = Nothing
    End If
    If oConexion1 Is Nothing Then
        oConexion.Close
        Set oTabla = Nothing
    End If
    Set oConexion = Nothing
    ms_MensajeError = oTabla.MensajeError
End Function

Function AgregarServiciosAlCatalogo(lIdTipoFinanciamiento As Long, rsServicios As Recordset)
Dim oConexion As New ADODB.Connection
Dim oTabla As New FinanciamientoCatalogoServ
Dim oDoTabla As New DOFinanciamientoCatalogoServ

    AgregarServiciosAlCatalogo = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not (rsServicios.EOF And rsServicios.BOF) Then
        rsServicios.MoveFirst
        Do While Not rsServicios.EOF
            oDoTabla.Activo = True
            oDoTabla.idProducto = rsServicios!idProducto
            oDoTabla.IdTipoFinanciamiento = lIdTipoFinanciamiento
            oDoTabla.IdUsuarioAuditoria = rsServicios!idUsuario
            oDoTabla.PrecioUnitario = rsServicios!PrecioUnitario
            If Not oTabla.Insertar(oDoTabla) Then
                Exit Function
            End If
            rsServicios.MoveNext
        Loop
    End If
    
    AgregarServiciosAlCatalogo = True
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function AtencionTieneServiciosFacturados(lIdAtencion As Long)
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios

    AtencionTieneServiciosFacturados = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    AtencionTieneServiciosFacturados = oTabla.AtencionTieneServiciosFacturados(lIdAtencion)
    
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function PacienteSePuedeEliminar(lidPaciente As Long)
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes

    PacienteSePuedeEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    PacienteSePuedeEliminar = oTabla.SePuedeEliminar(lidPaciente)
    
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CatalogoBienesInsumosSeleccionarPorTipoCatalogo(oDOCatalogoBienInsumo As DOCatalogoBienesInsumos, iTipoCatalogo As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    Set oTabla.Conexion = oConexion
    Set CatalogoBienesInsumosSeleccionarPorTipoCatalogo = oTabla.SeleccionarPorTipoCatalogo(oDOCatalogoBienInsumo, iTipoCatalogo)
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AgregarBienesInsumosAlCatalogo(lIdTipoFinanciamiento As Long, rsBienesInsumos As Recordset)
Dim oConexion As New ADODB.Connection
Dim oTabla As New FinanciamientoCatalogoBien
Dim oDoTabla As New DoFinanciamientoCatalogoBien

    AgregarBienesInsumosAlCatalogo = False
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    If Not (rsBienesInsumos.EOF And rsBienesInsumos.BOF) Then
        rsBienesInsumos.MoveFirst
        Do While Not rsBienesInsumos.EOF
            oDoTabla.Activo = True
            oDoTabla.idProducto = rsBienesInsumos!idProducto
            oDoTabla.IdTipoFinanciamiento = lIdTipoFinanciamiento
            oDoTabla.IdUsuarioAuditoria = rsBienesInsumos!idUsuario
            oDoTabla.PrecioUnitario = rsBienesInsumos!PrecioUnitario
            If Not oTabla.Insertar(oDoTabla) Then
                Exit Function
            End If
            rsBienesInsumos.MoveNext
        Loop
    End If
    
    AgregarBienesInsumosAlCatalogo = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'EFGL 2006
'--------------------------------------Caja-----------------------------------------
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function SeleccionarUltimaAtencion(idPaciente As Long, IdCuentaAtencion As Long) As DOAtencion
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones
Dim oDoTabla As New DOAtencion

    Set SeleccionarUltimaAtencion = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set oDoTabla = oTabla.SeleccionarUltimaAtencion(idPaciente, IdCuentaAtencion)
    
    Set SeleccionarUltimaAtencion = oDoTabla
    
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

ms_MensajeError = oTabla.MensajeError

End Function

Function AgregarAtencion(oAtencion As DOAtencion) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New Atenciones

    AgregarAtencion = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.Insertar(oAtencion) Then
       Exit Function
    End If
    
    AgregarAtencion = True
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

ms_MensajeError = oTabla.MensajeError

End Function

Function AgregarFacturacionServicios(oFacturacionServicio As DOFacturacionServicios) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios

    AgregarFacturacionServicios = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.Insertar(oFacturacionServicio) Then
        ms_MensajeError = oTabla.MensajeError
       Exit Function
    End If
    
    AgregarFacturacionServicios = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ModificarFacturacionServicios(oFacturacionServicio As DOFacturacionServicios) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios

    ModificarFacturacionServicios = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.modificar(oFacturacionServicio) Then
       Exit Function
    End If
    
    ModificarFacturacionServicios = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EliminarFacturacionServicios(oFacturacionServicio As DOFacturacionServicios) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios

    EliminarFacturacionServicios = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.Eliminar(oFacturacionServicio) Then
       Exit Function
    End If
    
    EliminarFacturacionServicios = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CatalogoBienesSeleccionarPorId(idCatalogoBienes As Long) As DOCatalogoBienesInsumos
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos
Dim doTabla As New DOCatalogoBienesInsumos

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    doTabla.idProducto = idCatalogoBienes

    Set oTabla.Conexion = oConexion
    If (oTabla.SeleccionarPorId(doTabla)) Then
        Set CatalogoBienesSeleccionarPorId = doTabla
    Else
        Set CatalogoBienesSeleccionarPorId = Nothing
    End If
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function

Function CatalogoBienesHospSeleccionarPorId(idCatalogoBien As Long, IdTipoFinanciamiento As Integer) As DoFinanciamientoCatalogoBien
Dim oConexion As New ADODB.Connection
Dim oTabla As New FinanciamientoCatalogoBien
Dim doTabla As New DoFinanciamientoCatalogoBien

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    doTabla.idProducto = idCatalogoBien
    doTabla.IdTipoFinanciamiento = IdTipoFinanciamiento

    Set oTabla.Conexion = oConexion
    If (oTabla.SeleccionarPorProductoYTipoFinanciamiento(doTabla)) Then
        Set CatalogoBienesHospSeleccionarPorId = doTabla
    Else
        Set CatalogoBienesHospSeleccionarPorId = Nothing
    End If
    oConexion.Close

    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing


End Function

Function AgregarFacturacionBienes(oFacturacionBien As DOFacturacionBienesInsumos) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos

    AgregarFacturacionBienes = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.Insertar(oFacturacionBien) Then
        ms_MensajeError = oTabla.MensajeError
       Exit Function
    End If
    
    AgregarFacturacionBienes = True
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function ModificarFacturacionBienes(oFacturacionBien As DOFacturacionBienesInsumos) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos

    ModificarFacturacionBienes = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.modificar(oFacturacionBien) Then
       Exit Function
    End If
    
    ModificarFacturacionBienes = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function EliminarFacturacionBienes(oFacturacionBien As DOFacturacionBienesInsumos) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos

    EliminarFacturacionBienes = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    
    If Not oTabla.Eliminar(oFacturacionBien) Then
       Exit Function
    End If
    
    EliminarFacturacionBienes = True
    oConexion.Close

ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'------------------------------ Estado de Cuenta -------------------------------
Function CuentaAtencionSeleccionarUltimaPorIdPaciente(lidPaciente As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CuentasAtencion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    
    Set oTabla.Conexion = oConexion
    Set CuentaAtencionSeleccionarUltimaPorIdPaciente = oTabla.ObtenerUltimaCuentaAtencionPorIdPaciente(lidPaciente)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioParaEstadoCuenta(IdCuentaAtencion As Long, IdPuntoCarga As Long, lEstadosFacturacion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioParaEstadoCuenta = oTabla.SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion, IdPuntoCarga, lEstadosFacturacion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function FacturacionServicioPorCodigo(Codigo As String, IdTipoFinanciamiento As Long, oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New FacturacionServicios
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioPorCodigo = oTabla.SeleccionarPorCodigo(Codigo, IdTipoFinanciamiento)
End Function
'***************daniel barrantes**************
'***************Filtros en la tabla "FacturacionServicios" por Codigo
'***************
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioPorCodigoDEBB(Codigo As String, IdTipoFinanciamiento As Long, IdPuntoCarga As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioPorCodigoDEBB = oTabla.SeleccionarPorCodigoDEBB(Codigo, IdTipoFinanciamiento, IdPuntoCarga)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CatalogoServicioSeleccionarPorIdYtipoFinanciamiento(lnIdProducto As Long, IdTipoFinanciamiento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set CatalogoServicioSeleccionarPorIdYtipoFinanciamiento = oTabla.SeleccionarPorIdYtipoFinanciamiento(lnIdProducto, IdTipoFinanciamiento)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function CatalogoBienesSeleccionarPorIdYtipoFinanciamiento(lnIdProducto As Long, IdTipoFinanciamiento As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoBienesInsumos
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set CatalogoBienesSeleccionarPorIdYtipoFinanciamiento = oTabla.SeleccionarPorIdYtipoFinanciamiento(lnIdProducto, IdTipoFinanciamiento)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesParaEstadoCuenta(IdCuentaAtencion As Long, IdEspecialidad As Long, lEstadosFacturacion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesParaEstadoCuenta = oTabla.SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion, IdEspecialidad, lEstadosFacturacion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'***************daniel barrantes**************
'***************Filtrar datos de la tabla "FacturacionBienesInsumos" por Codigo
'***************
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesPorCodigoDEBB(Codigo As String, IdTipoFinanciamiento As Long, lnidPtoCarga As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesPorCodigoDEBB = oTabla.SeleccionarPorCodigoDEBB(Codigo, IdTipoFinanciamiento, lnidPtoCarga)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
Function FacturacionBienesPorCodigo(Codigo As String, IdTipoFinanciamiento As Long, oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New FacturacionBienesInsumos
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesPorCodigo = oTabla.SeleccionarPorCodigo(Codigo, IdTipoFinanciamiento)
    Set oTabla = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionSegurosParaEstadoCuenta(IdCuentaAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionSeguros
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionSegurosParaEstadoCuenta = oTabla.SeleccionarPorCuentaAtencionParaEstadoCuenta(IdCuentaAtencion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function FacturacionServicioPorOrdenAtencion(IdOrden As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioPorOrdenAtencion = oTabla.SeleccionarPorOrdenAtencion(IdOrden, sEstadosFacturacion, sTiposFinanciamiento)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function FacturacionServicioPorCuentaAtencion(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.Open IIf(lIdAgrupador = 0, sighentidades.CadenaConexion, sighentidades.CadenaConexionShape)
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencion(IdCuentaAtencion, sEstadosFacturacion, sTiposFinanciamiento, lIdAgrupador)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'***************daniel barrantes**************
'***************Filtro de datos en "FacturacionServicios" por ID de cuenta Atencion
'***************
Function FacturacionServicioPorCuentaAtencionDEBB(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioPorCuentaAtencionDEBB = oTabla.SeleccionarPorCuentaAtencionDEBB(IdCuentaAtencion, sEstadosFacturacion, sTiposFinanciamiento, lIdAgrupador)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


'-------------------------------------------------------------------------------------
'                                       Facturacion Ordenes
'-------------------------------------------------------------------------------------
Function FactOrdenServicioAgregar(oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, _
                                  mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String, _
                                  lnIdServicioPaciente As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oFacturacionServicioFinan As New FacturacionServicioFin
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim lnPrecioUnitario As Double

    FactOrdenServicioAgregar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    
     ms_Texto = ""
     'Crear orden de servicio
     If oDoFactOrdenServ.IdCuentaAtencion > 0 Then
        oDoFactOrdenServ.IdServicioPaciente = lnIdServicioPaciente 'mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFactOrdenServ.IdCuentaAtencion, Format(oDoFactOrdenServ.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), Format(Now, "hh:mm"))
     Else
        oDoFactOrdenServ.IdServicioPaciente = 0
     End If
     If Not oFactOrdenesServicio.Insertar(oDoFactOrdenServ) Then
         bProcesoOK = False: GoTo ErrorManager
     End If
     '
     Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFactOrdenServ.IdUsuarioAuditoria, "A", oDoFactOrdenServ.IdOrden, "FactOrdenServicio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
     If TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
       '***********Contado ***********
       With oDOFactOrdenServicioPagos
           .FechaCreacion = oDoFactOrdenServ.FechaCreacion
           .IdEstadoFacturacion = 1   'Registrado
           .IdOrden = oDoFactOrdenServ.IdOrden
           .idUsuario = oDoFactOrdenServ.IdUsuarioAuditoria
           .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
       End With
       If Not oFactServicioPagos.Insertar(oDOFactOrdenServicioPagos) Then
           bProcesoOK = False: GoTo ErrorManager
       End If
       ms_Texto = Trim(Str(oDOFactOrdenServicioPagos.idOrdenPago))
     End If
     'Crear Detalle de servicio
     If rsFacturacionProductos.RecordCount = 0 Then
        bProcesoOK = False: GoTo ErrorManager
     End If
     rsFacturacionProductos.MoveFirst
     Do While Not rsFacturacionProductos.EOF
            'Despachos
            lnPrecioUnitario = rsFacturacionProductos.Fields!PrecioUnitario
            With oDoFacturacionServicioDesp
                .cantidad = rsFacturacionProductos.Fields!cantidad
                .IdOrden = oDoFactOrdenServ.IdOrden
                .idProducto = rsFacturacionProductos.Fields!idProducto
                .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                .Precio = lnPrecioUnitario
                .total = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
                .labConfHIS = IIf(IsNull(rsFacturacionProductos.Fields!labConfHIS), "", rsFacturacionProductos.Fields!labConfHIS)
                If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghRegistroAtencionCE Then   'debb-27/05/2015
                    .GrupoHIS = rsFacturacionProductos.Fields!Grupo                          'debb-27/05/2015
                    .SubGrupoHIS = rsFacturacionProductos.Fields!subGrupo                    'debb-27/05/2015
                    .labConfHIS = rsFacturacionProductos.Fields!labConfHIS                   ''debb-27/05/2015
                End If
            End With
            If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                   bProcesoOK = False: GoTo ErrorManager
            End If
            If TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = False Then
                lnPrecioUnitario = 0
                oDoCatalogoServicioHosp.PrecioUnitario = 0
                Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(rsFacturacionProductos.Fields!idProducto, oDoFactOrdenServ.IdTipoFinanciamiento)
                'If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
                    lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                'Else
                '    ms_MensajeError = "Tiene problemas con el ID SERVICIO: " & rsFacturacionProductos.Fields!IdProducto & " (parametros)" & Chr(13) & "para el PRODUCTO/PLAN: " & oDoFactOrdenServ.IdTipoFinanciamiento & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA...."
                '    bProcesoOK = False: GoTo ErrorManager
                'End If
            End If
            If TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
               '***********Contado ***********
               With oDoFacturacionServicioPagos
                   .cantidad = rsFacturacionProductos.Fields!cantidad
                   .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                   .idProducto = rsFacturacionProductos.Fields!idProducto
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .Precio = lnPrecioUnitario
                   .total = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
               End With
               If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                   bProcesoOK = False: GoTo ErrorManager
               End If
            Else
               '***********Sis,Soat,exoneracion, convenio ***********
               With oDoFacturacionServicioFinanc
                   .CantidadFinanciada = rsFacturacionProductos.Fields!cantidad
                   .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                   .IdOrden = oDoFactOrdenServ.IdOrden
                   .idProducto = rsFacturacionProductos.Fields!idProducto
                   .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
                   .PrecioFinanciado = lnPrecioUnitario
                   .TotalFinanciado = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
                   .idFuenteFinanciamiento = oDoFactOrdenServ.idFuenteFinanciamiento
               End With
               If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
                   bProcesoOK = False: GoTo ErrorManager
               End If
            End If
            rsFacturacionProductos.MoveNext
    Loop
            
    ms_MensajeError = ""
    FactOrdenServicioAgregar = True
ErrorManager:
    If FactOrdenServicioAgregar = True Then
        oConexion.CommitTrans
    Else
        ms_MensajeError = Err.Description
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set oFactServicioPagos = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set mo_ReglasFarmacia = Nothing
End Function

Function FactOrdenServicioModificar(oDoFactOrdenServ As DoFactOrdenServ, rsFacturacionProductos As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactOrdenesServicio As New FactOrdenServicio
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oFacturacionServicioPagos As New FacturacionServicioPagos
Dim oFacturacionServicioFinan As New FacturacionServicioFin
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim lnPrecioUnitario As Double
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim oDoFacturacionServicioPagos As New DoFacturacionServicioPagos
Dim oDoFacturacionServicioFinanc As New DoFacturacionServicioFin
Dim oDoFacturacionServicioDesp As New DoFacturacionServicioDespacho
Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFactServicioPagos As New FactOrdenServicioPagos
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    FactOrdenServicioModificar = False
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oFacturacionServicioPagos.Conexion = oConexion
    Set oFacturacionServicioFinan.Conexion = oConexion
    Set oFacturacionServicioDespacho.Conexion = oConexion
    Set oFactOrdenesServicio.Conexion = oConexion
    Set oFactServicioPagos.Conexion = oConexion
    
    ms_Texto = ""
    'Crear orden de servicio
    If oDoFactOrdenServ.IdCuentaAtencion > 0 Then
        oDoFactOrdenServ.IdServicioPaciente = mo_ReglasFarmacia.DevuelveServicioDondeSeEncuentraElPacienteSegunFechaHora(oDoFactOrdenServ.IdCuentaAtencion, Format(oDoFactOrdenServ.FechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY), Format(Now, "hh:mm"))
    Else
        oDoFactOrdenServ.IdServicioPaciente = 0
    End If
    If Not oFactOrdenesServicio.modificar(oDoFactOrdenServ) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
     '
     Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFactOrdenServ.IdUsuarioAuditoria, "M", oDoFactOrdenServ.IdOrden, "FactOrdenServicio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
    
     'Eliminar Detalle de servicio
     Dim oRsTmp1 As New Recordset
     Set oRsTmp1 = mo_ReglasFacturacion.FactOrdenServicioPagosSeleccionarPorIdOrden(oDoFactOrdenServ.IdOrden, oConexion)
     If oRsTmp1.RecordCount > 0 Then
       If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
            oDOFactOrdenServicioPagos.IdOrden = oRsTmp1.Fields!IdOrden
            oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
            oDoFacturacionServicioPagos.idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
            If Not oFacturacionServicioPagos.Eliminar(oDoFacturacionServicioPagos) Then
               bProcesoOK = False: GoTo ErrorManager
            End If
            ms_Texto = Trim(Str(oRsTmp1.Fields!idOrdenPago))
       End If
     End If
     oDoFacturacionServicioFinanc.IdOrden = oDoFactOrdenServ.IdOrden
     oDoFacturacionServicioDesp.IdOrden = oDoFactOrdenServ.IdOrden
     If Not oFacturacionServicioFinan.Eliminar(oDoFacturacionServicioFinanc) Then
       bProcesoOK = False: GoTo ErrorManager
     End If
     If Not oFacturacionServicioDespacho.Eliminar(oDoFacturacionServicioDesp) Then
       bProcesoOK = False: GoTo ErrorManager
     End If
     
     'Crear Detalle de servicio
     If rsFacturacionProductos.RecordCount = 0 Then
        bProcesoOK = False: GoTo ErrorManager
     End If
     rsFacturacionProductos.MoveFirst
     Do While Not rsFacturacionProductos.EOF
            'Despachos
            lnPrecioUnitario = rsFacturacionProductos.Fields!PrecioUnitario
            With oDoFacturacionServicioDesp
                .cantidad = rsFacturacionProductos.Fields!cantidad
                .IdOrden = oDoFactOrdenServ.IdOrden
                .idProducto = rsFacturacionProductos.Fields!idProducto
                .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                .Precio = lnPrecioUnitario
                .total = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
                .labConfHIS = IIf(IsNull(rsFacturacionProductos.Fields!labConfHIS), "", rsFacturacionProductos.Fields!labConfHIS)
            End With
            If Not oFacturacionServicioDespacho.Insertar(oDoFacturacionServicioDesp) Then
                   bProcesoOK = False: GoTo ErrorManager
            End If
            If oDoFactOrdenServ.IdTipoFinanciamiento <> 5 And oDoFactOrdenServ.IdTipoFinanciamiento <> 1 Then
                lnPrecioUnitario = 0
                oDoCatalogoServicioHosp.PrecioUnitario = 0
                Set oDoCatalogoServicioHosp = mo_ReglasFacturacion.CatalogoServiciosHospSeleccionarPorId(rsFacturacionProductos.Fields!idProducto, oDoFactOrdenServ.IdTipoFinanciamiento)
                'If oDoCatalogoServicioHosp.PrecioUnitario > 0 Then
                    lnPrecioUnitario = oDoCatalogoServicioHosp.PrecioUnitario
                'Else
                '    ms_MensajeError = "Tiene problemas con el ID SERVICIO: " & rsFacturacionProductos.Fields!IdProducto & " (parametros)" & Chr(13) & "para el PRODUCTO/PLAN: " & oDoFactOrdenServ.IdTipoFinanciamiento & Chr(13) & "...consulte al ADMINISTRADOR DEL SISTEMA...."
                '    bProcesoOK = False: GoTo ErrorManager
                'End If
            End If
            If TiposFinanciamientoGeneraReciboPago(oDoFactOrdenServ.IdTipoFinanciamiento, oConexion) = True Then
               '***********Contado ***********
               With oDoFacturacionServicioPagos
                   .cantidad = rsFacturacionProductos.Fields!cantidad
                   .idOrdenPago = oDOFactOrdenServicioPagos.idOrdenPago
                   .idProducto = rsFacturacionProductos.Fields!idProducto
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .Precio = lnPrecioUnitario
                   .total = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
               End With
               If Not oFacturacionServicioPagos.Insertar(oDoFacturacionServicioPagos) Then
                   bProcesoOK = False: GoTo ErrorManager
               End If
            Else
               '***********Sis,Soat,exoneracion, convenio ***********
               With oDoFacturacionServicioFinanc
                   .CantidadFinanciada = rsFacturacionProductos.Fields!cantidad
                   .FechaAutoriza = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                   .IdOrden = oDoFactOrdenServ.IdOrden
                   .idProducto = rsFacturacionProductos.Fields!idProducto
                   .IdTipoFinanciamiento = oDoFactOrdenServ.IdTipoFinanciamiento
                   .IdUsuarioAuditoria = oDoFactOrdenServ.IdUsuarioAuditoria
                   .IdUsuarioAutoriza = oDoFactOrdenServ.IdUsuarioAuditoria
                   .PrecioFinanciado = lnPrecioUnitario
                   .TotalFinanciado = rsFacturacionProductos.Fields!cantidad * lnPrecioUnitario
                   .idFuenteFinanciamiento = oDoFactOrdenServ.idFuenteFinanciamiento
               End With
               If Not oFacturacionServicioFinan.Insertar(oDoFacturacionServicioFinanc) Then
                   bProcesoOK = False: GoTo ErrorManager
               End If
            End If
            rsFacturacionProductos.MoveNext
    Loop
    ms_MensajeError = ""
    FactOrdenServicioModificar = True
ErrorManager:
    If FactOrdenServicioModificar = True Then
       oConexion.CommitTrans
    Else
       ms_MensajeError = Err.Description
       oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFactOrdenesServicio = Nothing
    Set oFacturacionServicioDespacho = Nothing
    Set oFacturacionServicioPagos = Nothing
    Set oFacturacionServicioFinan = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set oDoFacturacionServicioPagos = Nothing
    Set oDoFacturacionServicioFinanc = Nothing
    Set oDoFacturacionServicioDesp = Nothing
    Set oDoCatalogoServicioHosp = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oFactServicioPagos = Nothing
    Set mo_ReglasFarmacia = Nothing
End Function

Function FactOrdenServicioEliminar(oDOFactOrdenServicio As DoFactOrdenServ, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPaciente As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicio As New FactOrdenServicio
Dim oRsTmp1 As New Recordset
Dim oDOFactOrdenServicioPagos As New DoFactOrdenServPagos
Dim oFactOrdenServicioPagos  As New FactOrdenServicioPagos
Dim oFacturacionServicioFin As New FacturacionServicioFin

     FactOrdenServicioEliminar = False
     oConexion.Open sighentidades.CadenaConexion
    
     Set oFactOrdenServicio.Conexion = oConexion
     Set oFactOrdenServicioPagos.Conexion = oConexion
     Set oFacturacionServicioFin.Conexion = oConexion
         
     Set oRsTmp1 = FactOrdenServicioPagosSeleccionarPorIdOrden(oDOFactOrdenServicio.IdOrden, oConexion)
     If oRsTmp1.RecordCount > 0 Then
       If oRsTmp1.Fields!IdEstadoFacturacion = 1 Then
          oDOFactOrdenServicioPagos.idOrdenPago = oRsTmp1.Fields!idOrdenPago
          If oFactOrdenServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then
             oDOFactOrdenServicioPagos.IdEstadoFacturacion = 9
             If oFactOrdenServicioPagos.modificar(oDOFactOrdenServicioPagos) Then
                oDOFactOrdenServicio.IdEstadoFacturacion = 9
                If oFactOrdenServicio.modificar(oDOFactOrdenServicio) Then
                    FactOrdenServicioEliminar = True
                Else
                    ms_MensajeError = oFactOrdenServicio.MensajeError
                End If
             Else
                ms_MensajeError = oFactOrdenServicioPagos.MensajeError
             End If
          Else
             ms_MensajeError = oFactOrdenServicioPagos.MensajeError
          End If
       Else
          ms_MensajeError = "Ya fue Pagada"
       End If
    Else
       oDOFactOrdenServicio.IdEstadoFacturacion = 9
       If oFactOrdenServicio.modificar(oDOFactOrdenServicio) Then
          If oFacturacionServicioFin.ActualizaIdEstadoFacturacion(oDOFactOrdenServicio.IdOrden, 9) Then
             FactOrdenServicioEliminar = True
          Else
             ms_MensajeError = oFacturacionServicioFin.MensajeError
          End If
       Else
          ms_MensajeError = oFactOrdenServicio.MensajeError
       End If
    End If
    oRsTmp1.Close
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOFactOrdenServicio.IdUsuarioAuditoria, "E", oDOFactOrdenServicio.IdOrden, "FactOrdenServicio", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPaciente)          'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oDOFactOrdenServicioPagos = Nothing
    Set oRsTmp1 = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oFactOrdenServicio = Nothing
End Function

Function FactOrdenServicioSeleccionarPorId(IdOrden As Long) As DoFactOrdenServ
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicio As New FactOrdenServicio
Dim oDOFactOrdenServicio As New DoFactOrdenServ
    
    Set FactOrdenServicioSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFactOrdenServicio.Conexion = oConexion
    oDOFactOrdenServicio.IdOrden = IdOrden
    If oFactOrdenServicio.SeleccionarPorId(oDOFactOrdenServicio) Then
        Set FactOrdenServicioSeleccionarPorId = oDOFactOrdenServicio
    Else
    
    End If
    oConexion.Close

    ms_MensajeError = oFactOrdenServicio.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing

End Function

Function FactOrdenServicioSeleccionarPorIdComprobante(idComprobante As Long) As DOFactOrdenServicio
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicio As New FactOrdenesServicio
Dim oDOFactOrdenServicio As New DOFactOrdenServicio
    
    Set FactOrdenServicioSeleccionarPorIdComprobante = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFactOrdenServicio.Conexion = oConexion
    Set oDOFactOrdenServicio = oFactOrdenServicio.SeleccionarPorIdComprobante(idComprobante)
    
    Set FactOrdenServicioSeleccionarPorIdComprobante = oDOFactOrdenServicio
    
    oConexion.Close
    ms_MensajeError = oFactOrdenServicio.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing

End Function

Function FactOrdenServicioSeleccionarPorIdComprobanteDevolucion(idComprobante As Long) As DOFactOrdenServicio
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicio As New FactOrdenesServicio
Dim oDOFactOrdenServicio As New DOFactOrdenServicio
    
    Set FactOrdenServicioSeleccionarPorIdComprobanteDevolucion = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFactOrdenServicio.Conexion = oConexion
    Set oDOFactOrdenServicio = oFactOrdenServicio.SeleccionarPorIdComprobanteDevolucion(idComprobante)
    
    Set FactOrdenServicioSeleccionarPorIdComprobanteDevolucion = oDOFactOrdenServicio
    
    oConexion.Close
    ms_MensajeError = oFactOrdenServicio.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenServicio = Nothing

End Function

Function FactOrdenBienInsumoSeleccionarPorIdComprobante(idComprobante As Long) As DOFactOrdenBienInsumo
Dim oConexion As New ADODB.Connection
Dim oFactOrdenesBienesInsumo As New FactOrdenesBienesInsumo
Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
    
    
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFactOrdenesBienesInsumo.Conexion = oConexion
    Set oDOFactOrdenBienInsumo = oFactOrdenesBienesInsumo.SeleccionarPorIdComprobante(idComprobante)
    
    Set FactOrdenBienInsumoSeleccionarPorIdComprobante = oDOFactOrdenBienInsumo
    
    oConexion.Close
    ms_MensajeError = oFactOrdenesBienesInsumo.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenesBienesInsumo = Nothing

End Function

Function FactOrdenBienInsumoSeleccionarPorIdComprobanteDevolucion(idComprobante As Long) As DOFactOrdenBienInsumo
Dim oConexion As New ADODB.Connection
Dim oFactOrdenesBienesInsumo As New FactOrdenesBienesInsumo
Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
    
    Set FactOrdenBienInsumoSeleccionarPorIdComprobanteDevolucion = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oFactOrdenesBienesInsumo.Conexion = oConexion
    Set oDOFactOrdenBienInsumo = oFactOrdenesBienesInsumo.SeleccionarPorIdComprobanteDevolucion(idComprobante)
    
    Set FactOrdenBienInsumoSeleccionarPorIdComprobanteDevolucion = oDOFactOrdenBienInsumo
    
    oConexion.Close
    ms_MensajeError = oFactOrdenesBienesInsumo.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenesBienesInsumo = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       Facturacion Ordenes Bien e Insumos
'-------------------------------------------------------------------------------------
Function FactOrdenBienInsumoAgregar(oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, rsFacturacionProductos As Recordset, mo_ProductosEliminados As Collection, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDOFacturacionBienesInsumos As New DOFacturacionBienesInsumos
Dim oFacturacionBienes As New FacturacionBienesInsumos
Dim oFactOrdenesBienInsumo As New FactOrdenesBienesInsumo
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim PrSeguro As Double
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    FactOrdenBienInsumoAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oConexion.BeginTrans
    Set oFacturacionBienes.Conexion = oConexion
    Set oFactOrdenesBienInsumo.Conexion = oConexion
    
    'Crear orden de servicio
    If Not oFactOrdenesBienInsumo.Insertar(oDOFactOrdenBienInsumo) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
    
    'Cargar los servicios agregados y modificados
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
        oDOFacturacionBienesInsumos.IdFacturacionBienes = IIf(IsNull(rsFacturacionProductos!IdFacturacionProducto), 0, rsFacturacionProductos!IdFacturacionProducto)
        oDOFacturacionBienesInsumos.IdOrden = oDOFactOrdenBienInsumo.IdOrden
        oDOFacturacionBienesInsumos.idProducto = IIf(IsNull(rsFacturacionProductos!idProducto), 0, rsFacturacionProductos!idProducto)
        oDOFacturacionBienesInsumos.PrecioUnitario = IIf(IsNull(rsFacturacionProductos!PrecioUnitario), 0, rsFacturacionProductos!PrecioUnitario)
        oDOFacturacionBienesInsumos.cantidad = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
        oDOFacturacionBienesInsumos.IdEstadoFacturacion = IIf(IsNull(rsFacturacionProductos!IdEstadoFacturacion), 0, rsFacturacionProductos!IdEstadoFacturacion)
        oDOFacturacionBienesInsumos.idFuenteFinanciamiento = IIf(IsNull(rsFacturacionProductos!idFuenteFinanciamiento), 0, rsFacturacionProductos!idFuenteFinanciamiento)
        oDOFacturacionBienesInsumos.IdTipoFinanciamiento = IIf(IsNull(rsFacturacionProductos!IdTipoFinanciamiento), 0, rsFacturacionProductos!IdTipoFinanciamiento)
        oDOFacturacionBienesInsumos.IdPuntoCarga = IIf(IsNull(rsFacturacionProductos!IdPuntoCarga), 0, rsFacturacionProductos!IdPuntoCarga)
        oDOFacturacionBienesInsumos.IdUsuarioAutorizaSeguro = IIf(IsNull(rsFacturacionProductos!IdUsuarioAutorizaSeguro), 0, rsFacturacionProductos!IdUsuarioAutorizaSeguro)
        oDOFacturacionBienesInsumos.IdUsuarioAutorizaPendiente = IIf(IsNull(rsFacturacionProductos!IdUsuarioAutorizaPendiente), 0, rsFacturacionProductos!IdUsuarioAutorizaPendiente)
        oDOFacturacionBienesInsumos.IdCajero = IIf(IsNull(rsFacturacionProductos!IdCajero), 0, rsFacturacionProductos!IdCajero)
        oDOFacturacionBienesInsumos.FechaAutorizaSeguro = IIf(IsNull(rsFacturacionProductos!FechaAutorizaSeguro), 0, rsFacturacionProductos!FechaAutorizaSeguro)
        oDOFacturacionBienesInsumos.FechaCajero = IIf(IsNull(rsFacturacionProductos!FechaCajero), 0, rsFacturacionProductos!FechaCajero)
        oDOFacturacionBienesInsumos.idAtencion = IIf(IsNull(rsFacturacionProductos!idAtencion), 0, rsFacturacionProductos!idAtencion)
        oDOFacturacionBienesInsumos.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        oDOFacturacionBienesInsumos.IdComprobantePago = IIf(IsNull(rsFacturacionProductos!IdComprobantePago), 0, rsFacturacionProductos!IdComprobantePago)
        oDOFacturacionBienesInsumos.IdComprobantePagoDevolucion = IIf(IsNull(rsFacturacionProductos!IdComprobantePagoDevolucion), 0, rsFacturacionProductos!IdComprobantePagoDevolucion)
        '*****************Solo si tiene seguros=sis/soat (inicio)
        oDOFacturacionBienesInsumos.IdTipoFinanciamiento = oDOFactOrdenBienInsumo.IdFormaPago
        oDOFacturacionBienesInsumos.CantidadSIS = 0
        oDOFacturacionBienesInsumos.PrecioSIS = 0
        oDOFacturacionBienesInsumos.ImporteSIS = 0
        oDOFacturacionBienesInsumos.CantidadSOAT = 0
        oDOFacturacionBienesInsumos.PrecioSOAT = 0
        oDOFacturacionBienesInsumos.ImporteSOAT = 0
        oDOFacturacionBienesInsumos.importeEXO = 0
        oDOFacturacionBienesInsumos.cantidadConv = 0
        oDOFacturacionBienesInsumos.precConv = 0
        oDOFacturacionBienesInsumos.ImporteConv = 0
        PrSeguro = 0
        
   
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenBienInsumoConsultar"
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, oDOFactOrdenBienInsumo.IdFormaPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oDOFacturacionBienesInsumos.idProducto): .Parameters.Append oParameter
        Set oRsBuscaSeguro = .Execute
        Set oRsBuscaSeguro.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
        
         
        If oRsBuscaSeguro.RecordCount > 0 Then
           PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
        End If
        oRsBuscaSeguro.Close
        Set oRsBuscaSeguro = Nothing
        Select Case oDOFactOrdenBienInsumo.IdFormaPago
        Case 2       'SIS
            If PrSeguro > 0 Then
                oDOFacturacionBienesInsumos.CantidadSIS = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
                oDOFacturacionBienesInsumos.PrecioSIS = PrSeguro
                oDOFacturacionBienesInsumos.ImporteSIS = oDOFacturacionBienesInsumos.CantidadSIS * PrSeguro
                'oDOFacturacionBienesInsumos.CantidadPagar = oDOFacturacionBienesInsumos.cantidad
                'oDOFacturacionBienesInsumos.cantidad = 0
                oDOFacturacionBienesInsumos.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()

            End If
        Case 3         'SOAT
            If PrSeguro > 0 Then
                oDOFacturacionBienesInsumos.CantidadSOAT = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
                oDOFacturacionBienesInsumos.PrecioSOAT = PrSeguro
                oDOFacturacionBienesInsumos.ImporteSOAT = oDOFacturacionBienesInsumos.CantidadSOAT * PrSeguro
                'oDOFacturacionBienesInsumos.CantidadPagar = oDOFacturacionBienesInsumos.cantidad
                'oDOFacturacionBienesInsumos.cantidad = 0
                oDOFacturacionBienesInsumos.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            End If
        Case 4  'Convenios
            If PrSeguro > 0 Then
               oDOFacturacionBienesInsumos.cantidadConv = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
               oDOFacturacionBienesInsumos.precConv = PrSeguro
               oDOFacturacionBienesInsumos.ImporteConv = oDOFacturacionBienesInsumos.cantidadConv * PrSeguro
               oDOFacturacionBienesInsumos.PrecioUnitario = PrSeguro
               oDOFacturacionBienesInsumos.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            End If
        Case 9  'Exonerados
             oDOFacturacionBienesInsumos.importeEXO = oDOFacturacionBienesInsumos.cantidad * oDOFacturacionBienesInsumos.PrecioUnitario
             oDOFacturacionBienesInsumos.IdTipoFinanciamiento = 1
             'oDOFacturacionBienesInsumos.CantidadPagar = oDOFacturacionBienesInsumos.cantidad
             'oDOFacturacionBienesInsumos.cantidad = 0
             oDOFacturacionBienesInsumos.FechaAutorizaEXO = lcBuscaParametro.RetornaFechaHoraServidorSQL()
        Case Else
             oDOFacturacionBienesInsumos.IdTipoFinanciamiento = 1
        End Select
        '*****************Solo si tiene seguros=sis/soat (fin)
        Select Case rsFacturacionProductos!EstadoLocal
        Case "A"
            If Not oFacturacionBienes.InsertarBienInsumoSP(oDOFacturacionBienesInsumos, sResultado) Then
                MsgBox sResultado, vbInformation, "Facturacion Servicios"
                bProcesoOK = False: GoTo ErrorManager
            End If
        Case "M"
            If Not oFacturacionBienes.ModificarBienInsumoSP(oDOFacturacionBienesInsumos, sResultado) Then
                MsgBox sResultado, vbInformation, "Facturacion Servicios"
                bProcesoOK = False: GoTo ErrorManager
            End If
        End Select
        rsFacturacionProductos.MoveNext
    Loop
    
    'elimina servicios
    Dim I As Integer
    Dim lIdFacturacionServicio As Long
    For I = 1 To mo_ProductosEliminados.Count
        lIdFacturacionServicio = Val(mo_ProductosEliminados.Item(I))
        oDOFacturacionBienesInsumos.IdFacturacionBienes = lIdFacturacionServicio
        oDOFacturacionBienesInsumos.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        If Not oFacturacionBienes.Eliminar(oDOFacturacionBienesInsumos) Then
            bProcesoOK = False: GoTo ErrorManager
        End If
    Next
            
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    FactOrdenBienInsumoAgregar = True
    Set oConexion = Nothing
    Set oDOFacturacionBienesInsumos = Nothing
    Set oFacturacionBienes = Nothing
    Set oFactOrdenesBienInsumo = Nothing
    Set oRsBuscaSeguro = Nothing

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    FactOrdenBienInsumoAgregar = False
    
    Set oConexion = Nothing
    Set oDOFacturacionBienesInsumos = Nothing
    Set oFacturacionBienes = Nothing
    Set oFactOrdenesBienInsumo = Nothing
    Set oRsBuscaSeguro = Nothing

End Function

Function FactOrdenBienInsumoModificar(oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo, rsFacturacionProductos As Recordset, mo_ProductosEliminados As Collection, ml_IdUsuarioAuditoria As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oDoFacturacionBienes As New DOFacturacionBienesInsumos
Dim oFacturacionBienes As New FacturacionBienesInsumos
Dim oFactOrdenesBienInsumo As New FactOrdenesBienesInsumo
Dim bProcesoOK As Boolean
Dim sResultado As String
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim PrSeguro As Double
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion

    FactOrdenBienInsumoModificar = False

    
    oConexion.BeginTrans
    Set oFacturacionBienes.Conexion = oConexion
    Set oFactOrdenesBienInsumo.Conexion = oConexion
    
    'Crear orden de servicio
    
    oDOFactOrdenBienInsumo.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
    If Not oFactOrdenesBienInsumo.modificar(oDOFactOrdenBienInsumo) Then
        bProcesoOK = False: GoTo ErrorManager
    End If
    
    'Cargar los servicios agregados y modificados
    rsFacturacionProductos.MoveFirst
    Do While Not rsFacturacionProductos.EOF
    
        If rsFacturacionProductos!EstadoLocal = "A" Or rsFacturacionProductos!EstadoLocal = "M" Then
            oDoFacturacionBienes.IdFacturacionBienes = IIf(IsNull(rsFacturacionProductos!IdFacturacionProducto), 0, rsFacturacionProductos!IdFacturacionProducto)
            oDoFacturacionBienes.IdOrden = oDOFactOrdenBienInsumo.IdOrden
            oDoFacturacionBienes.idProducto = IIf(IsNull(rsFacturacionProductos!idProducto), 0, rsFacturacionProductos!idProducto)
            oDoFacturacionBienes.PrecioUnitario = IIf(IsNull(rsFacturacionProductos!PrecioUnitario), 0, rsFacturacionProductos!PrecioUnitario)
            oDoFacturacionBienes.cantidad = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
            oDoFacturacionBienes.IdEstadoFacturacion = IIf(IsNull(rsFacturacionProductos!IdEstadoFacturacion), 0, rsFacturacionProductos!IdEstadoFacturacion)
            oDoFacturacionBienes.idFuenteFinanciamiento = IIf(IsNull(rsFacturacionProductos!idFuenteFinanciamiento), 0, rsFacturacionProductos!idFuenteFinanciamiento)
            oDoFacturacionBienes.IdTipoFinanciamiento = IIf(IsNull(rsFacturacionProductos!IdTipoFinanciamiento), 0, rsFacturacionProductos!IdTipoFinanciamiento)
            'oDoFacturacionBienes.IdServicioInternamiento = IIf(IsNull(rsFacturacionProductos!IdServicioInternamiento), 0, rsFacturacionProductos!IdServicioInternamiento)
            oDoFacturacionBienes.IdPuntoCarga = IIf(IsNull(rsFacturacionProductos!IdPuntoCarga), 0, rsFacturacionProductos!IdPuntoCarga)
            oDoFacturacionBienes.IdUsuarioAutorizaSeguro = IIf(IsNull(rsFacturacionProductos!IdUsuarioAutorizaSeguro), 0, rsFacturacionProductos!IdUsuarioAutorizaSeguro)
            oDoFacturacionBienes.IdUsuarioAutorizaPendiente = IIf(IsNull(rsFacturacionProductos!IdUsuarioAutorizaPendiente), 0, rsFacturacionProductos!IdUsuarioAutorizaPendiente)
            oDoFacturacionBienes.IdUsuarioAutorizaDevolucion = IIf(IsNull(rsFacturacionProductos!IdUsuarioAutorizaDevolucion), 0, rsFacturacionProductos!IdUsuarioAutorizaDevolucion)
            oDoFacturacionBienes.IdCajero = IIf(IsNull(rsFacturacionProductos!IdCajero), 0, rsFacturacionProductos!IdCajero)
            oDoFacturacionBienes.FechaAutorizaSeguro = IIf(IsNull(rsFacturacionProductos!FechaAutorizaSeguro), 0, rsFacturacionProductos!FechaAutorizaSeguro)
            oDoFacturacionBienes.FechaAutorizaPendiente = IIf(IsNull(rsFacturacionProductos!FechaAutorizaPendiente), 0, rsFacturacionProductos!FechaAutorizaPendiente)
            oDoFacturacionBienes.FechaAutorizaDevolucion = IIf(IsNull(rsFacturacionProductos!FechaAutorizaDevolucion), 0, rsFacturacionProductos!FechaAutorizaDevolucion)
            oDoFacturacionBienes.FechaCajero = IIf(IsNull(rsFacturacionProductos!FechaCajero), 0, rsFacturacionProductos!FechaCajero)
            oDoFacturacionBienes.idAtencion = IIf(IsNull(rsFacturacionProductos!idAtencion), 0, rsFacturacionProductos!idAtencion)
            oDoFacturacionBienes.IdComprobantePago = IIf(IsNull(rsFacturacionProductos!IdComprobantePago), 0, rsFacturacionProductos!IdComprobantePago)
            oDoFacturacionBienes.IdComprobantePagoDevolucion = IIf(IsNull(rsFacturacionProductos!IdComprobantePagoDevolucion), 0, rsFacturacionProductos!IdComprobantePagoDevolucion)
            oDoFacturacionBienes.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
            '*****************Solo si tiene seguros=sis/soat (inicio)
            oDoFacturacionBienes.IdTipoFinanciamiento = oDOFactOrdenBienInsumo.IdFormaPago
            oDoFacturacionBienes.CantidadSIS = 0
            oDoFacturacionBienes.PrecioSIS = 0
            oDoFacturacionBienes.ImporteSIS = 0
            oDoFacturacionBienes.CantidadSOAT = 0
            oDoFacturacionBienes.PrecioSOAT = 0
            oDoFacturacionBienes.ImporteSOAT = 0
            oDoFacturacionBienes.importeEXO = 0
            oDoFacturacionBienes.cantidadConv = 0
            oDoFacturacionBienes.precConv = 0
            oDoFacturacionBienes.ImporteConv = 0
            PrSeguro = 0

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenBienInsumoConsultar"
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, oDOFactOrdenBienInsumo.IdFormaPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oRsBuscaSeguro = .Execute
        Set oRsBuscaSeguro.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
            
            If oRsBuscaSeguro.RecordCount > 0 Then
               PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
            End If
            oRsBuscaSeguro.Close
            Set oRsBuscaSeguro = Nothing
            Select Case oDOFactOrdenBienInsumo.IdFormaPago
            Case 2       'SIS
                If PrSeguro > 0 Then
                    oDoFacturacionBienes.CantidadSIS = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
                    oDoFacturacionBienes.PrecioSIS = PrSeguro
                    oDoFacturacionBienes.ImporteSIS = oDoFacturacionBienes.CantidadSIS * PrSeguro
                    'oDoFacturacionBienes.CantidadPagar = oDoFacturacionBienes.cantidad
                    'oDoFacturacionBienes.cantidad = 0
                    oDoFacturacionBienes.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                End If
            Case 3         'SOAT
                If PrSeguro > 0 Then
                    oDoFacturacionBienes.CantidadSOAT = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
                    oDoFacturacionBienes.PrecioSOAT = PrSeguro
                    oDoFacturacionBienes.ImporteSOAT = oDoFacturacionBienes.CantidadSOAT * PrSeguro
                    'oDoFacturacionBienes.CantidadPagar = oDoFacturacionBienes.cantidad
                    'oDoFacturacionBienes.cantidad = 0
                    oDoFacturacionBienes.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                End If
            Case 4  'Convenios
                If PrSeguro > 0 Then
                   oDoFacturacionBienes.cantidadConv = IIf(IsNull(rsFacturacionProductos!cantidad), 0, rsFacturacionProductos!cantidad)
                   oDoFacturacionBienes.precConv = PrSeguro
                   oDoFacturacionBienes.ImporteConv = oDoFacturacionBienes.cantidadConv * PrSeguro
                   oDoFacturacionBienes.PrecioUnitario = PrSeguro
                   oDoFacturacionBienes.FechaAutorizaSeguro = lcBuscaParametro.RetornaFechaHoraServidorSQL()
                End If
            Case 9  'Exonerados
                 oDoFacturacionBienes.importeEXO = oDoFacturacionBienes.cantidad * oDoFacturacionBienes.PrecioUnitario
                 oDoFacturacionBienes.IdTipoFinanciamiento = 1
                 'oDoFacturacionBienes.CantidadPagar = oDoFacturacionBienes.cantidad
                 'oDoFacturacionBienes.cantidad = 0
                 oDoFacturacionBienes.FechaAutorizaEXO = lcBuscaParametro.RetornaFechaHoraServidorSQL()
            Case Else
                 oDoFacturacionBienes.IdTipoFinanciamiento = 1
            End Select
            '*****************Solo si tiene seguros=sis/soat (fin)
            Select Case rsFacturacionProductos!EstadoLocal
            Case "A"
                If Not oFacturacionBienes.InsertarBienInsumoSP(oDoFacturacionBienes, sResultado) Then
                    If sResultado <> "" Then
                        MsgBox sResultado, vbInformation, "Facturacion Servicios"
                    End If
                    bProcesoOK = False: GoTo ErrorManager
                End If
            Case "M"
                If Not oFacturacionBienes.ModificarBienInsumoSP(oDoFacturacionBienes, sResultado) Then
                    MsgBox sResultado, vbInformation, "Facturacion Servicios"
                    bProcesoOK = False: GoTo ErrorManager
                End If
            End Select
        End If
        
        rsFacturacionProductos.MoveNext
    Loop
    
    'elimina servicios
    Dim I As Integer
    Dim lIdFacturacionBien As Long
    For I = 1 To mo_ProductosEliminados.Count
        lIdFacturacionBien = Val(mo_ProductosEliminados.Item(I))
        oDoFacturacionBienes.IdFacturacionBienes = lIdFacturacionBien
        oDoFacturacionBienes.IdUsuarioAuditoria = ml_IdUsuarioAuditoria
        If Not oFacturacionBienes.Eliminar(oDoFacturacionBienes) Then
            bProcesoOK = False: GoTo ErrorManager
        End If
    Next
            
    oConexion.CommitTrans
    oConexion.Close
    ms_MensajeError = ""
    FactOrdenBienInsumoModificar = True
    Set oConexion = Nothing
    Set oFacturacionBienes = Nothing
    Set oFactOrdenesBienInsumo = Nothing
    Set oRsBuscaSeguro = Nothing

    Exit Function
ErrorManager:
    oConexion.RollbackTrans
    oConexion.Close
    FactOrdenBienInsumoModificar = False
    Set oConexion = Nothing
    Set oFacturacionBienes = Nothing
    Set oFactOrdenesBienInsumo = Nothing
    Set oRsBuscaSeguro = Nothing
    
End Function

Function FactOrdenBienInsumoEliminar(oDOFactOrdenBienInsumo As DOFactOrdenBienInsumo) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactOrdenBienInsumo As New FactOrdenesBienesInsumo

    FactOrdenBienInsumoEliminar = False

    oConexion.Open sighentidades.CadenaConexion
    Set oFactOrdenBienInsumo.Conexion = oConexion
    If oFactOrdenBienInsumo.EliminarPorIdOrden(oDOFactOrdenBienInsumo.IdOrden) Then
        FactOrdenBienInsumoEliminar = True
    Else
        ms_MensajeError = oFactOrdenBienInsumo.MensajeError
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFactOrdenBienInsumo = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FactOrdenBienInsumoSeleccionarPorId(IdOrden As Long) As DOFactOrdenBienInsumo
Dim oConexion As New ADODB.Connection
Dim oFactOrdenBienInsumo As New FactOrdenesBienesInsumo
Dim oDOFactOrdenBienInsumo As New DOFactOrdenBienInsumo
    
    Set FactOrdenBienInsumoSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFactOrdenBienInsumo.Conexion = oConexion
    oDOFactOrdenBienInsumo.IdOrden = IdOrden
    If oFactOrdenBienInsumo.SeleccionarPorId(oDOFactOrdenBienInsumo) Then
        Set FactOrdenBienInsumoSeleccionarPorId = oDOFactOrdenBienInsumo
    End If
    oConexion.Close

    ms_MensajeError = oFactOrdenBienInsumo.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenBienInsumo = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13
Function FactOrdenesBienesInsumoSeleccionarPorId(IdOrden As Long) As DoFactOrdenesBienes
Dim oConexion As New ADODB.Connection
Dim oFactOrdenBienInsumo As New FactOrdenesBienes
Dim oDOFactOrdenBienInsumo As New DoFactOrdenesBienes
    
    Set FactOrdenesBienesInsumoSeleccionarPorId = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFactOrdenBienInsumo.Conexion = oConexion
    oDOFactOrdenBienInsumo.IdOrden = IdOrden
    If oFactOrdenBienInsumo.SeleccionarPorId(oDOFactOrdenBienInsumo) Then
        Set FactOrdenesBienesInsumoSeleccionarPorId = oDOFactOrdenBienInsumo
    End If
    oConexion.Close

    ms_MensajeError = oFactOrdenBienInsumo.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenBienInsumo = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenesBienesInsumoSeleccionarPorIdPreVenta(lnIdPreventa As Long) As DoFactOrdenesBienes
Dim oConexion As New ADODB.Connection
Dim oFactOrdenBienInsumo As New FactOrdenesBienes
Dim oDOFactOrdenBienInsumo As New DoFactOrdenesBienes
    
    Set FactOrdenesBienesInsumoSeleccionarPorIdPreVenta = Nothing
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFactOrdenBienInsumo.Conexion = oConexion
    oDOFactOrdenBienInsumo.idPreVenta = lnIdPreventa
    If oFactOrdenBienInsumo.SeleccionarPorIdpreventa(oDOFactOrdenBienInsumo) Then
        Set FactOrdenesBienesInsumoSeleccionarPorIdPreVenta = oDOFactOrdenBienInsumo
    End If
    oConexion.Close

    ms_MensajeError = oFactOrdenBienInsumo.MensajeError
    Set oConexion = Nothing
    Set oFactOrdenBienInsumo = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienInsumoPorOrdenAtencion(IdOrden As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set FacturacionBienInsumoPorOrdenAtencion = oTabla.SeleccionarPorOrdenAtencion(IdOrden, sEstadosFacturacion, sTiposFinanciamiento)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienInsumoPorCuentaAtencion(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.Open IIf(lIdAgrupador = 0, sighentidades.CadenaConexion, sighentidades.CadenaConexionShape)
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienInsumoPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencion(IdCuentaAtencion, sEstadosFacturacion, sTiposFinanciamiento, lIdAgrupador)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'***************daniel barrantes**************
'***************Filtrar datos de "FacturacionBienesInsumos" por Id Cuenta de Atencion
'***************
Function FacturacionBienInsumoPorCuentaAtencionDEBB(IdCuentaAtencion As Long, sEstadosFacturacion As String, sTiposFinanciamiento As String, lIdAgrupador As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    'oConexion.Open IIf(lIdAgrupador = 0, sighentidades.CadenaConexion, sighentidades.CadenaConexionShape)
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionBienInsumoPorCuentaAtencionDEBB = oTabla.SeleccionarPorCuentaAtencionDEBB(IdCuentaAtencion, sEstadosFacturacion, sTiposFinanciamiento, lIdAgrupador)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenesServicioPorCuentaAtencion(IdCuentaAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesServicio
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FactOrdenesServicioPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencion(IdCuentaAtencion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenesBienInsumoPorCuentaAtencion(IdCuentaAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FactOrdenesBienesInsumo
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set FactOrdenesBienInsumoPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencion(IdCuentaAtencion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioDevolucionPorIdComprobante(lIdComprobante As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicios
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioDevolucionPorIdComprobante = oTabla.SeleccionarDevolucionPorIdComprobante(lIdComprobante)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienInsumoDevolucionPorIdComprobante(lIdComprobante As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesInsumos
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oTabla.Conexion = oConexion
    Set FacturacionBienInsumoDevolucionPorIdComprobante = oTabla.SeleccionarDevolucionPorIdComprobante(lIdComprobante)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function


Function CuentasAtencionAnulada(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oCuentasAtencion As New CuentasAtencion
Dim oRsTmp As New ADODB.Recordset
Dim sSql As String
    On Error GoTo ErrCierre
    CuentasAtencionAnulada = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oCuentasAtencion.Conexion = oConexion
    oConexion.BeginTrans
    If oCuentasAtencion.Anulada(oDOCuentaAtencion) = True Then
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasActualizarXIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDOCuentaAtencion.idPaciente): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "M", oDOCuentaAtencion.IdCuentaAtencion, "Anuló Cta", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)           'ListBarItems.idListItem
    
    oConexion.CommitTrans
    oConexion.Close
    CuentasAtencionAnulada = True
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
    Exit Function
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
End Function

Function CuentasAtencionPagada(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCuentasAtencion As New CuentasAtencion
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsTmp As New ADODB.Recordset
Dim sSql As String
    On Error GoTo ErrCierre
    CuentasAtencionPagada = False

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set oCuentasAtencion.Conexion = oConexion
    oConexion.BeginTrans
    If oCuentasAtencion.Pagada(oDOCuentaAtencion) = True Then
    Else
        ms_MensajeError = oCuentasAtencion.MensajeError
        Exit Function
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasActualizarXIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDOCuentaAtencion.idPaciente): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
        '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "A", oDOCuentaAtencion.IdCuentaAtencion, "Pasó a Pagada", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)             'ListBarItems.idListItem
    oConexion.CommitTrans
    oConexion.Close
    CuentasAtencionPagada = True
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
    Exit Function
ErrCierre:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oCuentasAtencion = Nothing
    Set oRsTmp = Nothing
End Function

Function FarmMovimientoVentasPorCuentaAtencion(IdCuentaAtencion As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New farmMovimientoVentas
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FarmMovimientoVentasPorCuentaAtencion = oTabla.SeleccionarPorCuentaAtencion(IdCuentaAtencion)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesFinanciamSeleccionarPorIdProducto(lcMovNumero As String, lcMovTipo As String, _
                                               idProducto As Long, Optional oConexion1 As Connection) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesFinanciam
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CommandTimeout = 150
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    Set FacturacionBienesFinanciamSeleccionarPorIdProducto = oTabla.SeleccionarPorIdProducto(lcMovNumero, lcMovTipo, idProducto)
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function
Function FacturacionBienesPagosSeleccionarPorIdProducto(lcMovNumero As String, lcMovTipo As String, idProducto As Long, Optional oConexion1 As Connection) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionBienesPagos
    If oConexion1 Is Nothing Then
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 150
       Set oTabla.Conexion = oConexion
    Else
       Set oTabla.Conexion = oConexion1
    End If
    Set FacturacionBienesPagosSeleccionarPorIdProducto = oTabla.SeleccionarPorMovNumeroIdProducto(lcMovNumero, lcMovTipo, idProducto)
    Set oConexion = Nothing
    Set oTabla = Nothing
    
End Function

Function FacturacionBienesDevolucionesSeleccionarPorIdProducto(lcMovNumero As String, lcMovTipo As String, idProducto As Long, oConexion As Connection) As ADODB.Recordset
    Dim oTabla As New FacturacionBienesDevol
    Set oTabla.Conexion = oConexion
    Set FacturacionBienesDevolucionesSeleccionarPorIdProducto = oTabla.SeleccionarPorIdProducto(lcMovNumero, lcMovTipo, idProducto)
End Function


Function FactOrdenServicioFiltraPorIdCuenta(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion1 As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioFiltraPorIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion1.CommandTimeout = 300
    oConexion1.Open sighentidades.CadenaConexion
    oConexion1.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion1
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioFiltraPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioFiltraPorIdCuenta = oRecordset
   oConexion1.Close
   Set oConexion1 = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionServicioDespachoSeleccionarPorCuenta(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioDespachoSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDespachoSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioDespachoSeleccionarPorCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionServicioFinanciamSeleccionarPorIdOrdenIdProducto(lnIdOrden As Long, idProducto As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicioFin
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioFinanciamSeleccionarPorIdOrdenIdProducto = oTabla.SeleccionarPorIdOrdenIdProducto(lnIdOrden, idProducto)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function FacturacionServicioPagosSeleccionarPorIdOrdenIdProducto(lnIdOrden As Long, idProducto As Long, Optional oConexion1 As Connection) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicioPagos
    If oConexion1 Is Nothing Then
        oConexion.Open sighentidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 150
        Set oTabla.Conexion = oConexion
    Else
        Set oTabla.Conexion = oConexion1
    End If
    Set FacturacionServicioPagosSeleccionarPorIdOrdenIdProducto = oTabla.SeleccionarPorIdOrdenIdProducto(lnIdOrden, idProducto)
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProducto(lnIdOrden As Long, idProducto As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New FacturacionServicioDev
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    Set oTabla.Conexion = oConexion
    Set FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProducto = oTabla.SeleccionarPorIdOrdenIdProducto(lnIdOrden, idProducto)
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function FacturacionServicioFinanciamientosFiltraPorIdOrden(lidOrden As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioFinanciamientosFiltraPorIdOrden = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosFiltraPorIdOrden"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lidOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioFinanciamientosFiltraPorIdOrden = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionServicioPagosFiltraPorIdOrden(lidOrden As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosFiltraPorIdOrden = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosFiltraPorIdOrden"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lidOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosFiltraPorIdOrden = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPorFechas(ldFechaIni As Date, ldFechaFin As Date, lnIdPuntoCarga As Long, lnIdServicioActualPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idServicioPaciente", adInteger, adParamInput, 0, lnIdServicioActualPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPorFechasLab(ldFechaIni As Date, ldFechaFin As Date, lnIdPuntoCarga As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPorFechasLab = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPorFechasLab"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPorFechasLab = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FactOrdenServicioPagosSeleccionarPorIdOrden(lidOrden As Long, oConexion As Connection) As ADODB.Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FactOrdenServicioPagosSeleccionarPorIdOrden = Nothing

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactOrdenServicioPagosSeleccionarPorIdOrden"
            Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lidOrden): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
'            Set oRsTmp2.ActiveConnection = Nothing
        End With
        Set FactOrdenServicioPagosSeleccionarPorIdOrden = oRsTmp2
        Set oCommand = Nothing
        Set oParameter = Nothing
 '       Set oRsTmp2 = Nothing
        
        

End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPagosSeleccionarPorIdOrdenPago(lidOrdenPago As Long) As SIGHcomun.DoFactOrdenServPagos
Dim oConexion As New ADODB.Connection
Dim oFactOrdenServicioPagos As New FactOrdenServicioPagos
Dim oDOFactOrdenServicioPagos As New SIGHcomun.DoFactOrdenServPagos
    Set FactOrdenServicioPagosSeleccionarPorIdOrdenPago = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oFactOrdenServicioPagos.Conexion = oConexion
    oDOFactOrdenServicioPagos.idOrdenPago = lidOrdenPago
    If oFactOrdenServicioPagos.SeleccionarPorId(oDOFactOrdenServicioPagos) Then
        Set FactOrdenServicioPagosSeleccionarPorIdOrdenPago = oDOFactOrdenServicioPagos
    End If
    ms_MensajeError = oFactOrdenServicioPagos.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oFactOrdenServicioPagos = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioPagosFiltraPorIdOrdenPago(lidOrdenPago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosFiltraPorIdOrdenPago = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosFiltraPorIdOrdenPago"
        Set oParameter = .CreateParameter("@idOrdenPago", adInteger, adParamInput, 0, lidOrdenPago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosFiltraPorIdOrdenPago = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FacturacionServicioPagosPorCuenta(lnCuenta As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosPorCuenta = Nothing
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
    Else
        Set oConexion = oConexion1
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosPorCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionServicioPagosPorCuenta = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesFiltraIngresosPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesFiltraIngresosPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraIngresosPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesFiltraIngresosPorIdAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesFiltraEgresosPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesFiltraEgresosPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraEgresosPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesFiltraEgresosPorIdAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function AtencionesFiltraTransferenciasPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesFiltraTransferenciasPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraTransferenciasPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesFiltraTransferenciasPorIdAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function AtencionesFiltraDiagnosticosPorIdAtencion(lnIdAtencion As Long, lnIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesFiltraDiagnosticosPorIdAtencion = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesFiltraDiagnosticosPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesFiltraDiagnosticosPorIdAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionServicioDespachoFiltraPorIdOrden(lnIdOrden As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionServicioDespachoFiltraPorIdOrden = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDespachoFiltraPorIdOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionServicioDespachoFiltraPorIdOrden = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga(lnIdOrden As Long, lnIdPuntoCarga) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function ServicioParteDiario(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnCajero As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        
        ms_MensajeError = ""
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ServicioParteDiario"
            Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, lnCajero): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ServicioParteDiario = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FarmaciaParteDiario(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnCajero As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaParteDiario = Nothing
    ms_MensajeError = ""
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaParteDiario"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, lnCajero): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaParteDiario = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function FarmaciaConsolidado(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnCajero As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset


    
    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set FarmaciaConsolidado = Nothing
        ms_MensajeError = ""
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FarmaciaConsolidado"
            Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, lnCajero): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set FarmaciaConsolidado = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function ServicioConsolidado(lnIdCaja As Long, ldFechaIni As Date, ldFechaFin As Date, lnIdTurno As Long, lnIdCajero As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String

    Set ServicioConsolidado = oRecordset
    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
        Set ServicioConsolidado = Nothing
        ms_MensajeError = ""
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ServicioConsolidado"
            Set oParameter = .CreateParameter("@idCaja", adInteger, adParamInput, 0, lnIdCaja): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy hh:mm:ss")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idTurno", adInteger, adParamInput, 0, lnIdTurno): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, lnIdCajero): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set ServicioConsolidado = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
    'Set oRecordset.ActiveConnection = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioBoletasPorIdPacienteIdCuenta(lnIdPaciente As Long, lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioBoletasPorIdPacienteIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioBoletasPorIdPacienteIdCuenta"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioBoletasPorIdPacienteIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FarmaciaBoletasPorIdPacienteIdCuenta(lnIdPaciente As Long, lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaBoletasPorIdPacienteIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaBoletasPorIdPacienteIdCuenta"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaBoletasPorIdPacienteIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesConsultaExternaPorFechas(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesConsultaExternaPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesConsultaExternaPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesConsultaExternaPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioBoletasPorFechas(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioBoletasPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioBoletasPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioBoletasPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FarmaciaConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String


    Set FarmaciaConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Set oParameter = Nothing



Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServiciosConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServiciosConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosConsumosHospitalizadosFiltraPorFechasEgresoAdministrativo = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesFinanciamientosPorIdCuenta(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesFinanciamientosPorIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinanciamientosPorIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioFinanciamientosPorIdCuenta(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioFinanciamientosPorIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioFinanciamientosPorIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServiciosConsumoConveniosPorFechas(ldFechaIni As Date, ldFechaFin As Date, lnOrden As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServiciosConsumoConveniosPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioConsumoConveniosPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosConsumoConveniosPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServiciosConsumoConvenioConsultaExternaPorFechas(ldFechaIni As Date, ldFechaFin As Date, lnIdPaciente As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServiciosConsumoConvenioConsultaExternaPorFechas = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioConsumoConvenioConsultaExternaPorFechas"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosConsumoConvenioConsultaExternaPorFechas = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioFinanciamientosPorTipoFinanciamientoFechaLiquidacion(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioFinanciamientosPorTipoFinanciamientoFechaLiquidacion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorTipoFinanciamientoFechaLiquidacion"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioFinanciamientosPorTipoFinanciamientoFechaLiquidacion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FarmaciaFinanciamientosPorTipoFinanciamientoFechaLiquidacion(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaLiquidacion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorTipoFinanciamientoFechaLiquidacion"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaLiquidacion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FactOrdenServicioPagosPorIdAtencion(lnIdAtencion As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FactOrdenServicioPagosPorIdAtencion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosPorIdAtencion"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FactOrdenServicioPagosPorIdAtencion = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPagosPorIdAtencionIdProducto(lnIdAtencion As Long, lnIdProducto As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPagosPorIdAtencionIdProducto = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosPorIdAtencionIdProducto"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPagosPorIdAtencionIdProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPagosPorIdCuentaAtencion(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New Connection
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set FactOrdenServicioPagosPorIdCuentaAtencion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosPorIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPagosPorIdCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
   oConexion.Close
   Set oConexion = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-12/04/2016
Function DevuelveDeudaPacienteDeAntencionesAnteriores(lnIdPaciente As Long, oConexion As Connection, lnCuentaActual As Long) As String
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oRsCuentas As New Recordset
    Dim lnTotalPorPagar As Double, lnTotal As Double
    Dim lcDeuda As String, lcFallecido As String

    lcFallecido = DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(lnIdPaciente, oConexion)

    lcDeuda = ""
    On Error Resume Next
    Set oRsCuentas = mo_ReglasFarmacia.FacturacionCuentasAtencionSeleccionarPorIdPaciente(lnIdPaciente, oConexion)
    If oRsCuentas.RecordCount > 0 Then
       oRsCuentas.MoveFirst
       Do While Not oRsCuentas.EOF
          If oRsCuentas.Fields!IdEstado <> 4 And oRsCuentas.Fields!IdEstado <> 9 And _
                 oRsCuentas.Fields!IdCuentaAtencion <> lnCuentaActual And oRsCuentas.Fields!GeneraPago = 1 Then
                lnTotalPorPagar = oRsCuentas.Fields!totalPorPagar - (oRsCuentas.Fields!totalPagado + _
                                                                     oRsCuentas.Fields!totalExonerado)
                If lnTotalPorPagar <> 0 Then
                    lcDeuda = lcDeuda + Trim(Str(oRsCuentas.Fields!IdCuentaAtencion)) & " <" & Format(lnTotalPorPagar, "####,###,##0.00") & "> "
                End If
          End If
          oRsCuentas.MoveNext
       Loop
    End If
    oRsCuentas.Close
    If lcDeuda <> "" Then
       lcDeuda = "Ctas: " & lcDeuda
    End If
    ms_Texto = ""
    If lcFallecido <> "" Then
       lcDeuda = lcFallecido & " " & lcDeuda
       MsgBox lcDeuda, vbInformation, ""
    End If
    
    Set oRsCuentas = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasFacturacion = Nothing
    DevuelveDeudaPacienteDeAntencionesAnteriores = lcDeuda
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioSeleccionarPorIdComprobantePago = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@lnIdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPorIdComprobantePago = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 03-10-13 Se cambio a Conexion Base de datos
Function BuscaServicioActualDelPaciente(ml_IdServicioPaciente As Long) As String
    Dim oBuscaNombreServicioPaciente As New SIGHNegocios.ReglasServiciosHosp
    Dim DOServicio As New DOServicio
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set DOServicio = oBuscaNombreServicioPaciente.ServiciosSeleccionarPorId(ml_IdServicioPaciente, oConexion)
    BuscaServicioActualDelPaciente = DOServicio.nombre
    oConexion.Close
    Set oConexion = Nothing
    Set DOServicio = Nothing
    Set oBuscaNombreServicioPaciente = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenesBienesSeleccionarPorIdCuenta(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenesBienesSeleccionarPorIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesSeleccionarPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenesBienesSeleccionarPorIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioPagosSeleccionarPorId(lnIdOrdenPago As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosSeleccionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosSeleccionarPorId"
        Set oParameter = .CreateParameter("@idOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosSeleccionarPorId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino 03-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioSeleccionarPorIdOrdenPago(lnIdOrdenPago As Long, lnIdPuntoCarga As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioSeleccionarPorIdOrdenPago = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPorIdOrdenPago"
        Set oParameter = .CreateParameter("@IdOrdenPago", adInteger, adParamInput, 0, lnIdOrdenPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPorIdOrdenPago = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioSeleccionarPorIdCuenta(lidCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioSeleccionarPorIdCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarPorIdCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lidCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarPorIdCuenta = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga(lnIdOrden As Long, lnIdPuntoCarga, lnIdMovimiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, lnIdMovimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set LabFacturacionServicioDespachoFiltraPorIdOrdenIdPuntoCarga = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 03-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPorFechasYTipoFinanciamiento(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPorFechasYTipoFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPorFechasYTipoFinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPorFechasYTipoFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function farmMovimientoVentasSeleccionarPorFechasYTipoFinanciamiento(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmMovimientoVentasSeleccionarPorFechasYTipoFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoVentasSeleccionarPorFechasYTipoFinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmMovimientoVentasSeleccionarPorFechasYTipoFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarPorFechaEgresoMedicoYtipofinanciamiento(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorFechaEgresoMedicoYtipofinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaEgresoMedicoYtipofinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFechaEgresoMedicoYtipofinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarPorTipoServicio(lnIdTipoServicio As Long, ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorTipoServicio = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorTipoServicio"
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorTipoServicio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FarmaciaFinanciamientosPorFuenteFinanciamientoFechaLiquidacion(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaFinanciamientosPorFuenteFinanciamientoFechaLiquidacion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorFuenteFinanciamientoFechaLiquidacion"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaFinanciamientosPorFuenteFinanciamientoFechaLiquidacion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioFinanciamientosPorFuenteFinanciamientoFechaLiquidacion(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioFinanciamientosPorFuenteFinanciamientoFechaLiquidacion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorTipoFinanciamientoFechaLiquidacion"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioFinanciamientosPorFuenteFinanciamientoFechaLiquidacion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarPorFechaEgresoMedicoYfuenteFinanciamiento(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorFechaEgresoMedicoYfuenteFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaEgresoMedicoYfuentefinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFechaEgresoMedicoYfuenteFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarPorFechaEgresoMedico(ldFechaIni As Date, ldFechaFin As Date, lbSoloConsultaExterna As Boolean, oConexion As Connection, lnIdFuenteFinanciamiento As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String
    Dim LnSoloConsultaExterna As Integer
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        Set AtencionesSeleccionarPorFechaEgresoMedico = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "AtencionesSeleccionarPorFechaEgresoMedico"
            Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@SoloConsultaExterna", adInteger, adParamInput, 0, IIf(lbSoloConsultaExterna = True, 1, 0)): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       If lnIdFuenteFinanciamiento > 0 Then
          oRecordset.Filter = "idFuenteFinanciamiento=" & lnIdFuenteFinanciamiento
       End If
       Set AtencionesSeleccionarPorFechaEgresoMedico = oRecordset
       
       Set oRecordset = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesDevolucionesSeleccionarPorId(lcMovNumero As String, lcMovTipo As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesDevolucionesSeleccionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesDevolucionesSeleccionarPorId"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero):        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo):        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesDevolucionesSeleccionarPorId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesPagosSeleccionarPorCuentaTodos(lnCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesPagosSeleccionarPorCuentaTodos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorCuentaTodos"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagosSeleccionarPorCuentaTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesFinanciamientosSeleccionaXProdFinanciam(lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long, lnIdtipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesFinanciamientosSeleccionaXProdFinanciam = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosSeleccionaXProdFinanciam"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinanciamientosSeleccionaXProdFinanciam = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FacturacionBienesPagosSeleccionarPorMovNumeroProducto(lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oTabla As New DoFacturacionBienesPagos
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = oConexion
     .CommandText = "FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   Set FacturacionBienesPagosSeleccionarPorMovNumeroProducto = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionBienesFinanciamientosSeleccionarPorId(lcMovNumero As String, lcMovTipo As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionBienesFinanciamientosSeleccionarPorId = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesFinanciamientosSeleccionarPorId"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesFinanciamientosSeleccionarPorId = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioPagosPorCuentaTodos(lnCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosPorCuentaTodos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosPorCuentaTodos"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosPorCuentaTodos = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnNroCuenta As Long, Optional lbAunNoTieneReembolso As Boolean) As Double
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
Dim lnTotal As Double
    '
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorNroCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   'Consumos
   lnTotal = 0
   If lbAunNoTieneReembolso = False Then
      oRecordset.Filter = "idEstadoFacturacionFinanciamiento<>9 and idEstadoFacturacionFinanciamiento<>12 and idEstadoFacturacionFinanciamiento<>4"
   Else
      oRecordset.Filter = "idEstadoFacturacionFinanciamiento<>9"
   End If
   If oRecordset.RecordCount > 0 Then
      oRecordset.MoveFirst
      Do While Not oRecordset.EOF
          lnTotal = lnTotal + oRecordset.Fields!TotalFinanciado
          oRecordset.MoveNext
      Loop
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   RetornaConsumoPacienteServiciosConSeguroPorNroCuenta = lnTotal
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPreventasServicio(ldFechaIni As Date, ldFechaFin As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPreventasServicio = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPreventasServicio"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPreventasServicio = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmMovimientoVentasDetalleSeleccionarPorMovNumero(lcMovNumero As String, lcMovTipo As String, oConexion As Connection) As ADODB.Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim ms_MensajeError As String
  
  ms_MensajeError = ""
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = oConexion
    .CommandTimeout = 150
    .CommandText = "FarmMovimientoVentasDetalleSeleccionarPorMovNumero"
    Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero):    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo):    .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  Set FarmMovimientoVentasDetalleSeleccionarPorMovNumero = oRecordset
  Set oCommand = Nothing
  Set oRecordset = Nothing
  Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarPorFechaIngresoYfuenteFinanciamiento(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarPorFechaIngresoYfuenteFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaIngresoYfuenteFinanciamiento"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFechaIngresoYfuenteFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioFinanciamientosPorTipoFinanciamientoFechaAltaMedica(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioFinanciamientosPorTipoFinanciamientoFechaAltaMedica = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorTipoFinanciamientoFechaAltaMedica"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioFinanciamientosPorTipoFinanciamientoFechaAltaMedica = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FarmaciaFinanciamientosPorTipoFinanciamientoFechaEgresoMedico(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long, lnTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaEgresoMedico = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorTipoFinanciamientoFechaEgresoMedico"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoServicio", adInteger, adParamInput, 0, lnTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaEgresoMedico = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ServicioFinanciamientosPorTipoFinanciamientoFechaIngreso(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set ServicioFinanciamientosPorTipoFinanciamientoFechaIngreso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServicioFinanciamientosPorTipoFinanciamientoFechaIngreso"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServicioFinanciamientosPorTipoFinanciamientoFechaIngreso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FarmaciaFinanciamientosPorTipoFinanciamientoFechaIngreso(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaIngreso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaFinanciamientosPorTipoFinanciamientoFechaIngreso"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmaciaFinanciamientosPorTipoFinanciamientoFechaIngreso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function AreaTramitaSegurosDevuelveTodosSegunFiltro(lcFiltro As String) As ADODB.Recordset
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set AreaTramitaSegurosDevuelveTodosSegunFiltro = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AreaTramitaSegurosDevuelveTodosSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AreaTramitaSegurosDevuelveTodosSegunFiltro = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FuentesFinanciamientoDevuelveTodosSegunFiltro(lcFiltro As String) As ADODB.Recordset
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FuentesFinanciamientoDevuelveTodosSegunFiltro = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FuentesFinanciamientoDevuelveTodosSegunFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoDevuelveTodosSegunFiltro = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
 
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long, lnIdTipoServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarHospEmergPorFechaAltaMedicaYplan = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesSeleccionarCEPorFechaIngresoYplan(ldFechaIni As Date, ldFechaFin As Date, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesSeleccionarCEPorFechaIngresoYplan = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarCEPorFechaIngresoYplan"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarCEPorFechaIngresoYplan = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function RetornaConsumoFarmaciaServiciosPorNroCuenta(lnNroCuenta As Long, oConexion As Connection, Optional lbAunNoTieneReembolso As Boolean) As Double
   Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
   Dim lnPagosAdelantados As Double, lnConsumoServicio As Double, lnConsumoFarmacia As Double
   'Consumo Servicios-Seguros
   lnConsumoServicio = RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnNroCuenta, lbAunNoTieneReembolso)
   'Consumo Farmacia-Seguros
   lnConsumoFarmacia = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(lnNroCuenta)
   'Pagos a cuenta
   lnPagosAdelantados = RetornaImporteDePagosAdelantadosPorNroCuenta(lnNroCuenta, oConexion)
   Set mo_ReglasFarmacia = Nothing
   RetornaConsumoFarmaciaServiciosPorNroCuenta = lnConsumoServicio + lnConsumoFarmacia - lnPagosAdelantados
   
End Function

Function RetornaTotalPagosPendientesPorNroCuentaDEBB(lnNroCuenta As Long, oConexion As Connection) As Double
   Dim lnPagosAdelantados As Double, lnConsumoServicio As Double, lnConsumoFarmacia As Double
   Dim lnPagosXdevoluciones As Double
   'Consumo Servicios-Particular
   lnConsumoServicio = RetornaTotalPagosServiciosPendientesPorNroCuentaDEBB(lnNroCuenta, oConexion)
   'Consumo Farmacia-Particular
   lnConsumoFarmacia = RetornaTotalPagosFarmaciaPendientesPorNroCuentaDEBB(lnNroCuenta, oConexion)
   'Pagos a cuenta
   If lnConsumoServicio > 0 Or lnConsumoFarmacia > 0 Then
      lnPagosAdelantados = RetornaImporteDePagosAdelantadosPorNroCuenta(lnNroCuenta, oConexion)
      lnPagosXdevoluciones = RetornaImporteDePagosXdevolucionesPorNroCuenta(lnNroCuenta, oConexion)
   Else
      lnPagosAdelantados = 0
      lnPagosXdevoluciones = 0
   End If
   RetornaTotalPagosPendientesPorNroCuentaDEBB = lnConsumoServicio + lnConsumoFarmacia - lnPagosAdelantados + lnPagosXdevoluciones
End Function

'debb-25/02/2011
Function RetornaTotalPagosServiciosPendientesPorNroCuentaDEBB(lnNroCuenta As Long, Optional oConexion1 As Connection) As Double
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Dim lnIdOrdenPago As Long
    Dim oConexion As New Connection
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oRsTmp = FacturacionServicioPagosPorCuenta(lnNroCuenta, oConexion)
    Else
        Set oRsTmp = FacturacionServicioPagosPorCuenta(lnNroCuenta, oConexion1)
    End If
    
    lnTotal = 0
    oRsTmp.Filter = "EstadoFacOrdenServicio<>9"
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnIdOrdenPago = oRsTmp.Fields!idOrdenPago
            If oRsTmp.Fields!ImporteExonerado > 0 Then
               lnTotal = lnTotal - oRsTmp.Fields!ImporteExonerado
            End If
            Do While Not oRsTmp.EOF And lnIdOrdenPago = oRsTmp.Fields!idOrdenPago
               lnTotal = lnTotal + oRsTmp.Fields!totalPorPagar
               oRsTmp.MoveNext
               If oRsTmp.EOF Then
                  Exit Do
               End If
            Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
    RetornaTotalPagosServiciosPendientesPorNroCuentaDEBB = lnTotal
End Function

Function RetornaTotalPagosFarmaciaPendientesPorNroCuentaDEBB(lnNroCuenta As Long, Optional oConexion1 As Connection) As Double
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Dim lnIdOrden As Long
    Dim oConexion As New Connection
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set oRsTmp = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorCuenta(lnNroCuenta, oConexion)
        
    Else
        Set oRsTmp = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorCuenta(lnNroCuenta, oConexion1)
    End If
    oRsTmp.Filter = "movNumero<>null"
    lnTotal = 0
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnIdOrden = oRsTmp.Fields!IdOrden
            If oRsTmp.Fields!ImporteExonerado > 0 Then
               lnTotal = lnTotal - oRsTmp.Fields!ImporteExonerado
            End If
            Do While Not oRsTmp.EOF And lnIdOrden = oRsTmp.Fields!IdOrden
               lnTotal = lnTotal + oRsTmp.Fields!totalPorPagar
               oRsTmp.MoveNext
               If oRsTmp.EOF Then
                  Exit Do
               End If
            Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oConexion = Nothing
    RetornaTotalPagosFarmaciaPendientesPorNroCuentaDEBB = lnTotal
End Function

Function RetornaTotalExoneracionesFarmaciaPendientesPorNroCuenta(lnNroCuenta As Long) As Double
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Dim lnIdOrden As Long
    Set oRsTmp = mo_ReglasFarmacia.FacturacionBienesPagosSeleccionarPorCuenta(lnNroCuenta)
    oRsTmp.Filter = "movNumero<>null"
    lnTotal = 0
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnIdOrden = oRsTmp.Fields!IdOrden
            If oRsTmp.Fields!ImporteExonerado > 0 Then
               lnTotal = lnTotal + oRsTmp.Fields!ImporteExonerado
            End If
            Do While Not oRsTmp.EOF And lnIdOrden = oRsTmp.Fields!IdOrden
               oRsTmp.MoveNext
               If oRsTmp.EOF Then
                  Exit Do
               End If
            Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set mo_ReglasFarmacia = Nothing
    RetornaTotalExoneracionesFarmaciaPendientesPorNroCuenta = lnTotal
End Function

Function RetornaExoneracionesFarmaciaPendientesPorProducto(lcMovNumero As String, lnIdProducto As Long) As Double
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Set oRsTmp = mo_ReglasFacturacion.FacturacionBienesFinanciamientosSeleccionaXProdFinanciam(lcMovNumero, "S", lnIdProducto, 9)
    oRsTmp.Filter = "movNumero<>null"
    lnTotal = 0
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnTotal = lnTotal + oRsTmp.Fields!TotalFinanciado
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set mo_ReglasFacturacion = Nothing
    RetornaExoneracionesFarmaciaPendientesPorProducto = lnTotal
End Function


Function RetornaImporteDePagosAdelantadosPorNroCuenta(lnNroCuenta As Long, oConexion As Connection) As Double
   Dim lnPagosAdelantados  As Double
   Dim oRsTmp1 As New Recordset
   Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
   Set oRsTmp1 = mo_AdminCaja.CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta(lnNroCuenta, Val(lcBuscaParametro.SeleccionaFilaParametro(245)), oConexion)
   lnPagosAdelantados = 0
   If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lnPagosAdelantados = lnPagosAdelantados + oRsTmp1.Fields!total
            oRsTmp1.MoveNext
         Loop
   End If
   Set mo_AdminCaja = Nothing
   Set oRsTmp1 = Nothing
   RetornaImporteDePagosAdelantadosPorNroCuenta = lnPagosAdelantados
End Function

Function RetornaImporteDePagosRealizadosPorNroCuenta(lnNroCuenta As Long) As Double
   Dim lnPagos  As Double
   Dim oRsTmp1 As New Recordset
   Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
   Set oRsTmp1 = mo_AdminCaja.CajaComprobantesSeleccionarPorCuentaAtencion(lnNroCuenta)
   lnPagos = 0
   If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lnPagos = lnPagos + oRsTmp1.Fields!total
            oRsTmp1.MoveNext
         Loop
   End If
   Set mo_AdminCaja = Nothing
   Set oRsTmp1 = Nothing
   RetornaImporteDePagosRealizadosPorNroCuenta = lnPagos
End Function



Function FactReembolsosSeleccionarPorId(idFactReembolso As Long) As DoFactreembolsos
Dim oConexion As New ADODB.Connection
Dim oFactReembolsos As New FactReembolsos
Dim DoFactreembolsos As New DoFactreembolsos
    Set FactReembolsosSeleccionarPorId = Nothing
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFactReembolsos.Conexion = oConexion
    DoFactreembolsos.idFactReembolso = idFactReembolso
    If oFactReembolsos.SeleccionarPorId(DoFactreembolsos) Then
        Set FactReembolsosSeleccionarPorId = DoFactreembolsos
    End If
    oConexion.Close
    ms_MensajeError = oFactReembolsos.MensajeError
    Set oConexion = Nothing
    Set oFactReembolsos = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionReembolsosSeleccionarHospEmergPorIdFactReembolso(lnIdFactReembolso As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionReembolsosSeleccionarHospEmergPorIdFactReembolso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionReembolsosSeleccionarHospEmergPorIdFactReembolso"
        Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, lnIdFactReembolso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionReembolsosSeleccionarHospEmergPorIdFactReembolso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionReembolsosSeleccionarCEPorIdFactReembolso(lnIdFactReembolso As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionReembolsosSeleccionarCEPorIdFactReembolso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionReembolsosSeleccionarCEPorIdFactReembolso"
        Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, lnIdFactReembolso): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionReembolsosSeleccionarCEPorIdFactReembolso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactReembolsosSelecionarPorAreaTramitaSeguros(lnIdAreaTramitaSeguro As Long, lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactReembolsosSelecionarPorAreaTramitaSeguros = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactReembolsosSelecionarPorAreaTramitaSeguros"
        Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactReembolsosSelecionarPorAreaTramitaSeguros = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function DevuelveDxAltaMedica(lnIdAtencion As Long, lnIdTipoServicio As Long, Optional oConexion1 As Connection) As doDiagnostico
            Dim mo_AdminAdmision1 As New SIGHNegocios.ReglasAdmision
            Dim oRsTmp As New Recordset
            Dim oDODiagnostico As New doDiagnostico
            Dim oConexion As New Connection
            If oConexion1 Is Nothing Then
                oConexion.CommandTimeout = 300
                oConexion.CursorLocation = adUseClient
                oConexion.Open sighentidades.CadenaConexion
            End If
            If oConexion1 Is Nothing Then
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, IIf(lnIdTipoServicio = 1, 1, 3), oConexion)
            Else
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, IIf(lnIdTipoServicio = 1, 1, 3), oConexion1)
            End If
            oDODiagnostico.IdDiagnostico = 0
            oDODiagnostico.Descripcion = ""
            oDODiagnostico.CodigoCIE2004 = ""
            If lnIdTipoServicio = 1 Then
               oRsTmp.Filter = "IdTipoDiagnostico=102"    'CE, vivos
            Else
               oRsTmp.Filter = "IdTipoDiagnostico=301"    'Hosp/emerg, vivos
            End If
            If oRsTmp.RecordCount > 0 Then
               oDODiagnostico.CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
               oDODiagnostico.Descripcion = oRsTmp.Fields!Descripcion
               oDODiagnostico.IdDiagnostico = oRsTmp.Fields!IdDiagnostico
            ElseIf lnIdTipoServicio <> 1 Then
               If oConexion1 Is Nothing Then
                  Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 4, oConexion)
               Else
                  Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 4, oConexion1)
               End If
               oRsTmp.Filter = "IdTipoDiagnostico=303"    'Hosp/emerg, fallecidos
               If oRsTmp.RecordCount > 0 Then
                    oDODiagnostico.CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsTmp.Fields!Descripcion
                    oDODiagnostico.IdDiagnostico = oRsTmp.Fields!IdDiagnostico
               End If
            End If
            Set oRsTmp = Nothing
            Set DevuelveDxAltaMedica = oDODiagnostico
            Set oConexion = Nothing
            Set mo_AdminAdmision1 = Nothing
End Function
Function DevuelveDxAltaMedicaTodosDx(lnIdAtencion As Long, lnIdTipoServicio As Long) As String
            Dim mo_AdminAdmision1 As New SIGHNegocios.ReglasAdmision
            Dim oRsTmp As New Recordset
            Dim oDODiagnostico As New doDiagnostico
            Dim oConexion As New Connection
            oConexion.Open sighentidades.CadenaConexion
            oConexion.CursorLocation = adUseClient
            DevuelveDxAltaMedicaTodosDx = ""
            If lnIdTipoServicio = 1 Then
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 1, oConexion)
            Else
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 3, oConexion)
            End If
            If lnIdTipoServicio = 1 Then
               'oRsTmp.Filter = "IdTipoDiagnostico=102"    'CE, vivos
            Else
               oRsTmp.Filter = "IdTipoDiagnostico=301"    'Hosp/emerg, vivos
            End If
            If oRsTmp.RecordCount > 0 Then
               oRsTmp.MoveFirst
               Do While Not oRsTmp.EOF
                  DevuelveDxAltaMedicaTodosDx = DevuelveDxAltaMedicaTodosDx & "(" & Trim(oRsTmp.Fields!CodigoCIE2004) & "-" & Left(Trim(oRsTmp.Fields!DescripcionTipoDx), 2) & "-" & Trim(oRsTmp.Fields!Descripcion) & ") "
                  oRsTmp.MoveNext
               Loop
            ElseIf lnIdTipoServicio <> 1 Then
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 4, oConexion)
               oRsTmp.Filter = "IdTipoDiagnostico=303"    'Hosp/emerg, fallecidos
               If oRsTmp.RecordCount > 0 Then
                    oRsTmp.MoveFirst
                    Do While Not oRsTmp.EOF
                       DevuelveDxAltaMedicaTodosDx = DevuelveDxAltaMedicaTodosDx & "(" & Trim(oRsTmp.Fields!CodigoCIE2004) & "-" & Trim(oRsTmp.Fields!Descripcion) & ") "
                       oRsTmp.MoveNext
                    Loop
               End If
            End If
            Set oRsTmp = Nothing
            oConexion.Close
            Set oConexion = Nothing
            Set mo_AdminAdmision1 = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactReembolsosSelecionarPorAnioYAreatsYPlan(lnAnio As Long, lnIdAreaTramitaSeguro As Long, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactReembolsosSelecionarPorAnioYAreatsYPlan = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactReembolsosSelecionarPorAnioYAreatsYPlan"
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactReembolsosSelecionarPorAnioYAreatsYPlan = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionReembolsosSeleccionarPorCuentaPlanAreats(lnIdCuentaAtencion As Long, lnIdAreaTramitaSeguro As Long, lnIdFuenteFinanciamiento As Long, lnIdTipoConsumo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionReembolsosSeleccionarPorCuentaPlanAreats = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionReembolsosSeleccionarPorCuentaPlanAreats"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoConsumo", adInteger, adParamInput, 0, lnIdTipoConsumo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionReembolsosSeleccionarPorCuentaPlanAreats = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactReembolsosSeleccionarPorPlanAreats(lnIdAreaTramitaSeguro As Long, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactReembolsosSeleccionarPorPlanAreats = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactReembolsosSeleccionarPorPlanAreats"
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactReembolsosSeleccionarPorPlanAreats = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento(lnIdProducto As Long, lnIdtipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactCatalogoServiciosHospSeleccionarPorIdYtipoFinanciamiento = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function LabMovimientoLaboratorioSeleccionarPorIdOrden(lnIdOrden As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set LabMovimientoLaboratorioSeleccionarPorIdOrden = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "LabMovimientoLaboratorioSeleccionarPorIdOrden"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set LabMovimientoLaboratorioSeleccionarPorIdOrden = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function CuentasAtencionAltaConDeudaYGarante(oDOCuentaAtencion As DOCuentaAtencion, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcPlanActual As String) As Boolean
  Dim oConexion As New ADODB.Connection
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oCuentasAtencion As New CuentasAtencion
  Dim oRsTmp As New ADODB.Recordset
  Dim sSql As String
  On Error GoTo ErrCierre
  CuentasAtencionAltaConDeudaYGarante = False
  oConexion.Open sighentidades.CadenaConexion
  Set oCuentasAtencion.Conexion = oConexion
  oConexion.BeginTrans
  If oCuentasAtencion.AltaConDeudaYGarante(oDOCuentaAtencion) = True Then
  
  Else
    ms_MensajeError = oCuentasAtencion.MensajeError
    Exit Function
  End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasActualizarXIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oDOCuentaAtencion.idPaciente): .Parameters.Append oParameter
        .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
  
  Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCuentaAtencion.IdUsuarioAuditoria, "A", oDOCuentaAtencion.IdCuentaAtencion, "Pasó a Pagada", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcPlanActual)             'ListBarItems.idListItem
  oConexion.CommitTrans
  oConexion.Close
  Set oConexion = Nothing
  Set oCommand = Nothing
  Set oCuentasAtencion = Nothing
  Set oRsTmp = Nothing
  CuentasAtencionAltaConDeudaYGarante = True
  Exit Function
  
ErrCierre:
  oConexion.RollbackTrans
  oConexion.Close
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function atencionesSeleccionarPorFuenteFinanciamientoFAltaMedica(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long, lnidEstadoCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set atencionesSeleccionarPorFuenteFinanciamientoFAltaMedica = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSeleccionarPorFuenteFinanciamientoFAltaMedica"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEstadoCuenta", adInteger, adParamInput, 0, lnidEstadoCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set atencionesSeleccionarPorFuenteFinanciamientoFAltaMedica = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function atencionesSeleccionarPorFuenteFinanciamientoFingreso(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long, lnidEstadoCuenta As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set atencionesSeleccionarPorFuenteFinanciamientoFingreso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "atencionesSeleccionarPorFuenteFinanciamientoFingreso"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEstadoCuenta", adInteger, adParamInput, 0, lnidEstadoCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set atencionesSeleccionarPorFuenteFinanciamientoFingreso = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function TiposConsumosDevuelveTodos(lcFiltro As String) As ADODB.Recordset
Dim oRecordset As New ADODB.Recordset
Dim lcSql As String

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set TiposConsumosDevuelveTodos = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposConsumosDevuelveTodos"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set TiposConsumosDevuelveTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function ReembolsosAgregar(oDoFactReembolsos As DoFactreembolsos, oDoFactReembolsosDcto As DoFactReembolsosDcto, oRsListaCuentas As Recordset, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactReembolsos As New FactReembolsos
Dim oDoFacturacionReembolsos As New DoFacturacionReembolsos
Dim oFacturacionReembolsos As New FacturacionReembolsos
Dim oDoAtencion As New DOAtencion
Dim oAtenciones As New Atenciones
Dim oFactReembolsosDocumentos As New FactReembolsosDocumentos
Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
Dim oFacturacionServicioFin As New FacturacionServicioFin
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim lbContinua As Boolean
Dim oRsTmp As New Recordset
Dim lnReembolsoFarmacia As Double, lnReembolsoServicio As Double
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oFactReembolsos.Conexion = oConexion
    Set oFacturacionReembolsos.Conexion = oConexion
    Set oAtenciones.Conexion = oConexion
    Set oFactReembolsosDocumentos.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    If Not oFactReembolsos.Insertar(oDoFactReembolsos) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsos.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFactReembolsos.IdUsuarioAuditoria, "A", oDoFactReembolsos.idFactReembolso, "FactReembolsos", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)     'ListBarItems.idListItem
    '
    oDoFactReembolsosDcto.idFactReembolso = oDoFactReembolsos.idFactReembolso
    oDoFactReembolsosDcto.IdUsuarioAuditoria = oDoFactReembolsos.IdUsuarioAuditoria
    If Not oFactReembolsosDocumentos.Insertar(oDoFactReembolsosDcto) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsosDocumentos.MensajeError: GoTo Terminar
    End If
    '
    If oRsListaCuentas.RecordCount > 0 Then
        oDoFacturacionReembolsos.idFactReembolso = oDoFactReembolsos.idFactReembolso
        oDoFacturacionReembolsos.IdUsuarioAuditoria = oDoFactReembolsos.IdUsuarioAuditoria
        oRsListaCuentas.MoveFirst
        Do While Not oRsListaCuentas.EOF
           If oRsListaCuentas.Fields!PorReembolsar > 0 Or oRsListaCuentas.Fields!Reembolsado > 0 Then
                oDoFacturacionReembolsos.ConsumoPorReembolsar = oRsListaCuentas.Fields!consumo
                oDoFacturacionReembolsos.IdCuentaAtencion = oRsListaCuentas.Fields!nrocuenta
                oDoFacturacionReembolsos.IdDiagnostico = oRsListaCuentas.Fields!DxId
                oDoFacturacionReembolsos.ReembolsoPagadoFarmacia = oRsListaCuentas.Fields!ReembolsoPagadoFarmacia
                oDoFacturacionReembolsos.ReembolsoPagadoServicio = oRsListaCuentas.Fields!ReembolsoPagadoServicio
                oDoFacturacionReembolsos.ReembolsoPorPagar = oRsListaCuentas.Fields!PorReembolsar
                oDoFacturacionReembolsos.NroReferenciaDestino = IIf(IsNull(oRsListaCuentas.Fields!NroReferenciaDestino), "", oRsListaCuentas.Fields!NroReferenciaDestino)
                oDoFacturacionReembolsos.IdReembolsosAnteriores = IIf(IsNull(oRsListaCuentas.Fields!UltReembolsoPorCuenta), "", oRsListaCuentas.Fields!UltReembolsoPorCuenta)
                If Not oFacturacionReembolsos.Insertar(oDoFacturacionReembolsos) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionReembolsos.MensajeError: GoTo Terminar
                End If
'                If oDoFactReembolsos.idAreaTramitaSeguro = 4 Then
'                   'Referencias/contrareferencias
'                   oDoAtencion.idAtencion = oRsListaCuentas.Fields!idAtencion
'                   If Not oAtenciones.SeleccionarPorId(oDoAtencion) Then
'                       bProcesoOK = False: ms_MensajeError = oAtenciones.MensajeError: GoTo Terminar
'                   End If
'                   oDoAtencion.NroReferenciaDestino = oRsListaCuentas.Fields!NroReferenciaDestino
'                   If Not oAtenciones.Modificar(oDoAtencion) Then
'                       bProcesoOK = False: ms_MensajeError = oAtenciones.MensajeError: GoTo Terminar
'                   End If
'                End If
                'actualiza Cuenta = pagada
               
                If oRsListaCuentas.Fields!Reembolsado > 0 And oRsListaCuentas.Fields!PorReembolsar = 0 And oDoFactReembolsos.GrabaDefinitivamente = True Then
                    lbContinua = False
                    Select Case oDoFactReembolsos.idTipoConsumo
                    Case 1 'farmacia

                       With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacConsultaReembolsosxIdCuentaAtencionFarmacia"
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            Set oRsTmp = .Execute
                            Set oRsTmp.ActiveConnection = Nothing
                       End With
                       Set oParameter = Nothing
                       Set oCommand = Nothing
                                             
                       If oRsTmp.RecordCount > 0 Then
                          lbContinua = True
                       Else
                          If oRsListaCuentas.Fields!consumoS = 0 Then
                             lbContinua = True
                          End If
                       End If
                       oRsTmp.Close
                    Case 2 'servicio

                       With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacConsultaReembolsosxIdCuentaAtencionServicio"
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            Set oRsTmp = .Execute
                            Set oRsTmp.ActiveConnection = Nothing
                       End With
                       Set oCommand = Nothing
                       Set oParameter = Nothing
                       
                       If oRsTmp.RecordCount > 0 Then
                          lbContinua = True
                       Else
                          If oRsListaCuentas.Fields!consumoF = 0 Then
                             lbContinua = True
                          End If
                       End If
                       oRsTmp.Close
                    Case 3 'ambos
                        lbContinua = True
                    End Select
                    If lbContinua = True Then
                       'Actualiza Id de Reembolso total
                       
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 4): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                       
                       If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 4) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                          End If
                       End If
                       If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/15/", 4) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                          End If
                       End If
                    Else
                       'Actualiza Id de Reembolso parcial
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 15): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                        End With
                        Set oParameter = Nothing
                        Set oCommand = Nothing
                       
                       If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 15) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                          End If
                       End If
                       If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/", 15) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                          End If
                       End If
                    End If
                ElseIf oRsListaCuentas.Fields!Reembolsado > 0 And oRsListaCuentas.Fields!PorReembolsar > 0 And oDoFactReembolsos.GrabaDefinitivamente = True Then
                    'Actualiza Id de Reembolso parcial
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 15): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                        End With
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                    
                    If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                       If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 15) Then
                          bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                       End If
                    End If
                    If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                       If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/", 15) Then
                          bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                       End If
                    End If
                End If
                '
           End If
           oRsListaCuentas.MoveNext
        Loop
    End If
Terminar:
    If bProcesoOK Then
        ReembolsosAgregar = True
        oConexion.CommitTrans
    Else
        ReembolsosAgregar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFactReembolsos = Nothing
    Set oDoFacturacionReembolsos = Nothing
    Set oFacturacionReembolsos = Nothing
    Set oDoAtencion = Nothing
    Set oAtenciones = Nothing
    Set oRsTmp = Nothing
    Set oFactReembolsosDocumentos = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function ReembolsosModificar(oDoFactReembolsos As DoFactreembolsos, oDoFactReembolsosDcto As DoFactReembolsosDcto, oRsListaCuentas As Recordset, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactReembolsos As New FactReembolsos
Dim oDoFacturacionReembolsos As New DoFacturacionReembolsos
Dim oFacturacionReembolsos As New FacturacionReembolsos
Dim oDoAtencion As New DOAtencion
Dim oAtenciones As New Atenciones
Dim oFactReembolsosDocumentos As New FactReembolsosDocumentos
Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
Dim oFacturacionServicioFin As New FacturacionServicioFin
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim oRsTmp As New Recordset
Dim oRsTmp1 As New Recordset
Dim lcSql As String
Dim lbContinua As Boolean
Dim lnReembolsoFarmacia As Double, lnReembolsoServicio As Double
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oFacturacionReembolsos.Conexion = oConexion
    If oDoFactReembolsos.idAreaTramitaSeguro = 4 Then
       Set oAtenciones.Conexion = oConexion
    End If
    Set oFactReembolsosDocumentos.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    Set oFactReembolsos.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    If Not oFactReembolsos.modificar(oDoFactReembolsos) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsos.MensajeError: GoTo Terminar
    End If
    '
    oDoFactReembolsosDcto.idFactReembolso = oDoFactReembolsos.idFactReembolso
    oDoFactReembolsosDcto.IdUsuarioAuditoria = oDoFactReembolsos.IdUsuarioAuditoria
    If Not oFactReembolsosDocumentos.modificar(oDoFactReembolsosDcto) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsosDocumentos.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFactReembolsos.IdUsuarioAuditoria, "M", oDoFactReembolsos.idFactReembolso, "FactReembolsos", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria)     'ListBarItems.idListItem
    'Abre cuentas estado=pagada

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultarFacturacionReembolsosXidFactReembolso"
        Set oParameter = .CreateParameter("@IdFactReembolso", adInteger, adParamInput, 0, oDoFactReembolsos.idFactReembolso): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
    
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            Set oCommand = Nothing
            Set oParameter = Nothing
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 11): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!IdCuentaAtencion): .Parameters.Append oParameter
                .Execute
             End With
             Set oParameter = Nothing
             Set oCommand = Nothing
                        
            If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
               If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsTmp.Fields!IdCuentaAtencion, 1) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
               End If
            End If
            If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
               If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsTmp.Fields!IdCuentaAtencion, "/4/15/", 1) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
               End If
            End If
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    '
    oDoFacturacionReembolsos.idFactReembolso = oDoFactReembolsos.idFactReembolso
    oDoFacturacionReembolsos.IdUsuarioAuditoria = oDoFactReembolsos.IdUsuarioAuditoria
    If Not oFacturacionReembolsos.Eliminar(oDoFacturacionReembolsos) Then
        bProcesoOK = False: ms_MensajeError = oFacturacionReembolsos.MensajeError: GoTo Terminar
    End If
    If oRsListaCuentas.RecordCount > 0 Then
        oRsListaCuentas.MoveFirst
        Do While Not oRsListaCuentas.EOF
           If oRsListaCuentas.Fields!PorReembolsar > 0 Or oRsListaCuentas.Fields!Reembolsado > 0 Then
                oDoFacturacionReembolsos.ConsumoPorReembolsar = oRsListaCuentas.Fields!consumo
                oDoFacturacionReembolsos.IdCuentaAtencion = oRsListaCuentas.Fields!nrocuenta
                oDoFacturacionReembolsos.IdDiagnostico = oRsListaCuentas.Fields!DxId
                oDoFacturacionReembolsos.ReembolsoPagadoFarmacia = oRsListaCuentas.Fields!ReembolsoPagadoFarmacia
                oDoFacturacionReembolsos.ReembolsoPagadoServicio = oRsListaCuentas.Fields!ReembolsoPagadoServicio
                oDoFacturacionReembolsos.ReembolsoPorPagar = oRsListaCuentas.Fields!PorReembolsar
                oDoFacturacionReembolsos.NroReferenciaDestino = IIf(IsNull(oRsListaCuentas.Fields!NroReferenciaDestino), "", oRsListaCuentas.Fields!NroReferenciaDestino)
                oDoFacturacionReembolsos.IdReembolsosAnteriores = IIf(IsNull(oRsListaCuentas.Fields!UltReembolsoPorCuenta), "", oRsListaCuentas.Fields!UltReembolsoPorCuenta)
                If Not oFacturacionReembolsos.Insertar(oDoFacturacionReembolsos) Then
                    bProcesoOK = False: ms_MensajeError = oFacturacionReembolsos.MensajeError: GoTo Terminar
                End If
'                If oDoFactReembolsos.idAreaTramitaSeguro = 4 Then
'                   'Referencias/contrareferencias
'                   oDoAtencion.idAtencion = oRsListaCuentas.Fields!idAtencion
'                   If Not oAtenciones.SeleccionarPorId(oDoAtencion) Then
'                       bProcesoOK = False: ms_MensajeError = oAtenciones.MensajeError: GoTo Terminar
'                   End If
'                   oDoAtencion.NroReferenciaDestino = oRsListaCuentas.Fields!NroReferenciaDestino
'                   If Not oAtenciones.Modificar(oDoAtencion) Then
'                       bProcesoOK = False: ms_MensajeError = oAtenciones.MensajeError: GoTo Terminar
'                   End If
'                End If
                'actualiza Cuenta = pagada
                If oRsListaCuentas.Fields!Reembolsado > 0 And oRsListaCuentas.Fields!PorReembolsar = 0 And oDoFactReembolsos.GrabaDefinitivamente = True Then
                    lbContinua = False
                    Select Case oDoFactReembolsos.idTipoConsumo
                    Case 1 'farmacia
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        With oCommand
                             .CommandType = adCmdStoredProc
                             Set .ActiveConnection = oConexion
                             .CommandTimeout = 150
                             .CommandText = "FacConsultarFacturacionReembolsosXidFactReembolsoFarmacia"
                             Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                             Set oRsTmp = .Execute
                             Set oRsTmp.ActiveConnection = Nothing
                        End With
   
                       If oRsTmp.RecordCount > 0 Then
                          lbContinua = True
                       Else
                          If oRsListaCuentas.Fields!consumoS = 0 Then
                             lbContinua = True
                          End If
                       End If
                       oRsTmp.Close
                    Case 2 'servicio

                       Set oCommand = Nothing
                       Set oParameter = Nothing
                       With oCommand
                             .CommandType = adCmdStoredProc
                             Set .ActiveConnection = oConexion
                             .CommandTimeout = 150
                             .CommandText = "FacConsultarFacturacionReembolsosXidFactReembolsoServicio"
                             Set oParameter = .CreateParameter("@IdFactReembolso", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                             Set oRsTmp = .Execute
                             Set oRsTmp.ActiveConnection = Nothing
                        End With
                       
                       If oRsTmp.RecordCount > 0 Then
                          lbContinua = True
                       Else
                          If oRsListaCuentas.Fields!consumoF = 0 Then
                             lbContinua = True
                          End If
                       End If
                       oRsTmp.Close
                    Case 3 'ambos
                        lbContinua = True
                    End Select
                    If lbContinua = True Then
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 4): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                        End With
                       
                       If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 4) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                          End If
                       End If
                       If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/15/", 4) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                          End If
                       End If
                    Else
                       'Reembolso parcial
                        Set oCommand = Nothing
                        Set oParameter = Nothing
                        With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 15): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                        End With
                       
                       If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 15) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                          End If
                       End If
                       If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                          If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/", 15) Then
                             bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                          End If
                       End If
                    End If
                ElseIf oRsListaCuentas.Fields!Reembolsado > 0 And oRsListaCuentas.Fields!PorReembolsar > 0 And oDoFactReembolsos.GrabaDefinitivamente = True Then
                    'Reembolso parcial
                    Set oCommand = Nothing
                    Set oParameter = Nothing
                    With oCommand
                            .CommandType = adCmdStoredProc
                            Set .ActiveConnection = oConexion
                            .CommandTimeout = 150
                            .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                            Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 15): .Parameters.Append oParameter
                            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsListaCuentas.Fields!nrocuenta): .Parameters.Append oParameter
                            .Execute
                    End With
                        
                    If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                       If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, 15) Then
                          bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
                       End If
                    End If
                    If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
                       If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsListaCuentas.Fields!nrocuenta, "/1/", 15) Then
                          bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
                       End If
                    End If
                End If
           '
           End If
           oRsListaCuentas.MoveNext
        Loop
    End If
Terminar:
    If bProcesoOK Then
        ReembolsosModificar = True
        oConexion.CommitTrans
    Else
        ReembolsosModificar = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
    Set oFactReembolsos = Nothing
    Set oDoFacturacionReembolsos = Nothing
    Set oFacturacionReembolsos = Nothing
    Set oDoAtencion = Nothing
    Set oAtenciones = Nothing
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing
    Set oFactReembolsosDocumentos = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oFacturacionBienesFinanciam = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function ReembolsosAnular(oDoFactReembolsos As DoFactreembolsos, oDoFactReembolsosDcto As DoFactReembolsosDcto, oRsListaCuentas As Recordset, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcObservacionesAuditoria As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oFactReembolsos As New FactReembolsos
Dim oDoFacturacionReembolsos As New DoFacturacionReembolsos
Dim oFacturacionReembolsos As New FacturacionReembolsos
Dim oDoAtencion As New DOAtencion
Dim oAtenciones As New Atenciones
Dim oFactReembolsosDocumentos As New FactReembolsosDocumentos
Dim oFacturacionBienesFinanciam As New FacturacionBienesFinanciam
Dim oFacturacionServicioFin As New FacturacionServicioFin
Dim bProcesoOK As Boolean
Dim bPacienteNuevo As Boolean
Dim oRsTmp As New Recordset
Dim oRsTmp1 As New Recordset
Dim lcSql As String
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set oFactReembolsos.Conexion = oConexion
    Set oFacturacionReembolsos.Conexion = oConexion
    Set oAtenciones.Conexion = oConexion
    Set oFactReembolsosDocumentos.Conexion = oConexion
    Set oFacturacionServicioFin.Conexion = oConexion
    Set oFacturacionBienesFinanciam.Conexion = oConexion
    '
    oConexion.BeginTrans
    bProcesoOK = True
    '
    oDoFactReembolsos.IdEstadoReembolso = sghEstadoTabla.sghAnulado   'anulado
    If Not oFactReembolsos.modificar(oDoFactReembolsos) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsos.MensajeError: GoTo Terminar
    End If
    '
    oDoFactReembolsosDcto.idFactReembolso = oDoFactReembolsos.idFactReembolso
    oDoFactReembolsosDcto.IdUsuarioAuditoria = oDoFactReembolsos.IdUsuarioAuditoria
    If Not oFactReembolsosDocumentos.modificar(oDoFactReembolsosDcto) Then
        bProcesoOK = False: ms_MensajeError = oFactReembolsosDocumentos.MensajeError: GoTo Terminar
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoFactReembolsos.IdUsuarioAuditoria, "E", oDoFactReembolsos.idFactReembolso, "FactReembolsos", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcObservacionesAuditoria) 'ListBarItems.idListItem
    '

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "ReembolsosAnular"
            Set oParameter = .CreateParameter("@IdFactReembolso", adInteger, adParamInput, 0, oDoFactReembolsos.idFactReembolso): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
    
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF

            With oCommand
                   .CommandType = adCmdStoredProc
                   Set .ActiveConnection = oConexion
                   .CommandTimeout = 150
                   .CommandText = "FacturacionCuentasAtencionActualizarIdEstado"
                   Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, 11): .Parameters.Append oParameter
                   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsTmp.Fields!IdCuentaAtencion): .Parameters.Append oParameter
                   .Execute
             End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            
            If oDoFactReembolsos.idTipoConsumo = 1 Or oDoFactReembolsos.idTipoConsumo = 3 Then
               If Not oFacturacionBienesFinanciam.ActualizaIdEstadoFacturacionPorCuenta(oRsTmp.Fields!IdCuentaAtencion, 1) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionBienesFinanciam.MensajeError: GoTo Terminar
               End If
            End If
            If oDoFactReembolsos.idTipoConsumo = 2 Or oDoFactReembolsos.idTipoConsumo = 3 Then
               If Not oFacturacionServicioFin.ActualizaIdEstadoFacturacionPorCuenta(oRsTmp.Fields!IdCuentaAtencion, "/4/15/", 1) Then
                  bProcesoOK = False: ms_MensajeError = oFacturacionServicioFin.MensajeError: GoTo Terminar
               End If
            End If
            oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
Terminar:
    If bProcesoOK Then
        ReembolsosAnular = True
        oConexion.CommitTrans
    Else
        ReembolsosAnular = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oFactReembolsos = Nothing
    Set oDoFacturacionReembolsos = Nothing
    Set oFacturacionReembolsos = Nothing
    Set oDoAtencion = Nothing
    Set oAtenciones = Nothing
    Set oRsTmp = Nothing
    Set oFactReembolsosDocumentos = Nothing
    Set oFacturacionServicioFin = Nothing
    Set oFacturacionBienesFinanciam = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactReembolsosSeleccionarPorAnioMes(lnAnio As Long, lnMes As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactReembolsosSeleccionarPorAnioMes = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactReembolsosSeleccionarPorAnioMes"
        Set oParameter = .CreateParameter("Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactReembolsosSeleccionarPorAnioMes = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia(mo_conexion As Connection, lnIdFuenteFinanciamiento As Long) As Long
Dim oRsTmp As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia"
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
    End With
        Set oCommand = Nothing
        Set oParameter = Nothing
        
    If oRsTmp.RecordCount > 0 Then
       FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia = oRsTmp.Fields!idTipoConceptoFarmacia
    Else
       FuentesFinanciamientosDevuelveIdTipoConceptoFarmacia = 0
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Function


Function FacturacionBienesPorCodigoTipoFinanciamiento(mo_conexion As Connection, Codigo As String, IdTipoFinanciamiento As Long) As ADODB.Recordset
Dim oTabla As New FacturacionBienesInsumos
    Set oTabla.Conexion = mo_conexion
    Set FacturacionBienesPorCodigoTipoFinanciamiento = oTabla.SeleccionarPorCodigo(Codigo, IdTipoFinanciamiento)
    Set oTabla = Nothing
End Function

Function FactReembolsosDocumentosSeleccionarPorId(idFactReembolso As Long) As DoFactReembolsosDcto
Dim oConexion As New ADODB.Connection
Dim oFactReembolsosDctos As New FactReembolsosDocumentos
Dim DoFactreembolsosDctos As New DoFactReembolsosDcto
    Set FactReembolsosDocumentosSeleccionarPorId = Nothing
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oFactReembolsosDctos.Conexion = oConexion
    DoFactreembolsosDctos.idFactReembolso = idFactReembolso
    If oFactReembolsosDctos.SeleccionarPorId(DoFactreembolsosDctos) Then
        Set FactReembolsosDocumentosSeleccionarPorId = DoFactreembolsosDctos
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oFactReembolsosDctos.MensajeError
    Set oConexion = Nothing
    Set oFactReembolsosDctos = Nothing
End Function

Sub DevuelveRembolsoPagadoDeFarmaciaYservicioPorSeparado(lnIdCuentaAtencion As Long, lnAreaTramitaReembolso As Long, lnFuenteFinanciamiento As Long, lnMontoReembolsado As Double, lnIdTipoConsumo As Long, lnConsumoF3 As Double, lnAdelantosF3 As Double, lnReembolsoPagadosF3 As Double, ByRef lnReembolsoFarmacia As Double, ByRef lnReembolsoServicio As Double)
    lnReembolsoFarmacia = 0
    lnReembolsoServicio = 0
    Select Case lnIdTipoConsumo
    Case 1      'Farmacia
        lnReembolsoFarmacia = lnMontoReembolsado
        lnReembolsoServicio = 0
    Case 2      'Servicios
        lnReembolsoFarmacia = 0
        lnReembolsoServicio = lnMontoReembolsado
    Case 3      'ambos (Farmacia+servicios)
        If lnAdelantosF3 > lnConsumoF3 Then
           lnConsumoF3 = 0
        Else
           lnConsumoF3 = lnConsumoF3 - lnAdelantosF3
        End If
        lnConsumoF3 = lnConsumoF3 - lnReembolsoPagadosF3
        If lnMontoReembolsado < lnConsumoF3 Then
            lnReembolsoFarmacia = lnMontoReembolsado
            lnReembolsoServicio = 0
        Else
            If lnConsumoF3 = 0 Then
                lnReembolsoFarmacia = 0
                lnReembolsoServicio = lnMontoReembolsado
            Else
                lnReembolsoFarmacia = lnConsumoF3
                lnReembolsoServicio = lnMontoReembolsado - lnConsumoF3
            End If
        End If
    End Select
End Sub
'*********Reembolsos: solo cuando el TipoConsumo=Farmacia+servicios (ambos)
'*********            se usa en Agregar y Modificar Reembolso
Sub ActualizaRsTemporalConReembolsosPagadosAdelantosConsumos(lnIdFactReembolso As Long, lnAreaTramitaReembolso As Long, lnFuenteFinanciamiento As Long, lnTipoConsumo As Long, ByRef oRsCuentasRembolsos As Recordset, oConexion As Connection)
    Dim lnReembolsosPagados As Double, lnAdelantosF  As Double, lnConsumoF As Double, lnConsumoS As Double
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim oRsBuscaXcta As New Recordset
    Dim lnReembolsoPagadoFarmacia As Double, lnReembolsoPagadoServicio As Double
    oRsCuentasRembolsos.MoveFirst
    Do While Not oRsCuentasRembolsos.EOF
       If oRsCuentasRembolsos.Fields!Reembolsado > 0 Then
        If lnTipoConsumo = 3 Then
            '****Busca Reembolsos anteriores: farmacia + Servicios + ambos
            lnReembolsosPagados = 0
            lnConsumoS = 0
            Set oRsBuscaXcta = FacturacionReembolsosSeleccionarPorCuentaPlanAreats(oRsCuentasRembolsos.Fields!nrocuenta, lnAreaTramitaReembolso, lnFuenteFinanciamiento, 0)
            If oRsBuscaXcta.RecordCount > 0 Then
                oRsBuscaXcta.MoveFirst
                Do While Not oRsBuscaXcta.EOF
                   If lnIdFactReembolso <> oRsBuscaXcta.Fields!idFactReembolso Then
                      lnReembolsosPagados = lnReembolsosPagados + oRsBuscaXcta.Fields!ReembolsoPagadoFarmacia
                   End If
                   oRsBuscaXcta.MoveNext
                Loop
            End If
            lnAdelantosF = RetornaImporteDePagosAdelantadosPorNroCuenta(oRsCuentasRembolsos.Fields!nrocuenta, oConexion)
            lnConsumoF = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(oRsCuentasRembolsos.Fields!nrocuenta)
        Else
            lnReembolsosPagados = 0
            lnAdelantosF = 0
            lnConsumoF = 0
            lnConsumoS = 0
            If lnTipoConsumo = 1 Then
               lnConsumoS = RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(oRsCuentasRembolsos.Fields!nrocuenta)
            Else
               lnConsumoF = mo_ReglasFarmacia.RetornaConsumoPacienteFarmaciaConSeguroPorNroCuenta(oRsCuentasRembolsos.Fields!nrocuenta)
            End If
        End If
        DevuelveRembolsoPagadoDeFarmaciaYservicioPorSeparado oRsCuentasRembolsos.Fields!nrocuenta, lnAreaTramitaReembolso, lnFuenteFinanciamiento, oRsCuentasRembolsos.Fields!Reembolsado, lnTipoConsumo, lnConsumoF, lnAdelantosF, lnReembolsosPagados, lnReembolsoPagadoFarmacia, lnReembolsoPagadoServicio
        oRsCuentasRembolsos.Fields!adelantosF = lnAdelantosF
        oRsCuentasRembolsos.Fields!ReembolsosPagadosF = lnReembolsosPagados
        oRsCuentasRembolsos.Fields!consumoF = lnConsumoF
        oRsCuentasRembolsos.Fields!ReembolsoPagadoFarmacia = lnReembolsoPagadoFarmacia
        oRsCuentasRembolsos.Fields!ReembolsoPagadoServicio = lnReembolsoPagadoServicio
        oRsCuentasRembolsos.Fields!consumoS = lnConsumoS
        oRsCuentasRembolsos.Update
      Else
        oRsCuentasRembolsos.Fields!ReembolsoPagadoFarmacia = 0
        oRsCuentasRembolsos.Fields!ReembolsoPagadoServicio = 0
        oRsCuentasRembolsos.Update
      End If
      oRsCuentasRembolsos.MoveNext
    Loop
    Set mo_ReglasFarmacia = Nothing
    Set oRsBuscaXcta = Nothing
End Sub

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionReembolsosSeleccionarPorAnioMes(lnAnio As Long, lnMes As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionReembolsosSeleccionarPorAnioMes = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionReembolsosSeleccionarPorAnioMes"
        Set oParameter = .CreateParameter("Anio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionReembolsosSeleccionarPorAnioMes = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FacturacionReembolsosSeleccionarPorCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionReembolsosSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionReembolsosSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionReembolsosSeleccionarPorCuenta = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function DevuelveRedondeoEnCadaBoletaFarmacia(lnTotalBoleta As Double, lnTotalDctosEnBoleta As Double, lnIdComprobantePago As Long, lnIdEstadoComprobante As Long, lcCadenaConexion As String) As Double
    Dim lcSql As String
    Dim oRsTmpReglaFacturacion As New Recordset
    Dim lnTotalSinRedondeo As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion

    If lnIdEstadoComprobante = 4 Then
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveRedondeoEnCadaBoletaFarmacia"
         Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmpReglaFacturacion = .Execute
        Set oRsTmpReglaFacturacion.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
        
        
        
        If oRsTmpReglaFacturacion.RecordCount > 0 Then
           lnTotalSinRedondeo = 0
           oRsTmpReglaFacturacion.MoveFirst
           Do While Not oRsTmpReglaFacturacion.EOF
              lnTotalSinRedondeo = lnTotalSinRedondeo + oRsTmpReglaFacturacion.Fields!Importe
              oRsTmpReglaFacturacion.MoveNext
           Loop
           DevuelveRedondeoEnCadaBoletaFarmacia = (lnTotalBoleta - lnTotalDctosEnBoleta) - lnTotalSinRedondeo
        Else
           DevuelveRedondeoEnCadaBoletaFarmacia = 0
        End If
        oRsTmpReglaFacturacion.Close
        
    Else
        DevuelveRedondeoEnCadaBoletaFarmacia = 0
    End If
    Set oRsTmpReglaFacturacion = Nothing
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function TiposFinanciamientosSeleccionarPorGeneraPagos(lnGeneraPago As sghTipoFinanciamientoGeneraPago) As Recordset
    Dim oRsTmpReglaFacturacion As New Recordset
    
    
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim LnTipoGeneraPago As Integer
Dim oConexion As New ADODB.Connection
    Set TiposFinanciamientosSeleccionarPorGeneraPagos = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Select Case lnGeneraPago
            Case sghTodosLosQuePaganEnCaja, sghSoloSeguroSIS, sghSoloSeguroSOAT, sghSoloSeguroConvenios
                LnTipoGeneraPago = 1 'parametro genera pago
            Case sghTodosLosQueTienenAlgunSeguro
                LnTipoGeneraPago = 2 ' genera pago 2,3,4
    End Select
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientosSeleccionarPorGeneraPagos"
        Set oParameter = .CreateParameter("@generaPago", adInteger, adParamInput, 0, lnGeneraPago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoGeneraPago", adInteger, adParamInput, 0, LnTipoGeneraPago): .Parameters.Append oParameter
        Set oRsTmpReglaFacturacion = .Execute
        Set oRsTmpReglaFacturacion.ActiveConnection = Nothing
   End With
   Set TiposFinanciamientosSeleccionarPorGeneraPagos = oRsTmpReglaFacturacion
   oConexion.Close
   Set oRsTmpReglaFacturacion = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
 
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function TiposFinanciamientosTarifaSeleccionarPorPlan(lnIdFuenteFinanciamiento As Long) As Recordset
    Dim oRsTmpReglaFacturacion As New Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    Set TiposFinanciamientosTarifaSeleccionarPorPlan = Nothing
    

     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientosTarifaSeleccionarPorPlan"
        Set oParameter = .CreateParameter("@lnIdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRsTmpReglaFacturacion = .Execute
        Set oRsTmpReglaFacturacion.ActiveConnection = Nothing
   End With
   Set TiposFinanciamientosTarifaSeleccionarPorPlan = oRsTmpReglaFacturacion
   oConexion.Close
   Set oRsTmpReglaFacturacion = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta(lnIdtipoFinanciamiento As Long, oConexion As Connection) As sghComoSeTrabajaEnEstadoCuentaLosSeguros
    Dim oRsTmpReglaFacturacion As New Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta"
        Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
        Set oRsTmpReglaFacturacion = .Execute
        Set oRsTmpReglaFacturacion.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
     
    TiposFinanciamientoDevuelveComoSeTrabajaEnEstadoCuenta = oRsTmpReglaFacturacion.Fields!GeneraPago
End Function


Function DevuelveOrdenesPagoSegunCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As String
    Dim lcOrdenesPagos As String
    Dim oRsTmp3 As New ADODB.Recordset
    Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
    lcOrdenesPagos = ""
    Set oRsTmp3 = mo_AdminFacturacion.FactOrdenServicioPagosPorIdCuentaAtencion(lnIdCuentaAtencion)
    If oRsTmp3.RecordCount > 0 Then
       lcOrdenesPagos = "Orden Pago: "
       oRsTmp3.MoveFirst
       Do While Not oRsTmp3.EOF
          lcOrdenesPagos = lcOrdenesPagos & Trim(Str(oRsTmp3.Fields!idOrdenPago)) & ", "
          oRsTmp3.MoveNext
       Loop
    End If
    Set oRsTmp3 = Nothing
    Set mo_AdminFacturacion = Nothing
    DevuelveOrdenesPagoSegunCuenta = lcOrdenesPagos
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function DevuelveDescripcionTipoServicioSegunIdTipoServicio(lnIdTipoServicio As Long) As String
    Dim lcDescripcion As String
    Dim oRsTmp3 As New ADODB.Recordset
    lcDescripcion = ""
    

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "DevuelveDescripcionTipoServicioSegunIdTipoServicio"
        Set oParameter = .CreateParameter("@IdTipoServicio", adInteger, adParamInput, 0, lnIdTipoServicio): .Parameters.Append oParameter
        Set oRsTmp3 = .Execute
        Set oRsTmp3.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
    
    If oRsTmp3.RecordCount > 0 Then
          lcDescripcion = Trim(oRsTmp3.Fields!Descripcion)
    End If
    Set oRsTmp3 = Nothing
    DevuelveDescripcionTipoServicioSegunIdTipoServicio = lcDescripcion
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function AtencionesEstanciaHospitalariaSeleccionarPorIdCuentaAtencion(lnIdCuentaAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set AtencionesEstanciaHospitalariaSeleccionarPorIdCuentaAtencion = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaSeleccionarPorIdCuentaAtencion"
        Set oParameter = .CreateParameter("IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesEstanciaHospitalariaSeleccionarPorIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FacturacionServicioPagosPorIdCuentaIdProducto(lnIdCuentaAtencion As Long, lnIdProducto As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FacturacionServicioPagosPorIdCuentaIdProducto = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosPorIdCuentaIdProducto"
        Set oParameter = .CreateParameter("IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosPorIdCuentaIdProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Conexion Base de datos
Function FactOrdenServicioPorFechasLabPaciente(ldFechaIni As Date, ldFechaFin As Date, lnIdPuntoCarga As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set FactOrdenServicioPorFechasLabPaciente = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPorFechasLabPaciente"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaIni): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPorFechasLabPaciente = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function RetornaImporteDePagosXdevolucionesPorNroCuenta(lnNroCuenta As Long, oConexion As Connection) As Double
   Dim lnPagosXdevoluciones  As Double
   Dim oRsTmp1 As New Recordset
   Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
   Set oRsTmp1 = mo_AdminCaja.CajaComprobantesPagoSeleccionarPagosACuentaPorIdCuenta(lnNroCuenta, Val(lcBuscaParametro.SeleccionaFilaParametro(265)), oConexion)
   lnPagosXdevoluciones = 0
   If oRsTmp1.RecordCount > 0 Then
         oRsTmp1.MoveFirst
         Do While Not oRsTmp1.EOF
            lnPagosXdevoluciones = lnPagosXdevoluciones + oRsTmp1.Fields!total
            oRsTmp1.MoveNext
         Loop
   End If
   Set mo_AdminCaja = Nothing
   Set oRsTmp1 = Nothing
   RetornaImporteDePagosXdevolucionesPorNroCuenta = lnPagosXdevoluciones
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function DevuelveSiPagoConsultaMedicaEnCaja(lnIdAtencion As Long, lcMotivoSalidaHistoria As String, Optional oConexion1 As Connection) As Recordset
Dim oRsCitaPagada199 As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lcSql As String
Dim oConexion As New ADODB.Connection
On Error GoTo ErrCtas1
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "DevuelveSiPagoConsultaMedicaEnCaja"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MotivoSalidaHistoria", adVarChar, adParamInput, 1, lcMotivoSalidaHistoria): .Parameters.Append oParameter
        Set oRsCitaPagada199 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsCitaPagada199.ActiveConnection = Nothing
        End If
   End With
   Set DevuelveSiPagoConsultaMedicaEnCaja = oRsCitaPagada199
   Set oCommand = Nothing
   Set oParameter = Nothing
   Set oConexion = Nothing
   Exit Function
ErrCtas1:
    MsgBox Err.Description
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function EmpleadosSeleccionarPorFiltro(lcFiltro As String) As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set EmpleadosSeleccionarPorFiltro = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosSeleccionarPorFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set EmpleadosSeleccionarPorFiltro = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

End Function

'Sirve para Reporte: Resumen C.Costo
'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function ReembolsoDetalleSeleccionaPorIdComprobantePago(lnIdComprobantePago As Long, oConexion As Connection) As Recordset
    Dim oRsTmp As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ReembolsoDetalleSeleccionaPorIdComprobantePago = Nothing
    

     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReembolsoDetalleSeleccionaPorIdComprobantePago"
        Set oParameter = .CreateParameter("@lnIdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set ReembolsoDetalleSeleccionaPorIdComprobantePago = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing


End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
'Sirve para Reporte:  Consumo x punto de carga
Function ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta(lnIdComprobantePago As Long, oConexion As Connection, lnIdCuentaAtencion As Long) As Recordset
    Dim oRsTmp As New Recordset

    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta = Nothing
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
'Sirve para Reporte: Resumen C.Costo y Consumo x punto de carga
Function ReembolsoDetalleItemSeleccionarPorIdCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim oRsTmp As New Recordset
    

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set ReembolsoDetalleItemSeleccionarPorIdCuenta = Nothing
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ReembolsoDetalleItemSeleccionarPorIdCuenta"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   
   Set ReembolsoDetalleItemSeleccionarPorIdCuenta = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FacturacionServiciosPagosSeleccionarPorCuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim oRsTmp As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServiciosPagosSeleccionarPorCuenta = Nothing
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServiciosPagosSeleccionarPorCuenta"
        Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set FacturacionServiciosPagosSeleccionarPorCuenta = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function


'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FacturacionServiciosPagosSeleccionarPorIdComprobantePago(lnIdComprobantePago As Long, oConexion As Connection) As Recordset
    Dim oRsTmp As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServiciosPagosSeleccionarPorIdComprobantePago = Nothing
    
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServiciosPagosSeleccionarPorIdComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set FacturacionServiciosPagosSeleccionarPorIdComprobantePago = oRsTmp
   Set oRsTmp = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
     
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function ServiciosSeleccionarEmergenciaPorEspecialidad1(lnEspecialidad As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    On Error GoTo ManejadorDeError
    
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
     With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarEmergenciaPorEspecialidad"
        Set oParameter = .CreateParameter("@Especialidad", adInteger, adParamInput, 0, lnEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosSeleccionarEmergenciaPorEspecialidad1 = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function ServiciosSeleccionarCEPorEspecialidad(lnEspecialidad As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""
    

    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarCEPorEspecialidad"
        Set oParameter = .CreateParameter("@Especialidad", adInteger, adParamInput, 0, lnEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ServiciosSeleccionarCEPorEspecialidad = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing

    Exit Function 'Actualizado 07102014
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function AtencionesSeleccionarPorIdAtencion(lnIdAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    Set AtencionesSeleccionarPorIdAtencion = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ConsultaatencionesporidAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorIdAtencion = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function CatalogoBienesInsumosHospSeleccionarXIdProducto(lnIdProducto As Long) As Recordset
     Dim lcSql As String, oRsTmp1 As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     Set CatalogoBienesInsumosHospSeleccionarXIdProducto = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "CatalogoBienesInsumosHospSeleccionarXIdProducto"
         Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set CatalogoBienesInsumosHospSeleccionarXIdProducto = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
       
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioPagosSeleccionarXidProducto(lnIdProducto As Long) As Recordset
     Dim lcSql As String, oRsTmp1 As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     Set FacturacionServicioPagosSeleccionarXidProducto = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "FacturacionServicioPagosSeleccionarXidProducto"
         Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set FacturacionServicioPagosSeleccionarXidProducto = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProducto As Long, lnIdtipoFinanciamiento As Long, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New Connection
        
        If oConexion1 Is Nothing Then
            oConexion = New Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
        
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento"
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            If oConexion1 Is Nothing Then
               Set oRsTmp1.ActiveConnection = Nothing
            End If
        End With
        Set CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento = oRsTmp1
           Set oConexion = Nothing
        Set oCommand = Nothing
        Set oParameter = Nothing
        

End Function


'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function CatalogoServiciosHospEliminarXidProducto(lnIdProducto As Long) As Recordset
     Dim lcSql As String, oRsTmp1 As New Recordset
        
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     Set CatalogoServiciosHospEliminarXidProducto = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "CatalogoServiciosHospEliminarXidProducto"
         Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
         Set oRsTmp1 = .Execute
    End With
    Set CatalogoServiciosHospEliminarXidProducto = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
        
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function CatalogoServiciosHospSeleccionarXidProducto(lnIdProducto As Long) As Recordset
     Dim lcSql As String, oRsTmp1 As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     Set CatalogoServiciosHospSeleccionarXidProducto = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "CatalogoServiciosHospSeleccionarXidProducto"
         Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set CatalogoServiciosHospSeleccionarXidProducto = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Sub FactCatalogoServiciosActualizaCentroCostoXproducto(lnIdCentroCosto As Long, lnIdProducto As Long)
     Dim lcSql As String, oRsTmp1 As New Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "FactCatalogoServiciosActualizaCentroCostoXproducto"
         Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, lnIdCentroCosto): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
          .Execute
    End With
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
        
End Sub

Sub FactCatalogoServiciosActualizaCentroCostoXproductoXcentroCosto(lnIdCentroCosto As Long, lnIdProducto As Long)
'        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
  

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactCatalogoServiciosActualizaCentroCostoXproductoXcentroCosto"
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, lnIdCentroCosto): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing

        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

Sub FactCatalogoServiciosActualizaCentroCostoXcentroCosto(lnIdCentroCosto As Long)
'        Dim lcSql As String, oRsTmp1 As New Recordset
        
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactCatalogoServiciosActualizaCentroCostoXcentroCosto"
            Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, lnIdCentroCosto): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing

        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

Sub FactCatalogoBienesInsumosActualizaCentroCostoXcentroCosto(lnIdCentroCosto As Long)
'        Dim lcSql As String, oRsTmp1 As New Recordset
        
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactCatalogoBienesInsumosActualizaCentroCostoXcentroCosto"
            Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, lnIdCentroCosto): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing

        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarTodos() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
     Set CatalogoServiciosSeleccionarTodos = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "CatalogoServiciosSeleccionarTodos"
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set CatalogoServiciosSeleccionarTodos = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing

End Function

'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function FactCatalogoPaqueteSeleccionarPorFiltro(lcFiltro As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set FactCatalogoPaqueteSeleccionarPorFiltro = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoPaqueteSeleccionarPorFiltro"
                Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
                Set oRsTmp1 = .Execute
                Set oRsTmp1.ActiveConnection = Nothing
           End With
           Set FactCatalogoPaqueteSeleccionarPorFiltro = oRsTmp1
           oConexion.Close
           Set oRsTmp1 = Nothing
           Set oConexion = Nothing
           Set oCommand = Nothing
           Set oParameter = Nothing
        
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FacturacionCatalogoPaquetesSeleccionarTodos() As Recordset
     Dim lcSql As String, oRsTmp1 As New Recordset

     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
     Set FacturacionCatalogoPaquetesSeleccionarTodos = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "FacturacionCatalogoPaquetesSeleccionarTodos"
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set FacturacionCatalogoPaquetesSeleccionarTodos = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

Sub FacturacionCatalogoPaquetesEliminarXidFactPaquete(lnIdFactPaquete As Long)
'        Dim lcSql As String, oRsTmp1 As New Recordset
        
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionCatalogoPaquetesEliminarXidFactPaquete"
            Set oParameter = .CreateParameter("@IdFactPaquete", adInteger, adParamInput, 0, lnIdFactPaquete): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing

        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
        
End Sub

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function FacturacionPaquetesSeleccionarPorFiltro(lcFiltro As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set FacturacionPaquetesSeleccionarPorFiltro = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionPaquetesSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
            Set oRsTmp1 = .Execute
            Set oRsTmp1.ActiveConnection = Nothing
       End With
       Set FacturacionPaquetesSeleccionarPorFiltro = oRsTmp1
       oConexion.Close
       Set oRsTmp1 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
  
End Function

Sub FactCatalogoPaqueteEliminarXidFactPaquete(lnIdFactPaquete As Long)
'        Dim lcSql As String, oRsTmp1 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
  
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactCatalogoPaqueteEliminarXidFactPaquete"
            Set oParameter = .CreateParameter("@IdFactPaquete", adInteger, adParamInput, 0, lnIdFactPaquete): .Parameters.Append oParameter
            .Execute
        End With
        Set oParameter = Nothing
        
        oConexion.Close
        Set oConexion = Nothing
        Set oCommand = Nothing
End Sub

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionCatalogoPaquetesSeleccionarPorIdFactPaquete(lnIdFactPaquete As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set FacturacionCatalogoPaquetesSeleccionarPorIdFactPaquete = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCatalogoPaquetesSeleccionarPorIdFactPaquete"
        Set oParameter = .CreateParameter("@IdFactPaquete", adInteger, adParamInput, 0, lnIdFactPaquete): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FacturacionCatalogoPaquetesSeleccionarPorIdFactPaquete = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function CatalogoBienesSeleccionarPorIdProducto(lnIdProducto As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set CatalogoBienesSeleccionarPorIdProducto = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set CatalogoBienesSeleccionarPorIdProducto = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function CatalogoServiciosSeleccionarPorIdProductoDebb(lnIdProducto As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CatalogoServiciosSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set CatalogoServiciosSeleccionarPorIdProductoDebb = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-09-13 Se cambio a Store Procedure
Function FuentesFinanciamientoSeleccionarSoloConSeguros() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Set FuentesFinanciamientoSeleccionarSoloConSeguros = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FuentesFinanciamientoSeleccionarSoloConSeguros"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FuentesFinanciamientoSeleccionarSoloConSeguros = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
 End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function TiposFinanciamientoSeleccionarSoloConPlan() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Set TiposFinanciamientoSeleccionarSoloConPlan = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientoSeleccionarSoloConPlan"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set TiposFinanciamientoSeleccionarSoloConPlan = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
    
End Function


'Modificado por Yamill Palomino 03-10-13 Se cambio a Store Procedure
Function ServiciosSeleccionarPorFiltro(lcFiltro As String, lnOrdenadoPor As sghTipoOrden) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set ServiciosSeleccionarPorFiltro = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarPorFiltro"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@OrdenadoPor", adVarChar, adParamInput, 50, IIf(lnOrdenadoPor = sghPorCodigo, " order by codigo", " order by nombre")): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set ServiciosSeleccionarPorFiltro = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 30-10-13 Se cambio a Store Procedure
Function FactOrdenServicioSeleccionarSoloAdmisionCE(lnIdAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioSeleccionarSoloAdmisionCE"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioSeleccionarSoloAdmisionCE = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 30-10-13 Se cambio a Store Procedure
Function FacturacionPaquetesSeleccionarPorId(lcFiltro As String) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    Set FacturacionPaquetesSeleccionarPorId = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionPaquetesSeleccionarPorFiltro1"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 500, lcFiltro): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FacturacionPaquetesSeleccionarPorId = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
            
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function ServiciosFacturadosPorIdAtencion(lnIdAtencion As Long) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    Set ServiciosFacturadosPorIdAtencion = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosFacturadosPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set ServiciosFacturadosPorIdAtencion = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function CitasMuestraAteriores(lnIdPaciente As Long, Optional oConexion1 As Connection) As Recordset
        Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    Set CitasMuestraAteriores = Nothing
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "CitasMuestraAteriores"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        If oConexion1 Is Nothing Then
           Set oRsTmp1.ActiveConnection = Nothing
        End If
   End With
   Set CitasMuestraAteriores = oRsTmp1
   If oConexion1 Is Nothing Then
      oConexion.Close
      Set oRsTmp1 = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function ServiciosDevuelveIdDepartamento(lnIdServicio As Long) As Long
   Dim oRsTmp As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosDevuelveIdDepartamento"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
   
   If oRsTmp.RecordCount > 0 Then
      ServiciosDevuelveIdDepartamento = oRsTmp.Fields!IdDepartamento
   Else
      ServiciosDevuelveIdDepartamento = 0
   End If
   Set oRsTmp = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function ProveedoresSeleccionarPorRUC(lcRuc As String) As Recordset
   Dim oRsTmp As New Recordset
   Dim lcSql As String

    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
    Dim oParameter As ADODB.Parameter
    Set ProveedoresSeleccionarPorRUC = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ProveedoresSeleccionarPorRUC"
        Set oParameter = .CreateParameter("@Ruc", adVarChar, adParamInput, 50, lcRuc): .Parameters.Append oParameter
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
   End With
   Set ProveedoresSeleccionarPorRUC = oRsTmp
   oConexion.Close
   Set oRsTmp = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function SeleccionaServicios(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTComprobante As Long, oConexion As Connection) As ADODB.Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaServicios"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionaServicios = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function SeleccionaFarmacia(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTComprobante As Long, oConexion As Connection) As ADODB.Recordset
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaFarmacia"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionaFarmacia = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function SeleccionaDatosCajeroRpt(lnIdCajero As Long, lnDatoDelEmpleado As sghDatoDelEmpleado, oConexion As Connection) As String
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
' Dim ssql As String
  Dim oParameter As ADODB.Parameter

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaDatosCajeroRpt"
        Set oParameter = .CreateParameter("@lnIdCajero", adInteger, adParamInput, 0, lnIdCajero): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
  
  Select Case lnDatoDelEmpleado
    Case sghIniciales
      SeleccionaDatosCajeroRpt = UCase(Left(oRecordset.Fields!ApellidoPaterno, 1) & Left(oRecordset.Fields!ApellidoMaterno, 1) & Left(oRecordset.Fields!Nombres, 1))
    Case sghUsuario
      SeleccionaDatosCajeroRpt = oRecordset.Fields!Usuario
    Case sghApellidosYnombres
      SeleccionaDatosCajeroRpt = Trim(oRecordset.Fields!ApellidoPaterno) & " " & Trim(oRecordset.Fields!ApellidoMaterno) & " " & Trim(oRecordset.Fields!Nombres)
  End Select
  Exit Function
  
ManejadorDeError:
  MsgBox Err.Description
  Exit Function
End Function

Function SeleccionaServiciosDebb(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTComprobante As Long, oConexion As Connection, lnIdCajero As Long, ml_IdTipoReporte As Long) As ADODB.Recordset
  'On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
'  Dim sSQL1 As String
'  Dim Caja As String, Turno As String, TComprobante As String, lcCajero As String, lcTipoOrden As String
  Dim lcFiltro As String
  Dim oParameter As ADODB.Parameter
  Set SeleccionaServiciosDebb = Nothing
  
    
  If IdCaja <> 0 Then lcFiltro = "idCaja=" & IdCaja & " AND "
  If IdTurno <> 0 Then lcFiltro = lcFiltro & "idTurno=" & IdTurno & " AND "
  If lnIdCajero <> 0 Then lcFiltro = lcFiltro & "idCAjero=" & lnIdCajero & " AND "
  If idTComprobante <> 0 Then lcFiltro = lcFiltro & "IdTipoComprobante=" & idTComprobante & " AND "
  If ml_IdTipoReporte <> 0 Then
     lcFiltro = lcFiltro & IIf(ml_IdTipoReporte = 2, "IdTipoOrden=1 AND ", "IdTipoOrden>1 AND ")
  End If
      
  If lcFiltro <> "" Then
        lcFiltro = Mid(lcFiltro, 1, Len(lcFiltro) - 5)
  End If
    
  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaServiciosDebb"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(FI, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(FF, sighentidades.DevuelveFechaSoloFormato_DMY_HMS)): .Parameters.Append oParameter
        Set oRecordset = .Execute
  End With
  If lcFiltro <> "" Then
     oRecordset.Filter = "(" & lcFiltro & ")"
  End If
  Set SeleccionaServiciosDebb = oRecordset
   
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function SeleccionaFarmaciaDebbSql2000(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTComprobante As Long, oConexion As Connection) As ADODB.Recordset
  'On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim sSQL1 As String
  Dim lcFiltro As String
  Dim Caja As String, Turno As String, TComprobante As String
  Dim oParameter As ADODB.Parameter
  

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaFarmaciaDebbSql2000"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adVarChar, adParamInput, 50, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adVarChar, adParamInput, 50, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionaFarmaciaDebbSql2000 = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function ChequeaSiEsPagosAcuenta(lnIdComprobantePago As Long, oConexion As Connection, lnIdPagosACuenta As Long, _
                                 ByRef lnTotalRedondeo As Double, IdTipoOrden As Long, lnExoneraciones As Double, _
                                 lnAdelantos As Double, lnIdEstadoComprobante As Long, _
                                 lnTotalBoleta As Double) As Boolean
'    Dim lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnTotalItems As Double
    Dim lbNOtieneDETALLE As Boolean
    ChequeaSiEsPagosAcuenta = False
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ChequeaSiEsPagosAcuenta"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPagosACuenta", adInteger, adParamInput, 0, lnIdPagosACuenta): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoOrden", adInteger, adParamInput, 0, IdTipoOrden): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set oCommand = Nothing
   Set oParameter = Nothing
    lbNOtieneDETALLE = True
    lnTotalItems = 0
    If oRsTmp1.RecordCount > 0 Then
       lbNOtieneDETALLE = False
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          If lnIdPagosACuenta = oRsTmp1.Fields!idProducto And IdTipoOrden = 1 Then
             ChequeaSiEsPagosAcuenta = True
          End If
          lnTotalItems = lnTotalItems + oRsTmp1.Fields!total
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    '***** Redondeo ********
    If lbNOtieneDETALLE = True Then
        lnTotalRedondeo = 9999    'no puede haber BOLETA SIN DETALLE
    Else
        lnTotalRedondeo = 0
        If lnIdEstadoComprobante = 4 Then
           lnTotalRedondeo = Round((lnTotalBoleta + lnExoneraciones + lnAdelantos) - lnTotalItems, 3)
        End If
    End If
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FarmaciaVendedorChequeaSiEsSuBoleta(lnIdComprobantePago As Long, lnIdVendedorFarmacia As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set FarmaciaVendedorChequeaSiEsSuBoleta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaVendedorChequeaSiEsSuBoleta"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdVendedorFarmacia", adInteger, adParamInput, 0, lnIdVendedorFarmacia): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FarmaciaVendedorChequeaSiEsSuBoleta = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioDespachoSeleccionarXidCuenta(ml_idCuentaAtencion As Long) As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set FacturacionServicioDespachoSeleccionarXidCuenta = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDespachoSeleccionarXidCuenta"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FacturacionServicioDespachoSeleccionarXidCuenta = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function Sismedv2SeleccionarPrecios(oConexionFox As Connection, lcTablaDBF As String, lcWhere As String) As Recordset
    Dim lcSql As String, oRsFox1 As New Recordset
    lcSql = "select * from " & lcTablaDBF & lcWhere
    oRsFox1.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
   Set Sismedv2SeleccionarPrecios = oRsFox1
    
 End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function TiposFinanciamientoSeleccionarSoloIngresaPrecios() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Set TiposFinanciamientoSeleccionarSoloIngresaPrecios = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "TiposFinanciamientoSeleccionarSoloIngresaPrecios"
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set TiposFinanciamientoSeleccionarSoloIngresaPrecios = oRsTmp1
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion(lnCuenta As Long, oConexion As Connection) As ADODB.Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion = Nothing
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion"
        Set oParameter = .CreateParameter("@Cuenta", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion = oRsTmp1
   Set oCommand = Nothing
   Set oRsTmp1 = Nothing
   Set oParameter = Nothing
    
   End Function

Function FacturacionServicioPagosPorCuentaTodosConnection(lnCuenta As Long, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionServicioPagosPorCuentaTodosConnection = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosPorCuentaTodos"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosPorCuentaTodosConnection = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Sub FacturacionCuentasAtencionPtosActualizar(ml_idCuenta As Long, lbTodasLasCuentas As Boolean, lnNumeroMinutosDeProceso As Long)
10            On Error GoTo ErCtaCta
              Dim mrs_Tmp As New Recordset
              Dim lcBuscaParametro As New SIGHDatos.Parametros
              Dim oConexion As New Connection
              Dim rsReporte As New ADODB.Recordset
              Dim rsTmp As New Recordset
              Dim rsTmp10 As New Recordset
              Dim mrs_Tmp1 As New Recordset
              Dim mrs_Ctas As New Recordset
              Dim rsTmpReemb2 As New Recordset, oRsTmpReemb1 As New Recordset
              Dim oRsPuntosDeCarga As New Recordset
              Dim oDoImagMovimientoIngresos As New DoImagMovimientoIngresos
              Dim oDOImagMovimiento As New DoImagMovimiento
              Dim ml_idUsuario As Long
20            Dim lnSaldoInicial As Long: Dim lnSaldofinal As Long
30            Dim lnIngresos As Long: Dim lnSalidas As Long: Dim lnSalidasImg  As Long
              Dim ldFechaPrincipio As Date
40            Dim lcCodigo As String: Dim lcNombre As String: Dim lnIdProducto As Long
              Dim lnPrecio As Double
              Dim lbPrimeraVez As Boolean
50            Dim lcTexto1 As String: Dim lcTexto2 As String: Dim lcTexto3 As String
              Dim lnDebioPagar As Double, lnTotalPorCuenta As Double
              Dim lnIdTipoServicio As Long, lnIdDepartamento As Long, lnIdEspacialidad As Long, lnIdServicioPaciente As Long
              Dim lcDtipoServicio As String, lcDpto As String, lcDespecialidad As String, lcDServicio As String
              Dim lnSalieron As Integer, lnRetornaron As Integer, lnSalieronYnoRetornaron As Integer
              Dim lnTSalieron As Integer, lnTRetornaron As Integer, lnTSalieronYnoRetornaron As Integer
              Dim lcSql As String, lbEsCuentaPagante As Boolean
              Dim lbProcesaEnServidor As Boolean, lcHoraInicio As String, lcHoraFinal As String
              Dim lcParametro245 As String, lcError As String
              Dim lnFarmacia As Double, lnServicio As Double, lnRx As Double, lnLaboratorio As Double, lnOtros As Double, lnPucho As Double, lnDctos As Double
              Dim lnTFarmacia As Double, lnTRx As Double, lnTLaboratorio As Double, lnTOtros As Double, lnTDctos As Double, lnDctosXitem As Double
              Dim lnTtotalFacturadoCuenta As Double, LnTotalPagadoCuenta As Double, lnTotalDeudaCuenta As Double
              Dim lnTotalFacturadoCuenta As Double
              Dim lnTotalConsumos As Double, lnTotalPagos As Double, LnTodasLasCuentas As Long
              Dim lnTotalConsumosCta As Double, lnTotalPagadoCta As Double, lnTotalExoneradoCta As Double, lnTotalReembolsoCta As Double
              Dim lcHoraActual As String, lnCuentasProcesadas As Long, lnIdPuntoCargaActual As Long
              Dim oCommand As New ADODB.Command
              Dim oParameter As ADODB.Parameter
60            If ml_idCuenta = 0 Then
70               Exit Sub
80            End If
90            lcHoraInicio = Time
100           lcParametro245 = lcBuscaParametro.SeleccionaFilaParametro(245)
110           oConexion.CommandTimeout = 300
120           oConexion.CursorLocation = adUseClient
130           oConexion.Open sighentidades.CadenaConexion
              '
140           If lbTodasLasCuentas = True Then
150              LnTodasLasCuentas = 1
160           Else
170              LnTodasLasCuentas = 0
180           End If
190           With oCommand
200             .CommandType = adCmdStoredProc
210             Set .ActiveConnection = oConexion
220             .CommandTimeout = 150
230             .CommandText = "FacturacionCuentasAtencionxIdCuentaAtencion"
240             Set oParameter = .CreateParameter("@TodasLasCuentas", adInteger, adParamInput, 0, LnTodasLasCuentas): .Parameters.Append oParameter
250             Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
260             Set mrs_Tmp1 = .Execute
270             Set mrs_Tmp1.ActiveConnection = Nothing
280           End With
290           Set oParameter = Nothing
300           Set oCommand = Nothing
              '
310           If mrs_Tmp1.RecordCount > 0 Then
320              lnCuentasProcesadas = 0
330              mrs_Tmp1.MoveFirst
340              Do While Not mrs_Tmp1.EOF
350                   If lbTodasLasCuentas = True Then
360                           lcHoraActual = Time
370                           lcSql = DateDiff("n", CDate(lcHoraInicio), CDate(lcHoraActual))
380                           If Val(lcSql) > lnNumeroMinutosDeProceso Then
390                              Exit Do
400                           End If
410                           lnCuentasProcesadas = lnCuentasProcesadas + 1
420                   End If
                      '
430                   ml_idCuenta = mrs_Tmp1.Fields!IdCuentaAtencion
440                   lbEsCuentaPagante = IIf(mrs_Tmp1.Fields!GeneraPago = 1, True, False)
450                   lnTotalConsumosCta = 0: lnTotalPagadoCta = 0: lnTotalExoneradoCta = 0: lnTotalReembolsoCta = 0
                      'Limpia tablas de cabecera y detalle
460                   With oCommand
470                       .CommandType = adCmdStoredProc
480                       Set .ActiveConnection = oConexion
490                       .CommandTimeout = 150
500                       .CommandText = "FacturacionCuentasAtencionPtosEliminarXcuenta"
510                       Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
520                        .Execute
530                   End With
540                   Set oCommand = Nothing
550                   Set oParameter = Nothing
                      '
                      'servicios
560                   If lbEsCuentaPagante = True Then
570                       Set rsTmp = FacturacionServicioPagosPorCuentaTodosConnection(ml_idCuenta, oConexion)
580                       rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
590                   Else
600                       With oCommand
610                         .CommandType = adCmdStoredProc
620                         Set .ActiveConnection = oConexion
630                         .CommandTimeout = 150
640                         .CommandText = "FacturacionServicioFinanciamientosConsultarXIdCuentaAtencion"
650                         Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
660                         Set rsTmp = .Execute
670                         Set rsTmp.ActiveConnection = Nothing
680                       End With
690                       Set oCommand = Nothing
700                       Set oParameter = Nothing
710                       rsTmp.Filter = "(idEstadoFacturacion=1 or idEstadoFacturacion=4)"
720                   End If
730                   If rsTmp.RecordCount > 0 Then
740                      rsTmp.MoveFirst
750                      Do While Not rsTmp.EOF
760                         If rsTmp.Fields!totalPorPagar > 0 And rsTmp.Fields!idProducto <> Val(lcParametro245) Then
                              '
770                           lbPrimeraVez = True
780                           Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
790                           If mrs_Tmp.RecordCount > 0 Then
800                              mrs_Tmp.MoveFirst
810                              Do While Not mrs_Tmp.EOF
820                                 If mrs_Tmp.Fields!IdCuentaAtencion = ml_idCuenta And mrs_Tmp.Fields!IdPuntoCarga = rsTmp.Fields!IdPuntoCarga Then
830                                    lbPrimeraVez = False
840                                    Exit Do
850                                 End If
860                                 mrs_Tmp.MoveNext
870                              Loop
880                           End If
                              '
890                           With oCommand
900                             .CommandType = adCmdStoredProc
910                             Set .ActiveConnection = oConexion
920                             .CommandTimeout = 150
930                             .CommandText = "FacturacionCuentasAtencionPtosActualizar"
940                             If lbPrimeraVez = True Then
950                                   lnTotalConsumos = rsTmp.Fields!totalPorPagar
960                                   Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
970                                   Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
980                                   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, rsTmp.Fields!IdPuntoCarga): .Parameters.Append oParameter
990                                   Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, rsTmp.Fields!totalPorPagar): .Parameters.Append oParameter
1000                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1010                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1020                            Else
1030                                  lnTotalConsumos = mrs_Tmp.Fields!TotalConsumos + rsTmp.Fields!totalPorPagar
1040                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
1050                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
1060                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, rsTmp.Fields!IdPuntoCarga): .Parameters.Append oParameter
1070                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, lnTotalConsumos): .Parameters.Append oParameter
1080                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1090                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1100                            End If
1110                            .Execute
1120                          End With
1130                          Set oParameter = Nothing
1140                          Set oCommand = Nothing
                              '
1150                          lnTotalConsumosCta = lnTotalConsumosCta + rsTmp.Fields!totalPorPagar
1160                        End If
1170                        rsTmp.MoveNext
1180                     Loop
1190                  End If
1200                  rsTmp.Close
                      'farmacia
1210                  If lbEsCuentaPagante = True Then
1220                      Set rsTmp = FacturacionBienesPagosSeleccionarPorCuentaTodosOconexion(ml_idCuenta, oConexion)
1230                      rsTmp.Filter = "idEstadoFacturacion=1 or idEstadoFacturacion=4"
1240                  Else
1250                      With oCommand
1260                        .CommandType = adCmdStoredProc
1270                        Set .ActiveConnection = oConexion
1280                        .CommandTimeout = 150
1290                        .CommandText = "FacturacionfarmMovimientoVentasConsultarXIdCuentaAtencion"
1300                        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
1310                        Set rsTmp = .Execute
1320                        Set rsTmp.ActiveConnection = Nothing
1330                      End With
1340                      Set oCommand = Nothing
1350                      Set oParameter = Nothing
1360                      rsTmp.Filter = "totalFinanciado>0 and idTipoFinanciamiento=" & mrs_Tmp1.Fields!IdFormaPago
1370                  End If
1380                  If rsTmp.RecordCount > 0 Then
1390                     rsTmp.MoveFirst
1400                     Do While Not rsTmp.EOF
1410                        If rsTmp.Fields!idEstadoMovimiento = 1 Then
1420                          lbPrimeraVez = True
1430                          Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
1440                          If mrs_Tmp.RecordCount > 0 Then
1450                             mrs_Tmp.MoveFirst
1460                             Do While Not mrs_Tmp.EOF
1470                                If mrs_Tmp.Fields!IdCuentaAtencion = ml_idCuenta And mrs_Tmp.Fields!IdPuntoCarga = 5 Then
1480                                   lbPrimeraVez = False
1490                                   Exit Do
1500                                End If
1510                                mrs_Tmp.MoveNext
1520                             Loop
1530                          End If
                              '
1540                          With oCommand
1550                            .CommandType = adCmdStoredProc
1560                            Set .ActiveConnection = oConexion
1570                            .CommandTimeout = 150
1580                            .CommandText = "FacturacionCuentasAtencionPtosActualizar"
1590                            If lbPrimeraVez = True Then
1600                                  lnTotalConsumos = rsTmp.Fields!TotalFinanciado
1610                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
1620                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
1630                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
1640                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, lnTotalConsumos): .Parameters.Append oParameter
1650                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1660                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1670                            Else
1680                                  lnTotalConsumos = mrs_Tmp.Fields!TotalConsumos + rsTmp.Fields!TotalFinanciado
1690                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
1700                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
1710                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
1720                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, lnTotalConsumos): .Parameters.Append oParameter
1730                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1740                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
1750                            End If
1760                            .Execute
1770                          End With
1780                          Set oParameter = Nothing
1790                          Set oCommand = Nothing
                              '
1800                          lnTotalConsumosCta = lnTotalConsumosCta + rsTmp.Fields!TotalFinanciado
1810                        End If
1820                        rsTmp.MoveNext
1830                     Loop
1840                  End If
1850                  rsTmp.Close
                      'Pagos realizados de la Cuenta
1860                  With oCommand
1870                        .CommandType = adCmdStoredProc
1880                        Set .ActiveConnection = oConexion
1890                        .CommandTimeout = 150
1900                        .CommandText = "FacturacionCajaComprobantesPagoConsultarXIdCuentaAtencion"
1910                        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
1920                        Set rsTmp10 = .Execute
1930                        Set rsTmp10.ActiveConnection = Nothing
1940                  End With
1950                  Set oCommand = Nothing
1960                  Set oParameter = Nothing
                      
1970                  rsTmp10.Filter = "idEstadoComprobante=4"
1980                  If rsTmp10.RecordCount > 0 Then
1990                      lnTFarmacia = 0: lnTRx = 0: lnTLaboratorio = 0: lnTOtros = 0: lnTDctos = 0
2000                      rsTmp10.MoveFirst
2010                      Do While Not rsTmp10.EOF
2020                         If rsTmp10.Fields!idEstadoComprobante = 4 Then   'solo pagados
2030                              lnTotalPagadoCta = lnTotalPagadoCta + rsTmp10.Fields!total
2040                              lnTotalExoneradoCta = lnTotalExoneradoCta + rsTmp10.Fields!exoneraciones + rsTmp10.Fields!dctos
2050                              If lbEsCuentaPagante = True Then
2060                                      lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
2070                                      If rsTmp10.Fields!IdTipoOrden <> 1 Then
                                              '**** Cuenta de un Paciente PAGANTE - BOLETA DE FARMACIA
2080                                          lnTFarmacia = rsTmp10.Fields!total
2090                                          lnServicio = 0
2100                                          lnTDctos = lnTDctos + (rsTmp10.Fields!total - rsTmp10.Fields!Subtotal)
2110                                          Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
2120                                          If mrs_Tmp.RecordCount > 0 Then
2130                                              mrs_Tmp.MoveFirst
2140                                              mrs_Tmp.Find "idPuntoCarga=" & sghPuntosCargaBasicos.sghPtoCargaFarmacia
2150                                          End If
                                              '
2160                                          With oCommand
2170                                            .CommandType = adCmdStoredProc
2180                                            Set .ActiveConnection = oConexion
2190                                            .CommandTimeout = 150
2200                                            .CommandText = "FacturacionCuentasAtencionPtosActualizar"
2210                                            If mrs_Tmp.EOF Then
2220                                                  lnTotalPagos = lnTFarmacia
2230                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
2240                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
2250                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
2260                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2270                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
2280                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2290                                            Else
2300                                                  lnTotalPagos = mrs_Tmp.Fields!TotalPagos + lnTFarmacia
2310                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
2320                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
2330                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
2340                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2350                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
2360                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2370                                            End If
2380                                            .Execute
2390                                          End With
2400                                          Set oParameter = Nothing
2410                                          Set oCommand = Nothing
                                              '
2420                                      Else
                                              '**** Cuenta de un Paciente PAGANTE - BOLETA DE SERVICIOS
2430                                          lnServicio = rsTmp10.Fields!total
2440                                          lnDctos = (rsTmp10.Fields!total - rsTmp10.Fields!Subtotal)
2450                                          lnTDctos = lnTDctos + lnDctos
2460                                          Set oRsTmpReemb1 = FacturacionServiciosPagosSeleccionarPorIdComprobantePago(rsTmp10.Fields!IdComprobantePago, oConexion)
2470                                          If oRsTmpReemb1.RecordCount > 0 Then
2480                                             oRsTmpReemb1.MoveFirst
2490                                             Do While Not oRsTmpReemb1.EOF
2500                                                   lnDctosXitem = 0
2510                                                   If lnDctos <> 0 Then
2520                                                      With oCommand
2530                                                             .CommandType = adCmdStoredProc
2540                                                             Set .ActiveConnection = oConexion
2550                                                             .CommandTimeout = 150
2560                                                             .CommandText = "FacturacionServicioFinanciamientosConsultarX@IdOrdenIdProducto"
2570                                                             Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRsTmpReemb1.Fields!IdOrden): .Parameters.Append oParameter
2580                                                             Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmpReemb1.Fields!idProducto): .Parameters.Append oParameter
2590                                                             Set rsTmpReemb2 = .Execute
2600                                                             Set rsTmpReemb2.ActiveConnection = Nothing
2610                                                      End With
2620                                                      Set oParameter = Nothing
2630                                                      Set oCommand = Nothing
                                                          
2640                                                      If rsTmpReemb2.RecordCount > 0 Then
2650                                                         rsTmpReemb2.MoveFirst
2660                                                         Do While Not rsTmpReemb2.EOF
2670                                                            lnDctosXitem = lnDctosXitem + rsTmpReemb2.Fields!TotalFinanciado
2680                                                            rsTmpReemb2.MoveNext
2690                                                         Loop
2700                                                      End If
2710                                                      rsTmpReemb2.Close
2720                                                   End If
2730                                                   lnIdPuntoCargaActual = IIf(IsNull(oRsTmpReemb1.Fields!IdPuntoCarga), 1, oRsTmpReemb1.Fields!IdPuntoCarga)
2740                                                   Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
2750                                                   If mrs_Tmp.RecordCount > 0 Then
2760                                                      mrs_Tmp.MoveFirst
2770                                                      mrs_Tmp.Find "idPuntoCarga=" & lnIdPuntoCargaActual
2780                                                   End If
                                                       '
2790                                                   With oCommand
2800                                                    .CommandType = adCmdStoredProc
2810                                                    Set .ActiveConnection = oConexion
2820                                                    .CommandTimeout = 150
2830                                                    .CommandText = "FacturacionCuentasAtencionPtosActualizar"
2840                                                    If mrs_Tmp.EOF Then
2850                                                          lnTotalPagos = oRsTmpReemb1.Fields!total - lnDctosXitem
2860                                                          Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
2870                                                          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
2880                                                          Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
2890                                                          Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2900                                                          Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
2910                                                          Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2920                                                    Else
2930                                                          lnTotalPagos = mrs_Tmp.Fields!TotalPagos + (oRsTmpReemb1.Fields!total - lnDctosXitem)
2940                                                          Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
2950                                                          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
2960                                                          Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
2970                                                          Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
2980                                                          Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
2990                                                          Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3000                                                    End If
3010                                                    .Execute
3020                                                   End With
3030                                                   Set oParameter = Nothing
3040                                                   Set oCommand = Nothing
                                                       '
3050                                                   oRsTmpReemb1.MoveNext
3060                                             Loop
3070                                          End If
3080                                          oRsTmpReemb1.Close
3090                                      End If
3100                              Else
                                          'Otros pagos que no cubre el SEGURO
3110                                      lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
3120                                      If rsTmp10.Fields!IdTipoOrden <> 1 Then
                                              '**** Cuenta de un Paciente SEGUROS - BOLETA DE FARMACIA
3130                                          lnTFarmacia = rsTmp10.Fields!total
3140                                          lnServicio = 0
3150                                          lnTDctos = lnTDctos + (rsTmp10.Fields!total - rsTmp10.Fields!Subtotal)
3160                                          Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
3170                                          If mrs_Tmp.RecordCount > 0 Then
3180                                              mrs_Tmp.MoveFirst
3190                                              mrs_Tmp.Find "idPuntoCarga=" & sghPuntosCargaBasicos.sghPtoCargaFarmacia
3200                                          End If
                                              '
3210                                          With oCommand
3220                                            .CommandType = adCmdStoredProc
3230                                            Set .ActiveConnection = oConexion
3240                                            .CommandTimeout = 150
3250                                            .CommandText = "FacturacionCuentasAtencionPtosActualizar"
3260                                            If mrs_Tmp.EOF Then
3270                                                  lnTotalPagos = lnTFarmacia
3280                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
3290                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
3300                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
3310                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3320                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
3330                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3340                                            Else
3350                                                  lnTotalPagos = mrs_Tmp.Fields!TotalPagos + lnTFarmacia
3360                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
3370                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
3380                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
3390                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3400                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
3410                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3420                                            End If
3430                                            .Execute
3440                                          End With
3450                                          Set oParameter = Nothing
3460                                          Set oCommand = Nothing
                                              
3470                                      Else
                                              '**** Cuenta de un Paciente SEGUROS - BOLETA DE SERVICIOS
3480                                          lnServicio = rsTmp10.Fields!total
3490                                          lnDctos = (rsTmp10.Fields!total - rsTmp10.Fields!Subtotal)
3500                                          lnTDctos = lnTDctos + lnDctos
3510                                          Set oRsTmpReemb1 = FacturacionServiciosPagosSeleccionarPorIdComprobantePago(rsTmp10.Fields!IdComprobantePago, oConexion)
3520                                          If oRsTmpReemb1.RecordCount > 0 Then
3530                                             oRsTmpReemb1.MoveFirst
3540                                             Do While Not oRsTmpReemb1.EOF
3550                                                   lnDctosXitem = 0
3560                                                   If lnDctos <> 0 Then
3570                                                      With oCommand
3580                                                             .CommandType = adCmdStoredProc
3590                                                             Set .ActiveConnection = oConexion
3600                                                             .CommandTimeout = 150
3610                                                             .CommandText = "FacturacionServicioFinanciamientosConsultarX@IdOrdenIdProducto"
3620                                                             Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, oRsTmpReemb1.Fields!IdOrden): .Parameters.Append oParameter
3630                                                             Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oRsTmpReemb1.Fields!idProducto): .Parameters.Append oParameter
3640                                                             Set rsTmpReemb2 = .Execute
3650                                                             Set rsTmpReemb2.ActiveConnection = Nothing
3660                                                      End With
3670                                                      Set oParameter = Nothing
3680                                                      Set oCommand = Nothing
                                                          
3690                                                      If rsTmpReemb2.RecordCount > 0 Then
3700                                                         rsTmpReemb2.MoveFirst
3710                                                         Do While Not rsTmpReemb2.EOF
3720                                                            lnDctosXitem = lnDctosXitem + rsTmpReemb2.Fields!TotalFinanciado
3730                                                            rsTmpReemb2.MoveNext
3740                                                         Loop
3750                                                      End If
3760                                                      rsTmpReemb2.Close
3770                                                   End If
3780                                                   lnIdPuntoCargaActual = IIf(IsNull(oRsTmpReemb1.Fields!IdPuntoCarga), 1, oRsTmpReemb1.Fields!IdPuntoCarga)
3790                                                   Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
3800                                                   If mrs_Tmp.RecordCount > 0 Then
3810                                                      mrs_Tmp.MoveFirst
3820                                                      mrs_Tmp.Find "idPuntoCarga=" & lnIdPuntoCargaActual
3830                                                   End If
                                                       '
3840                                                   With oCommand
3850                                                    .CommandType = adCmdStoredProc
3860                                                    Set .ActiveConnection = oConexion
3870                                                    .CommandTimeout = 150
3880                                                    .CommandText = "FacturacionCuentasAtencionPtosActualizar"
3890                                                    If mrs_Tmp.EOF Then
3900                                                          lnTotalPagos = oRsTmpReemb1.Fields!total - lnDctosXitem
3910                                                          Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
3920                                                          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
3930                                                          Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
3940                                                          Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3950                                                          Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
3960                                                          Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
3970                                                    Else
3980                                                          lnTotalPagos = mrs_Tmp.Fields!TotalPagos + (oRsTmpReemb1.Fields!total - lnDctosXitem)
3990                                                          Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
4000                                                          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
4010                                                          Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
4020                                                          Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4030                                                          Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
4040                                                          Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4050                                                    End If
4060                                                    .Execute
4070                                                   End With
4080                                                   Set oParameter = Nothing
4090                                                   Set oCommand = Nothing
                                                       '
4100                                                   oRsTmpReemb1.MoveNext
4110                                             Loop
4120                                          End If
4130                                          oRsTmpReemb1.Close
4140                                      End If
4150                             End If
4160                         End If
4170                         rsTmp10.MoveNext
4180                      Loop
4190                  End If
4200                  rsTmp10.Close
                      'solo para Reporte HRA- Si es un Reembolso
4210                  If lbEsCuentaPagante = False Then
4220                      With oCommand
4230                          .CommandType = adCmdStoredProc
4240                          Set .ActiveConnection = oConexion
4250                          .CommandTimeout = 150
4260                          .CommandText = "FacturacionReembolsosXcuentaAtencion"
4270                          Set oParameter = .CreateParameter("@ml_idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
4280                          Set rsTmp10 = .Execute
4290                          Set rsTmp10.ActiveConnection = Nothing
4300                      End With
4310                      Set oCommand = Nothing
4320                      Set oParameter = Nothing
                          
4330                      If rsTmp10.RecordCount > 0 Then
4340                         rsTmp10.MoveFirst
4350                         Do While Not rsTmp10.EOF
4360                              lnFarmacia = 0: lnServicio = 0: lnRx = 0: lnLaboratorio = 0: lnOtros = 0
4370                              If rsTmp10.Fields!IdComprobantePago > 0 Then
4380                              Set rsTmpReemb2 = ReembolsoDetalleSeleccionaPorIdComprobantePagoIdCuenta(rsTmp10.Fields!IdComprobantePago, oConexion, ml_idCuenta)
4390                              If rsTmpReemb2.RecordCount > 0 Then
4400                                 lnTotalReembolsoCta = lnTotalReembolsoCta + rsTmp10.Fields!total
                                      'Reembolsos
4410                                  rsTmpReemb2.MoveFirst
4420                                  Do While Not rsTmpReemb2.EOF
4430                                     lnFarmacia = lnFarmacia + rsTmpReemb2.Fields!ReembolsoPagadoFarmacia
4440                                     lnServicio = lnServicio + rsTmpReemb2.Fields!ReembolsoPagadoServicio
4450                                     rsTmpReemb2.MoveNext
4460                                  Loop
4470                                  rsTmpReemb2.Close
4480                                  lnTFarmacia = lnTFarmacia + lnFarmacia
4490                                  If lnFarmacia > 0 Then
4500                                          Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
4510                                          If mrs_Tmp.RecordCount > 0 Then
4520                                              mrs_Tmp.MoveFirst
4530                                              mrs_Tmp.Find "idPuntoCarga=" & sghPuntosCargaBasicos.sghPtoCargaFarmacia
4540                                          End If
                                              '
4550                                          With oCommand
4560                                            .CommandType = adCmdStoredProc
4570                                            Set .ActiveConnection = oConexion
4580                                            .CommandTimeout = 150
4590                                            .CommandText = "FacturacionCuentasAtencionPtosActualizar"
4600                                            If mrs_Tmp.EOF Then
4610                                                  lnTotalPagos = lnFarmacia
4620                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
4630                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
4640                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
4650                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4660                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4670                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
4680                                            Else
4690                                                  lnTotalPagos = mrs_Tmp.Fields!TotalPagosReembolso + lnFarmacia
4700                                                  Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
4710                                                  Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
4720                                                  Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, sghPuntosCargaBasicos.sghPtoCargaFarmacia): .Parameters.Append oParameter
4730                                                  Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4740                                                  Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
4750                                                  Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
4760                                            End If
4770                                            .Execute
4780                                          End With
4790                                          Set oParameter = Nothing
4800                                          Set oCommand = Nothing
                                              
                                              '
4810                                  End If
4820                                  If lnServicio > 0 Then
4830                                     Set oRsTmpReemb1 = ReembolsoDetalleItemSeleccionarPorIdCuenta(ml_idCuenta, oConexion)
4840                                     If oRsTmpReemb1.RecordCount > 0 Then
4850                                        oRsTmpReemb1.MoveFirst
4860                                        Do While Not oRsTmpReemb1.EOF
4870                                              If oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre1 Or oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaBancoSangre2 Then
                                                     'Laboratorio
4880                                                 lnLaboratorio = lnLaboratorio + oRsTmpReemb1.Fields!TotalFinanciado
4890                                              ElseIf oRsTmpReemb1.Fields!IdPuntoCarga = sghPuntosCargaBasicos.sghPtoCargaRayosX Then
                                                     'Rx
4900                                                 lnRx = lnRx + oRsTmpReemb1.Fields!TotalFinanciado
4910                                              Else
4920                                                 lnOtros = lnOtros + oRsTmpReemb1.Fields!TotalFinanciado
4930                                              End If
4940                                              lnIdPuntoCargaActual = IIf(IsNull(oRsTmpReemb1.Fields!IdPuntoCarga), 1, oRsTmpReemb1.Fields!IdPuntoCarga)
4950                                              Set mrs_Tmp = FacturacionCuentasAtencionPtosXIdCuenta(ml_idCuenta, oConexion)
4960                                              If mrs_Tmp.RecordCount > 0 Then
4970                                                  mrs_Tmp.MoveFirst
4980                                                  mrs_Tmp.Find "idPuntoCarga=" & lnIdPuntoCargaActual
4990                                              End If
                                                  '
5000                                              With oCommand
5010                                                .CommandType = adCmdStoredProc
5020                                                Set .ActiveConnection = oConexion
5030                                                .CommandTimeout = 150
5040                                                .CommandText = "FacturacionCuentasAtencionPtosActualizar"
5050                                                If mrs_Tmp.EOF Then
5060                                                      lnTotalPagos = oRsTmpReemb1.Fields!TotalFinanciado
5070                                                      Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, True): .Parameters.Append oParameter
5080                                                      Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
5090                                                      Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
5100                                                      Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
5110                                                      Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
5120                                                      Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
5130                                                Else
5140                                                      lnTotalPagos = mrs_Tmp.Fields!TotalPagosReembolso + oRsTmpReemb1.Fields!TotalFinanciado
5150                                                      Set oParameter = .CreateParameter("@lbNueva", adBoolean, adParamInput, 0, False): .Parameters.Append oParameter
5160                                                      Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
5170                                                      Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, lnIdPuntoCargaActual): .Parameters.Append oParameter
5180                                                      Set oParameter = .CreateParameter("@TotalConsumos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
5190                                                      Set oParameter = .CreateParameter("@TotalPagos", adCurrency, adParamInput, 0, 0): .Parameters.Append oParameter
5200                                                      Set oParameter = .CreateParameter("@TotalPagosReembolso", adCurrency, adParamInput, 0, lnTotalPagos): .Parameters.Append oParameter
5210                                                End If
5220                                                .Execute
5230                                              End With
5240                                              Set oParameter = Nothing
5250                                              Set oCommand = Nothing
                                                  
                                                  '
5260                                              oRsTmpReemb1.MoveNext
5270                                        Loop
5280                                     End If
5290                                     oRsTmpReemb1.Close
5300                                  End If
5310                            End If
5320                            End If
5330                            rsTmp10.MoveNext
5340                         Loop
5350                      End If
5360                      rsTmp10.Close
5370                  End If
                      
                      'Actualiza Totales en la cuenta
5380                  With oCommand
5390                    .CommandType = adCmdStoredProc
5400                    Set .ActiveConnection = oConexion
5410                    .CommandTimeout = 150
5420                    .CommandText = "FacturacionCuentasAtencionActualizaImportes"
5430                    Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
5440                    Set oParameter = .CreateParameter("@TotalExonerado", adCurrency, adParamInput, 0, lnTotalExoneradoCta): .Parameters.Append oParameter
5450                    Set oParameter = .CreateParameter("@TotalAsegurado", adCurrency, adParamInput, 0, IIf(lbEsCuentaPagante = True, 0, lnTotalConsumosCta)): .Parameters.Append oParameter
5460                    Set oParameter = .CreateParameter("@TotalPagado", adCurrency, adParamInput, 0, IIf(lbEsCuentaPagante = True, lnTotalPagadoCta, lnTotalReembolsoCta)): .Parameters.Append oParameter
5470                    Set oParameter = .CreateParameter("@totalPorPagar", adCurrency, adParamInput, 0, IIf(lbEsCuentaPagante = True, lnTotalConsumosCta, (lnTotalConsumosCta - lnTotalReembolsoCta))): .Parameters.Append oParameter
5480                    .Execute
5490                  End With
5500                  Set oParameter = Nothing
5510                  Set oCommand = Nothing
                      
                      '
5520                  mrs_Tmp1.MoveNext
5530             Loop
5540      End If
5550      If lbTodasLasCuentas = True Then
5560         MsgBox "Nro cuentas procesadas: " & lnCuentasProcesadas, vbInformation, "Resultado"
5570      End If
          
5580      Set mrs_Tmp = Nothing
5590      Set lcBuscaParametro = Nothing
5600      Set oConexion = Nothing
5610      Set rsReporte = Nothing
5620      Set rsTmp = Nothing
5630      Set rsTmp10 = Nothing
5640      Set mrs_Tmp1 = Nothing
5650      Set mrs_Ctas = Nothing
5660      Set rsTmpReemb2 = Nothing
5670      Set oRsTmpReemb1 = Nothing
5680      Set oRsPuntosDeCarga = Nothing
5690      Set oDoImagMovimientoIngresos = Nothing
5700      Set oDOImagMovimiento = Nothing
          
5710      Exit Sub
ErCtaCta:
5720      MsgBox Err.Description & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                   "Sub FacturacionCuentasAtencionPtosActualizar", "reglasFacturacion.cls") 'debb-02/05/2016
          'Resume
5730      lcError = Err.Description


End Sub
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FarmaciaChequeaSiBoletaEsDeLaFarmacia(lnIdComprobantePago As Long, lnIdFarmacia As Long) As Boolean
    FarmaciaChequeaSiBoletaEsDeLaFarmacia = False
    Dim lcSql As String, oRsTmp1 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmaciaChequeaSiBoletaEsDeLaFarmacia"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   oConexion.Close
   
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   
    If oRsTmp1.RecordCount > 0 Then
       If oRsTmp1.Fields!idAlmacenOrigen = lnIdFarmacia Then
           FarmaciaChequeaSiBoletaEsDeLaFarmacia = True
       End If
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Function
'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarPorFechaIngresoMedicoSinAlta(ldFechaIni As Date, ldFechaFin As Date, lbSoloConsultaExterna As Boolean, oConexion As Connection, lnIdFuenteFinanciamiento As Long) As Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim SoloConsultaExterna As Integer
    Set AtencionesSeleccionarPorFechaIngresoMedicoSinAlta = Nothing

            If lbSoloConsultaExterna = True Then
                SoloConsultaExterna = 1
            Else
                SoloConsultaExterna = 2
            End If
            If lnIdFuenteFinanciamiento > 0 Then
            End If

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaIngresoMedicoSinAlta"
        Set oParameter = .CreateParameter("@FechaIni", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SoloConsultaExterna", adInteger, adParamInput, 0, SoloConsultaExterna): .Parameters.Append oParameter
        Set oRecordset = .Execute
       
   End With
   Set AtencionesSeleccionarPorFechaIngresoMedicoSinAlta = oRecordset
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarPorFechaIngresoMedicoYfuenteFinanciamientoSinAlta(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long, oConexion As Connection) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesSeleccionarPorFechaIngresoMedicoYfuenteFinanciamientoSinAlta"
        Set oParameter = .CreateParameter("@ldFechaIni", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set AtencionesSeleccionarPorFechaIngresoMedicoYfuenteFinanciamientoSinAlta = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function AtencionesSeleccionarPorFuenteFinanciamientoFingresoSinAlta1(ldFechaIni As Date, ldFechaFin As Date, lnIdFuenteFinanciamiento As Long, lnidEstadoCuenta As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String

    
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacAtencionesSeleccionarPorFuenteFinanciamientoFingresoSinAlta"
        Set oParameter = .CreateParameter("@ldFechaIni", adDBTimeStamp, adParamInput, 0, Format(ldFechaIni, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ldFechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy")): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idEstadoCuenta", adInteger, adParamInput, 0, lnidEstadoCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set AtencionesSeleccionarPorFuenteFinanciamientoFingresoSinAlta1 = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Set oRecordset = Nothing
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioPagosFiltraPorIdOrdenConexion(lnIdOrden As Long, oConexion As Connection) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioPagosFiltraPorIdOrdenConexion = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosFiltraPorIdOrdenConexion"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionServicioPagosFiltraPorIdOrdenConexion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
    
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioFinanciamientosFiltraPorIdOrdenConexion(lnIdOrden As Long, oConexion As Connection) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioFinanciamientosFiltraPorIdOrdenConexion = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosFiltraPorIdOrdenConexion"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set FacturacionServicioFinanciamientosFiltraPorIdOrdenConexion = oRecordset
   Set oRecordset = Nothing
   Set oCommand = Nothing
    
End Function

Function DevuelveDxIngresoSoloHospEmerg(ml_IdTipoServicio As Long, ml_idAtencion As Long, oConexion As Connection) As String
    Dim rsDiagnosticos As New Recordset
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    DevuelveDxIngresoSoloHospEmerg = ""
    If ml_IdTipoServicio <> sghConsultaExterna Then
       Set rsDiagnosticos = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarPorAtencion(ml_idAtencion, sghHospitalizacionIngreso, oConexion)
       If rsDiagnosticos.RecordCount > 0 Then
          DevuelveDxIngresoSoloHospEmerg = "(" & Trim(rsDiagnosticos!CodigoCIE2004) & ") " & rsDiagnosticos!Descripcion
       End If
       rsDiagnosticos.Close
    End If
    Set rsDiagnosticos = Nothing
    Set mo_AdminAdmision = Nothing
End Function

Function DevuelveDxAltaMedicaConexion(lnIdAtencion As Long, lnIdTipoServicio As Long, oConexion As Connection) As doDiagnostico
            Dim mo_AdminAdmision1 As New SIGHNegocios.ReglasAdmision
            Dim oRsTmp As New Recordset
            Dim oDODiagnostico As New doDiagnostico
            If lnIdTipoServicio = 1 Then
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 1, oConexion)
            Else
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 3, oConexion)
            End If
            oDODiagnostico.IdDiagnostico = 0
            oDODiagnostico.Descripcion = ""
            oDODiagnostico.CodigoCIE2004 = ""
            If lnIdTipoServicio = 1 Then
               oRsTmp.Filter = "IdTipoDiagnostico=102"    'CE, vivos
            Else
               oRsTmp.Filter = "IdTipoDiagnostico=301"    'Hosp/emerg, vivos
            End If
            If oRsTmp.RecordCount > 0 Then
               oDODiagnostico.CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
               oDODiagnostico.Descripcion = oRsTmp.Fields!Descripcion
               oDODiagnostico.IdDiagnostico = oRsTmp.Fields!IdDiagnostico
            ElseIf lnIdTipoServicio <> 1 Then
               Set oRsTmp = mo_AdminAdmision1.AtencionesDiagnosticosSeleccionarPorAtencion(lnIdAtencion, 4, oConexion)
               oRsTmp.Filter = "IdTipoDiagnostico=303"    'Hosp/emerg, fallecidos
               If oRsTmp.RecordCount > 0 Then
                    oDODiagnostico.CodigoCIE2004 = oRsTmp.Fields!CodigoCIE2004
                    oDODiagnostico.Descripcion = oRsTmp.Fields!Descripcion
                    oDODiagnostico.IdDiagnostico = oRsTmp.Fields!IdDiagnostico
               End If
            End If
            Set oRsTmp = Nothing
            Set DevuelveDxAltaMedicaConexion = oDODiagnostico
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioDespachoSeleccionarPorIdCuenta(lnIdCuentaAtencion As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim lcSql As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set FacturacionServicioDespachoSeleccionarPorIdCuenta = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioDespachoSeleccionarPorIdCuenta"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
       End With
       Set FacturacionServicioDespachoSeleccionarPorIdCuenta = oRecordset
       oConexion.Close
       Set oRecordset = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionServicioPagosPorFechasComprobante(ldFechaInicio As Date, ldFechaFin As Date) As Recordset
         Dim oRsCitaPagada1 As New Recordset
         Dim lcSql As String
         Dim oCommand As New ADODB.Command
         Dim oParameter As ADODB.Parameter
         Dim oConexion As New ADODB.Connection

         oConexion.CursorLocation = adUseClient
         oConexion.CommandTimeout = 300
         oConexion.Open sighentidades.CadenaConexion
         Set FacturacionServicioPagosPorFechasComprobante = Nothing
         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "FacturacionServicioPagosPorFechasComprobante"
             Set oParameter = .CreateParameter("@ldFechaIni", adDBTimeStamp, adParamInput, 0, Format(ldFechaInicio, "dd/mm/yyyy") & " 00:00:01"): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@ldFechaFin", adDBTimeStamp, adParamInput, 0, Format(ldFechaFin, "dd/mm/yyyy") & " 23:59:59"): .Parameters.Append oParameter
             Set oRsCitaPagada1 = .Execute
             Set oRsCitaPagada1.ActiveConnection = Nothing
        End With
        Set FacturacionServicioPagosPorFechasComprobante = oRsCitaPagada1
        oConexion.Close
        Set oParameter = Nothing
        Set oConexion = Nothing
        Set oRsCitaPagada1 = Nothing
        Set oCommand = Nothing
         
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function RetornaTotalConsumoServiciosParaPagantePorNroCuenta(lnNroCuenta As Long) As Double
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Dim lnIdOrdenPago As Long
    Dim lcSql As String

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RetornaTotalConsumoServiciosParaPagantePorNroCuenta"
            Set oParameter = .CreateParameter("@NroCuenta", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set oCommand = Nothing
       Set oParameter = Nothing

    lnTotal = 0
     If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnIdOrdenPago = oRsTmp.Fields!idOrdenPago
            If oRsTmp.Fields!ImporteExonerado > 0 Then
               lnTotal = lnTotal - oRsTmp.Fields!ImporteExonerado
            End If
            Do While Not oRsTmp.EOF And lnIdOrdenPago = oRsTmp.Fields!idOrdenPago
               lnTotal = lnTotal + oRsTmp.Fields!totalPorPagar
               oRsTmp.MoveNext
               If oRsTmp.EOF Then
                  Exit Do
               End If
            Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    RetornaTotalConsumoServiciosParaPagantePorNroCuenta = lnTotal
End Function
'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta(lnNroCuenta As Long) As Double
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim oRsTmp As New Recordset
    Dim lnTotal As Double
    Dim lnIdOrden As Long
    Dim lcSql As String

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta"
            Set oParameter = .CreateParameter("@NroCuenta", adInteger, adParamInput, 0, lnNroCuenta): .Parameters.Append oParameter
            Set oRsTmp = .Execute
            Set oRsTmp.ActiveConnection = Nothing
       End With
       Set oCommand = Nothing
       Set oParameter = Nothing
    
    oRsTmp.Filter = "movNumero<>null"
    lnTotal = 0
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
            lnIdOrden = oRsTmp.Fields!IdOrden
            If oRsTmp.Fields!ImporteExonerado > 0 Then
               lnTotal = lnTotal - oRsTmp.Fields!ImporteExonerado
            End If
            Do While Not oRsTmp.EOF And lnIdOrden = oRsTmp.Fields!IdOrden
               lnTotal = lnTotal + oRsTmp.Fields!totalPorPagar
               oRsTmp.MoveNext
               If oRsTmp.EOF Then
                  Exit Do
               End If
            Loop
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set mo_ReglasFarmacia = Nothing
    RetornaTotalConsumoFarmaciaParaPagantePorNroCuenta = lnTotal
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function FactOrdenServicioSeleccionarXidOrden(lnIdOrden As Long) As Recordset
         Dim oRsCitaPagada1 As New Recordset
         Dim lcSql As String

        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
        Set FactOrdenServicioSeleccionarXidOrden = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactOrdenServicioSeleccionarXidOrden"
                Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
                Set oRsCitaPagada1 = .Execute
                Set oRsCitaPagada1.ActiveConnection = Nothing
           End With
           Set FactOrdenServicioSeleccionarXidOrden = oRsCitaPagada1
           oConexion.Close
           Set oRsCitaPagada1 = Nothing
           Set oConexion = Nothing
           Set oCommand = Nothing
           Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function FactReembolsosSelecionarPorAreaTramitaSegurosDEBB(lnIdAreaTramitaSeguro As Long, lnIdCuentaAtencion As Long, _
                                           lbSoloImporteReembolsadoDifConsumido As Boolean) As ADODB.Recordset
     'Dim lcSql As String
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oParameter As ADODB.Parameter
     Dim oConexion As New ADODB.Connection
     
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
     If lnIdCuentaAtencion > 0 Then

        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FactReembolsosSelecionarPorAreaTramitaSeguroYCuentaAtencion"
          Set oParameter = .CreateParameter("@IdAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
          Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
  
      ElseIf lnIdCuentaAtencion = 0 And lbSoloImporteReembolsadoDifConsumido = False Then
        With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "FactReembolsosSelecionarPorAreaTramitaSeguro"
          Set oParameter = .CreateParameter("@IdAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
          Set oRecordset = .Execute
          Set oRecordset.ActiveConnection = Nothing
        End With
        
      ElseIf lbSoloImporteReembolsadoDifConsumido = True Then
          Dim oRsTmp1 As New Recordset
          Dim lnIdFactReembolso As Long
          Dim lnAnio As Long, lnMes As Long, lcPlan As String, lcNroSerie As String, lcNroDocumento As String
          Dim lcDescripcion As String, lnConsumoPorReembolsar As Double, lnReembolsoPagado As Double, lcEstado As String
          Dim lnIdEstadoReembolso As Long, lbGrabaDefinitivamente As Boolean, lnTotalServFarmacia As Double
          With oRecordset
                .Fields.Append "idFactReembolso", adInteger, 4, adFldIsNullable
                .Fields.Append "Anio", adInteger, 4, adFldIsNullable
                .Fields.Append "Mes", adInteger, 4, adFldIsNullable
                .Fields.Append "Plan", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroSerie", adVarChar, 100, adFldIsNullable
                .Fields.Append "NroDocumento", adVarChar, 100, adFldIsNullable
                .Fields.Append "Descripcion", adVarChar, 100, adFldIsNullable
                .Fields.Append "ConsumoPorReembolsar", adDouble
                .Fields.Append "ReembolsoPagado", adDouble
                .Fields.Append "Estado", adVarChar, 100, adFldIsNullable
                .Fields.Append "idEstadoReembolso", adInteger, 4, adFldIsNullable
                .Fields.Append "GrabaDefinitivamente", adBoolean
                .LockType = adLockOptimistic
                .Open
          End With
            With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "FactReembolsosSelecionarPorIdAreaTramitaSeguro"
              Set oParameter = .CreateParameter("@IdAreaTramitaSeguro", adInteger, adParamInput, 0, lnIdAreaTramitaSeguro): .Parameters.Append oParameter
              Set oRsTmp1 = .Execute
              Set oRsTmp1.ActiveConnection = Nothing
            End With
          
          If oRsTmp1.RecordCount > 0 Then
             oRsTmp1.MoveFirst
             Do While Not oRsTmp1.EOF
                lnIdFactReembolso = oRsTmp1.Fields!idFactReembolso
                lnAnio = oRsTmp1.Fields!Anio
                lnMes = oRsTmp1.Fields!Mes
                lcPlan = oRsTmp1.Fields!Plan
                lcNroSerie = IIf(IsNull(oRsTmp1.Fields!NroSerie), "", oRsTmp1.Fields!NroSerie)
                lcNroDocumento = IIf(IsNull(oRsTmp1.Fields!NroDocumento), "", oRsTmp1.Fields!NroDocumento)
                lcDescripcion = oRsTmp1.Fields!Descripcion
                lnConsumoPorReembolsar = oRsTmp1.Fields!ConsumoPorReembolsar
                lnReembolsoPagado = oRsTmp1.Fields!ReembolsoPagado
                lcEstado = oRsTmp1.Fields!Estado
                lnIdEstadoReembolso = oRsTmp1.Fields!IdEstadoReembolso
                lbGrabaDefinitivamente = oRsTmp1.Fields!GrabaDefinitivamente
                lnTotalServFarmacia = 0
                Do While Not oRsTmp1.EOF And lnIdFactReembolso = oRsTmp1.Fields!idFactReembolso
                   lnTotalServFarmacia = lnTotalServFarmacia + _
                                         IIf(IsNull(oRsTmp1.Fields!ReembPagadoServ), 0, oRsTmp1.Fields!ReembPagadoServ) + _
                                         IIf(IsNull(oRsTmp1.Fields!ReembPagadoFarm), 0, oRsTmp1.Fields!ReembPagadoFarm)
                   oRsTmp1.MoveNext
                   If oRsTmp1.EOF Then
                      Exit Do
                   End If
                Loop
                If Round(lnTotalServFarmacia, 0) <> Round(lnReembolsoPagado, 0) Then
                    oRecordset.AddNew
                    oRecordset.Fields!idFactReembolso = lnIdFactReembolso
                    oRecordset.Fields!Anio = lnAnio
                    oRecordset.Fields!Mes = lnMes
                    oRecordset.Fields!Plan = lcPlan
                    oRecordset.Fields!NroSerie = lcNroSerie
                    oRecordset.Fields!NroDocumento = lcNroDocumento
                    oRecordset.Fields!Descripcion = lcDescripcion
                    oRecordset.Fields!ConsumoPorReembolsar = lnConsumoPorReembolsar
                    oRecordset.Fields!ReembolsoPagado = lnReembolsoPagado
                    oRecordset.Fields!Estado = lcEstado
                    oRecordset.Fields!IdEstadoReembolso = lnIdEstadoReembolso
                    oRecordset.Fields!GrabaDefinitivamente = lbGrabaDefinitivamente
                    oRecordset.Update
                End If
             Loop
          End If
          oRsTmp1.Close
          Set oRsTmp1 = Nothing
      End If
      Set FactReembolsosSelecionarPorAreaTramitaSegurosDEBB = oRecordset
      Set oCommand = Nothing
        
      Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function FacturacionServicioDespachoSeleccionarPorIdProducto(lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oFacturacionServicioDespacho As New FacturacionServicioDespacho
Dim oConexion As New Connection
   oConexion.CommandTimeout = 300
   oConexion.CursorLocation = adUseClient
   oConexion.Open sighentidades.CadenaConexion
   Set oFacturacionServicioDespacho.Conexion = oConexion
   Set FacturacionServicioDespachoSeleccionarPorIdProducto = oFacturacionServicioDespacho.SeleccionarPorIdProducto(lnIdProducto)
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc(lnIdOrden As Long, lnIdProducto As Long, _
                                                            lnIdtipoFinanciamiento As Long, oConexion As Connection) As ADODB.Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc"
            Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioFinanciamFiltrarXIdOrdenProductoTFinanc = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 01-10-13 Se cambio a Store Procedure
Function FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc(lcMovNumero As String, lnIdProducto As Long, _
                                                            lnIdtipoFinanciamiento As Long, lcMovTipo As String, _
                                                            oConexion As Connection) As ADODB.Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc"
            Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 50, lcMovNumero): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 50, lcMovTipo): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionBienesFinanciamFiltrarXmovNumeroProductoTFinanc = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function FacturacionCatalogoPaqueteFarmSeleccionarXid(lnIdPaquete As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set FacturacionCatalogoPaqueteFarmSeleccionarXid = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FacturacionCatalogoPaqueteFarmSeleccionarXid"
                Set oParameter = .CreateParameter("@lnIdPaquete", adInteger, adParamInput, 0, lnIdPaquete): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, 5): .Parameters.Append oParameter
                Set oRsTmp2 = .Execute
                Set oRsTmp2.ActiveConnection = Nothing
           End With
           Set FacturacionCatalogoPaqueteFarmSeleccionarXid = oRsTmp2
           oConexion.Close
           Set oRsTmp2 = Nothing
           Set oConexion = Nothing
           Set oCommand = Nothing
           Set oParameter = Nothing
End Function


Function TiposFinanciamientoSoloFarmacia() As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

        Dim oCommand As New ADODB.Command
        Dim oParameter As ADODB.Parameter
        Dim oConexion As New ADODB.Connection
            Set TiposFinanciamientoSoloFarmacia = Nothing
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "TiposFinanciamientoSoloFarmacia"
                Set oRsTmp2 = .Execute
                Set oRsTmp2.ActiveConnection = Nothing
           End With
           Set TiposFinanciamientoSoloFarmacia = oRsTmp2
           oConexion.Close
           Set oRsTmp2 = Nothing
           Set oConexion = Nothing
           Set oCommand = Nothing
           Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 04-10-13 Se cambio a Store Procedure
Function TiposFinanciamientoXidentificador(lnIdtipoFinanciamiento As Long, Optional oConexion1 As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
   
     
         If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
     
        With oCommand
             .CommandType = adCmdStoredProc
             If oConexion1 Is Nothing Then
                Set .ActiveConnection = oConexion
             Else
                Set .ActiveConnection = oConexion1
             End If
             .CommandTimeout = 150
             .CommandText = "TiposFinanciamientoXidentificador"
             Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
             Set oRsTmp2 = .Execute
        End With
        Set TiposFinanciamientoXidentificador = oRsTmp2
        Set oCommand = Nothing
        Set oParameter = Nothing
        Set oConexion = Nothing


End Function

Function FactReembolsosSeleccionarXidentificador(lnIdFactReembolso As Long, Optional oConexion1 As Connection) As Recordset
'    Dim ssql As String
    Dim oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Set FactReembolsosSeleccionarXidentificador = Nothing
    
         If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
    
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactReembolsosSeleccionarXidentificador"
        Set oParameter = .CreateParameter("@IdFactReembolso", adInteger, adParamInput, 0, lnIdFactReembolso): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
   End With
   Set FactReembolsosSeleccionarXidentificador = oRsTmp2
   Set oCommand = Nothing
   Set oParameter = Nothing
   Set oConexion = Nothing
    
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function FactOrdenServicioPagosXImporteExoneradoMayorAcero(lnIdEmpleado As Long, mda_FechaInicio As Date, mda_FechaFin As Date) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
  
    Set FactOrdenServicioPagosXImporteExoneradoMayorAcero = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenServicioPagosXImporteExoneradoMayorAcero"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(mda_FechaInicio, "dd/mm/yyyy") & " 00:00:01"): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(mda_FechaFin, "dd/mm/yyyy") & " 23:59:59"): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
   End With
   Set FactOrdenServicioPagosXImporteExoneradoMayorAcero = oRsTmp2
   Set oRsTmp2 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function


'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function FactOrdenesBienesXimporteExoneradoMayorAcero(lnIdEmpleado As Long, mda_FechaInicio As Date, mda_FechaFin As Date) As Recordset
'    Dim ssql As String
    Dim oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactOrdenesBienesXimporteExoneradoMayorAcero"
        Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, lnIdEmpleado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, Format(mda_FechaInicio, "dd/mm/yyyy") & " 00:00:01"): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, Format(mda_FechaFin, "dd/mm/yyyy") & " 23:59:59"): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
    End With
    
    Set FactOrdenesBienesXimporteExoneradoMayorAcero = oRsTmp2
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionBienesFinanciamientosXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionBienesFinanciamientosXcuenta = Nothing
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionBienesFinanciamientosXcuenta"
            Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionBienesFinanciamientosXcuenta = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
  
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioFinanciamientosXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioFinanciamientosXcuenta = Nothing
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioFinanciamientosXcuenta"
            Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioFinanciamientosXcuenta = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionBienesFinanciamientosXcuentaYconexion(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionBienesFinanciamientosXcuentaYconexion = Nothing
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionBienesFinanciamientosXcuentaYconexion"
            Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionBienesFinanciamientosXcuentaYconexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioFinanciamientosXidOrdenIdProducto(lnIdOrden As Long, lnIdProducto As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioFinanciamientosXidOrdenIdProducto = Nothing
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioFinanciamientosXidOrdenIdProducto"
            Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioFinanciamientosXidOrdenIdProducto = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionCuentasAtencionPtosXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionCuentasAtencionPtosXcuenta = Nothing
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionCuentasAtencionPtosXcuenta"
            Set oParameter = .CreateParameter("@idCuenta", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionCuentasAtencionPtosXcuenta = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Sub CatalogoServiciosHospEliminarXtipoFinanciamientoIdProducto(lnIdtipoFinanciamiento As Long, lnIdProducto As Long, oConexion As Connection)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
  
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandTimeout = 150
      .CommandText = "FactCatalogoServiciosHospEliminarXtipoFinanciamientoIdProducto"
      Set oParameter = .CreateParameter("@IdtipoFinanciamiento", adInteger, adParamInput, 0, lnIdtipoFinanciamiento): .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
      .Execute
    End With
    Set oCommand = Nothing
End Sub

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FuentesFinanciamientoSoloParticular() As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    
    Dim oCommand As New ADODB.Command
    Dim oConexion As New ADODB.Connection
        Set FuentesFinanciamientoSoloParticular = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FuentesFinanciamientoSoloParticular"
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
       End With
       Set FuentesFinanciamientoSoloParticular = oRsTmp2
       oConexion.Close
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
    
End Function

Function FacturacionPaquetesXidComprobantePago(lnIdComprobantePago As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FacturacionPaquetesXidComprobantePago = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionPaquetesXidComprobantePago"
        Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
   End With
   Set FacturacionPaquetesXidComprobantePago = oRsTmp2
   oConexion.Close
   Set oRsTmp2 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
   
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioDespachoXcuenta(ml_idCuentaAtencion As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        Set FacturacionServicioDespachoXcuenta = Nothing
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioDespachoXcuenta"
            Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
       End With
       Set FacturacionServicioDespachoXcuenta = oRsTmp2
       oConexion.Close
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

Function FacturacionServicioFinanciamPorIdOrdenIdProducto(mo_conexion As Connection, lnIdOrden, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionServicioFinanciamPorIdOrdenIdProducto = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosSeleccionarPorIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioFinanciamPorIdOrdenIdProducto = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

Function FacturacionServicioPagosPorIdOrdenIdProducto(mo_conexion As Connection, lnIdOrden As Long, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionServicioPagosPorIdOrdenIdProducto = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioPagosSeleccionarPorIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioPagosPorIdOrdenIdProducto = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesPagosPorIdProducto(mo_conexion As Connection, lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionBienesPagosPorIdProducto = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesPagosSeleccionarPorMovnumeroIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesPagosPorIdProducto = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioFinanciamientosXcuentaConexion(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioFinanciamientosXcuentaConexion = Nothing
      With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioFinanciamientosXcuentaConexion"
            Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioFinanciamientosXcuentaConexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioPagosXcuentaYconexion(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioPagosXcuentaYconexion = Nothing
      With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioPagosXcuentaYconexion"
            Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioPagosXcuentaYconexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioDevolucionesXcuenta(ml_idCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioDevolucionesXcuenta = Nothing
      With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioDevolucionesXcuenta"
            Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_idCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioDevolucionesXcuenta = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
    
End Function

Function FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProductoConexion(mo_conexion As Connection, lnIdOrden As Long, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProductoConexion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProducto"
        Set oParameter = .CreateParameter("@idOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionServicioDevolucionesSeleccionarPorIdOrdenIdProductoConexion = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FacturacionBienesDevolucionesSeleccionarPorIdProductoConexion(mo_conexion As Connection, lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FacturacionBienesDevolucionesSeleccionarPorIdProductoConexion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_conexion
        .CommandTimeout = 150
        .CommandText = "FacturacionBienesDevolucionesSeleccionarPorIdProducto"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   Set FacturacionBienesDevolucionesSeleccionarPorIdProductoConexion = oRecordset
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionCatalogoPaquetesXidPaquete(lnIdPaquete As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    Set FacturacionCatalogoPaquetesXidPaquete = Nothing
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCatalogoPaquetesXidPaquete"
        Set oParameter = .CreateParameter("@IdPaquete", adInteger, adParamInput, 0, lnIdPaquete): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
   End With
   Set FacturacionCatalogoPaquetesXidPaquete = oRsTmp2
   oConexion.Close
   Set oRsTmp2 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
    
End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Function facturacionBienesFinanciamientosSoloExonerados(lcMovNumero As String, lcMovTipo As String, lnIdProducto As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "facturacionBienesFinanciamientosSoloExonerados"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 50, lcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 50, lcMovTipo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRsTmp2 = .Execute
        Set oRsTmp2.ActiveConnection = Nothing
   End With
   Set facturacionBienesFinanciamientosSoloExonerados = oRsTmp2
   oConexion.Close
   Set oRsTmp2 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function facturacionServicioFinanciamientosSoloExonerados(lnIdOrden As Long, lnIdProducto As Long) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
        
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "facturacionServicioFinanciamientosSoloExonerados"
            Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
            Set oRsTmp2.ActiveConnection = Nothing
       End With
       Set facturacionServicioFinanciamientosSoloExonerados = oRsTmp2
       oConexion.Close
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionReembolsosXboletaConexion(lnIdComprobantePago As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionReembolsosXboletaConexion"
            Set oParameter = .CreateParameter("@IdComprobantePago ", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionReembolsosXboletaConexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionReembolsosXcuenta(ml_idCuenta As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionReembolsosXcuenta = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionReembolsosXcuenta"
            Set oParameter = .CreateParameter("@idCuenta ", adInteger, adParamInput, 0, ml_idCuenta): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionReembolsosXcuenta = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    
End Function

Function FactOrdenServicioXidOrdenConexion(lnIdOrden As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactOrdenServicioXidOrdenConexion"
            Set oParameter = .CreateParameter("@IdOrden ", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FactOrdenServicioXidOrdenConexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FactOrdenServicioPoridOrdenConexion(lnIdOrden As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FactOrdenServicioPoridOrdenConexion = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FactOrdenServicioPoridOrdenConexion"
            Set oParameter = .CreateParameter("@IdOrden ", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set oCommand = Nothing
       Set oParameter = Nothing
       Set FactOrdenServicioPoridOrdenConexion = oRsTmp2
End Function
'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionServicioPagosXidComprobantePagoConexion(lnIdComprobantePago As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionServicioPagosXidComprobantePagoConexion = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionServicioPagosXidComprobantePagoConexion"
            Set oParameter = .CreateParameter("@IdComprobantePago ", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionServicioPagosXidComprobantePagoConexion = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing

End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FactCatalogoPaqueteSeleccionarTodos() As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

        Dim oCommand As New ADODB.Command
        Dim oConexion As New ADODB.Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
            With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "FactCatalogoPaqueteSeleccionarTodos"
                Set oRsTmp2 = .Execute
                Set oRsTmp2.ActiveConnection = Nothing
           End With
           Set FactCatalogoPaqueteSeleccionarTodos = oRsTmp2
           oConexion.Close
           Set oRsTmp2 = Nothing
           Set oConexion = Nothing
           Set oCommand = Nothing
   
End Function

'Modificado por Yamill Palomino 02-10-13 Se cambio a Store Procedure
Function FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad(lnIdCuentaAtencion As Long, lnIdEspecialidad As Long, lnIdPuntoCarga As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset

    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Set FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad = Nothing
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad"
            Set oParameter = .CreateParameter("@idPuntoCarga ", adInteger, adParamInput, 0, lnIdPuntoCarga): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdCuentaAtencion ", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdEspecialidad ", adInteger, adParamInput, 0, lnIdEspecialidad): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad = oRsTmp2
       Set oRsTmp2 = Nothing
       Set oCommand = Nothing
       Set oParameter = Nothing
    End Function

'Modificado por Yamill Palomino 06-10-13 Se cambio a Store Procedure
Sub DesagregaCuentasItemsXreembolso(ByVal oRsItemsXCuenta1 As Recordset, lnIdComprobantePago As Long, _
                                    lnTotalIdProducto As Double, moConexion As Connection)
    Dim lcSql As String
    Dim oRsItems As New Recordset
    Dim oRsCuentas As New Recordset
    Dim lbNuevo As Boolean
    Dim lnImpTotal As Double, lnImpProrrateo As Double, lnTotal1 As Double
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    On Error GoTo ErrCtas
    If oRsItemsXCuenta1.State = adStateOpen Then
         oRsItemsXCuenta1.Close
    End If
    With oRsItemsXCuenta1
       .Fields.Append "IdPartida", adInteger
       .Fields.Append "IdProducto", adInteger
       .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
       .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
       .Fields.Append "Precio", adDouble
       .Fields.Append "Cantidad", adInteger
       .Fields.Append "Importe", adDouble
       .Fields.Append "ImporteProrrateado", adDouble
       .LockType = adLockOptimistic
       .Open
    End With
    '

    With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = moConexion
          .CommandTimeout = 150
          .CommandText = "FactReembolsosDocumentosXIdComprobantePago"
          Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, lnIdComprobantePago): .Parameters.Append oParameter
          Set oRsCuentas = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    lnImpTotal = 0
    If oRsCuentas.RecordCount > 0 Then
       oRsCuentas.MoveFirst
       Do While Not oRsCuentas.EOF

            With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = moConexion
                 .CommandTimeout = 150
                 .CommandText = "FactCatalogoServiciosXIdCuentaAtencionIdFuenteFinanciamiento"
                 Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, oRsCuentas.Fields!IdCuentaAtencion): .Parameters.Append oParameter
                 Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 0, oRsCuentas.Fields!idFuenteFinanciamiento): .Parameters.Append oParameter
                 Set oRsItems = .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
                  
            If oRsItems.RecordCount > 0 Then
               oRsItems.MoveFirst
               Do While Not oRsItems.EOF
                  lbNuevo = True
                  If oRsItemsXCuenta1.RecordCount > 0 Then
                     oRsItemsXCuenta1.MoveFirst
                     oRsItemsXCuenta1.Find "idProducto=" & oRsItems.Fields!idProducto
                     If Not oRsItemsXCuenta1.EOF Then
                        lbNuevo = False
                     End If
                  End If
                  If lbNuevo = True Then
                     oRsItemsXCuenta1.AddNew
                     oRsItemsXCuenta1.Fields!IdPartida = oRsItems.Fields!IdPartida
                     oRsItemsXCuenta1.Fields!idProducto = oRsItems.Fields!idProducto
                     oRsItemsXCuenta1.Fields!Codigo = oRsItems.Fields!Codigo
                     oRsItemsXCuenta1.Fields!nombreProducto = Left(oRsItems.Fields!nombre, 150)
                     oRsItemsXCuenta1.Fields!Precio = oRsItems.Fields!PrecioFinanciado
                  End If
                  oRsItemsXCuenta1.Fields!cantidad = oRsItemsXCuenta1.Fields!cantidad + oRsItems.Fields!CantidadFinanciada
                  oRsItemsXCuenta1.Fields!Importe = oRsItemsXCuenta1.Fields!Importe + oRsItems.Fields!TotalFinanciado
                  oRsItemsXCuenta1.Update
                  lnImpTotal = lnImpTotal + oRsItems.Fields!TotalFinanciado
                  oRsItems.MoveNext
               Loop
            End If
            oRsItems.Close
            oRsCuentas.MoveNext
       Loop
    End If
    oRsCuentas.Close
    '***proceso de prorrateo
    If oRsItemsXCuenta1.RecordCount > 0 And lnImpTotal <> 0 Then
       lnTotal1 = 0
       oRsItemsXCuenta1.MoveFirst
       Do While Not oRsItemsXCuenta1.EOF
          lnImpProrrateo = Round((lnTotalIdProducto * oRsItemsXCuenta1.Fields!Importe) / lnImpTotal, 2)
          lnTotal1 = lnTotal1 + lnImpProrrateo
          oRsItemsXCuenta1.Fields!ImporteProrrateado = lnImpProrrateo
          oRsItemsXCuenta1.Update
          oRsItemsXCuenta1.MoveNext
       Loop
       If lnTotal1 <> lnTotalIdProducto Then
          lnImpProrrateo = lnTotalIdProducto - lnTotal1
          oRsItemsXCuenta1.MoveFirst
          oRsItemsXCuenta1.Fields!ImporteProrrateado = oRsItemsXCuenta1.Fields!ImporteProrrateado + lnImpProrrateo
          oRsItemsXCuenta1.Update
       End If
    End If
    Exit Sub
ErrCtas:
    MsgBox Err.Description
    Resume
End Sub

Sub DesagregaCuentasItemsXServHosp(ByVal oRsItemsXCuenta1 As Recordset, lnIdCuentaAtencion As Long, _
                                   lnTotalIdProducto As Double, moConexion As Connection, lnIdServiciosHospitalarios As Long)
'    Dim lcSql As String
    Dim oRsItems As New Recordset
    Dim oRsCuentas As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lbNuevo As Boolean
    Dim lnImpTotal As Double, lnImpProrrateo As Double, lnTotal1 As Double, lnConsumoServicio As Double
    On Error GoTo ErrCtas1
    If oRsItemsXCuenta1.State = adStateOpen Then
         oRsItemsXCuenta1.Close
    End If
    With oRsItemsXCuenta1
       .Fields.Append "IdPartida", adInteger
       .Fields.Append "IdProducto", adInteger
       .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
       .Fields.Append "NombreProducto", adVarChar, 255, adFldIsNullable
       .Fields.Append "Precio", adDouble
       .Fields.Append "Cantidad", adInteger
       .Fields.Append "Importe", adDouble
       .Fields.Append "ImporteProrrateado", adDouble
       .LockType = adLockOptimistic
       .Open
    End With
    'Farmacia
    lnImpTotal = 0

     With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = moConexion
       .CommandTimeout = 150
       .CommandText = "FacDesagregaCuentasItemsXServHosp"
       Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
       Set oRsItems = .Execute
     End With
     Set oCommand = Nothing
     Set oParameter = Nothing
    
    If oRsItems.RecordCount > 0 Then
       oRsItems.MoveFirst
       Do While Not oRsItems.EOF
            oRsItemsXCuenta1.AddNew
            oRsItemsXCuenta1.Fields!IdPartida = oRsItems.Fields!IdPartida
            oRsItemsXCuenta1.Fields!idProducto = oRsItems.Fields!idProducto
            oRsItemsXCuenta1.Fields!Codigo = oRsItems.Fields!Codigo
            oRsItemsXCuenta1.Fields!nombreProducto = Left(oRsItems.Fields!nombre, 150)
            oRsItemsXCuenta1.Fields!Precio = oRsItems.Fields!Precio
            oRsItemsXCuenta1.Fields!cantidad = oRsItemsXCuenta1.Fields!cantidad + oRsItems.Fields!cantidad
            oRsItemsXCuenta1.Fields!Importe = oRsItemsXCuenta1.Fields!Importe + oRsItems.Fields!total
            oRsItemsXCuenta1.Update
            lnImpTotal = lnImpTotal + oRsItems.Fields!total
            oRsItems.MoveNext
       Loop
    End If
    oRsItems.Close
    'ServicioHospitalario

    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = moConexion
       .CommandTimeout = 150
       .CommandText = "FactCatalogoServiciosXIdProducto"
       Set oParameter = .CreateParameter("@IdServiciosHospitalarios", adInteger, adParamInput, 0, lnIdServiciosHospitalarios): .Parameters.Append oParameter
       Set oRsItems = .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
    
    If oRsItems.RecordCount > 0 Then
        lnConsumoServicio = RetornaConsumoPacienteServiciosConSeguroPorNroCuenta(lnIdCuentaAtencion)
        oRsItemsXCuenta1.AddNew
        oRsItemsXCuenta1.Fields!IdPartida = oRsItems.Fields!IdPartida
        oRsItemsXCuenta1.Fields!idProducto = oRsItems.Fields!idProducto
        oRsItemsXCuenta1.Fields!Codigo = oRsItems.Fields!Codigo
        oRsItemsXCuenta1.Fields!nombreProducto = oRsItems.Fields!nombre
        oRsItemsXCuenta1.Fields!Precio = 1
        oRsItemsXCuenta1.Fields!cantidad = 1
        oRsItemsXCuenta1.Fields!Importe = oRsItemsXCuenta1.Fields!Importe + lnConsumoServicio
        oRsItemsXCuenta1.Update
        lnImpTotal = lnImpTotal + lnConsumoServicio
    End If
    oRsItems.Close
    '***proceso de prorrateo
    If oRsItemsXCuenta1.RecordCount > 0 Then
       lnTotal1 = 0
       oRsItemsXCuenta1.MoveFirst
       Do While Not oRsItemsXCuenta1.EOF
If oRsItemsXCuenta1.Fields!idProducto = lnIdServiciosHospitalarios Then
lnImpProrrateo = 0
End If
          If oRsItemsXCuenta1.Fields!Importe > 0 Then
            lnImpProrrateo = Round((lnTotalIdProducto * oRsItemsXCuenta1.Fields!Importe) / lnImpTotal, 2)
            lnTotal1 = lnTotal1 + lnImpProrrateo
            oRsItemsXCuenta1.Fields!ImporteProrrateado = lnImpProrrateo
            oRsItemsXCuenta1.Update
          End If
          oRsItemsXCuenta1.MoveNext
       Loop
       If lnTotal1 <> lnTotalIdProducto Then
          lnImpProrrateo = lnTotalIdProducto - lnTotal1
          oRsItemsXCuenta1.MoveFirst
          oRsItemsXCuenta1.Fields!ImporteProrrateado = oRsItemsXCuenta1.Fields!ImporteProrrateado + lnImpProrrateo
          oRsItemsXCuenta1.Update
       End If
    End If
    Exit Sub
ErrCtas1:
    MsgBox Err.Description
    
End Sub



Function FacturacionCuentasAtencionPtosXIdCuenta(lnIdCuentaAtencion As Long, oConexion As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionCuentasAtencionPtosXIdCuenta"
            Set oParameter = .CreateParameter("@IdCuentaAtencion ", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
       End With
       Set FacturacionCuentasAtencionPtosXIdCuenta = oRsTmp2
       Set oCommand = Nothing
       Set oParameter = Nothing
    End Function




Function FactCatalogoServiciosConPrecioMayor() As Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset

     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
     Set FactCatalogoServiciosConPrecioMayor = Nothing
     oConexion.CursorLocation = adUseClient
     oConexion.CommandTimeout = 300
     oConexion.Open sighentidades.CadenaConexion
     With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "FactCatalogoServiciosConPrecioMayor"
         Set oRsTmp1 = .Execute
         Set oRsTmp1.ActiveConnection = Nothing
    End With
    Set FactCatalogoServiciosConPrecioMayor = oRsTmp1
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing

End Function

'yamill
Function FacturacionCuentasAtencionExoneracion(lnIdCuentaAtencion As Long, Optional oConexion1 As Connection) As Recordset
    Dim sSql As String, oRsTmp2 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
        If oConexion1 Is Nothing Then
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
        End If
        With oCommand
            .CommandType = adCmdStoredProc
            If oConexion1 Is Nothing Then
               Set .ActiveConnection = oConexion
            Else
               Set .ActiveConnection = oConexion1
            End If
            .CommandTimeout = 150
            .CommandText = "FacturacionCuentasAtencionExoneracion"
            Set oParameter = .CreateParameter("@IdCuentaAtencion ", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oRsTmp2 = .Execute
           If oConexion1 Is Nothing Then
              Set oRsTmp2.ActiveConnection = Nothing
           End If
       End With
       Set FacturacionCuentasAtencionExoneracion = oRsTmp2
       If oConexion1 Is Nothing Then
          oConexion.Close
       End If
       Set oCommand = Nothing
       Set oParameter = Nothing
       Set oRsTmp2 = Nothing
       Set oConexion = Nothing
    End Function
'Yamill
Function FacturacionCuentasAtencionExonAgregar(lnIdCuentaAtencion As Long, lnNumeroExoneracion As Long, oConexion As Connection) As Boolean
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    FacturacionCuentasAtencionExonAgregar = False

        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "FacturacionCuentasAtencionExonAgregar"
            Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@NumeroExoneracion", adInteger, adParamInput, 0, lnNumeroExoneracion): .Parameters.Append oParameter
            .Execute
        End With
    FacturacionCuentasAtencionExonAgregar = True
    Set oCommand = Nothing
    Set oParameter = Nothing
    
End Function

    
    



Function NroExoneracionXcuenta(lnCuenta As Long, Optional oConexion As Connection) As Long
        Dim oRsTmp3 As New Recordset
        Set oRsTmp3 = FacturacionCuentasAtencionExoneracion(lnCuenta, oConexion)
        NroExoneracionXcuenta = 0
        If oRsTmp3.RecordCount > 0 Then
            oRsTmp3.MoveFirst
            NroExoneracionXcuenta = oRsTmp3.Fields!NumeroExoneracion
        End If
        Set oRsTmp3 = Nothing
End Function



'modificado por Samuel 07/08/2014
Function SeleccionaCierreCaja(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTReporte As Long, IdCajero As Long, idTComprobante As Long, idFarmacia As Long, idVendedor As Long) As ADODB.Recordset
 On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        If idTReporte = 1 Then
            .CommandText = "SeleccionaCierreCajaServicios"
        Else
            .CommandText = "SeleccionaCierreCajaInsumos"
        End If
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 0, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 0, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        If idTReporte = 2 Then
            Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, idFarmacia): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idVendedor", adInteger, adParamInput, 0, idVendedor): .Parameters.Append oParameter
        End If
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionaCierreCaja = oRecordset
    'Set oRecordset = Nothing
    Set oCommand = Nothing
    'Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'modificado por Samuel 09/08/2014
Function SeleccionaCierreCajaTotales(FI As String, FF As String, IdCaja As Long, IdTurno As Long, idTReporte As Long, IdCajero As Long, idTComprobante As Long, idFarmacia As Long, idVendedor As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        If idTReporte = 1 Then
            .CommandText = "SeleccionaCierreCajaServiciosTotales"
        Else
            .CommandText = "SeleccionaCierreCajaInsumosTotales"
        End If
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        If idTReporte = 2 Then
            Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, idFarmacia): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@idVendedor", adInteger, adParamInput, 0, idVendedor): .Parameters.Append oParameter
        End If
        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionaCierreCajaTotales = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'modificado por Samuel 17/08/2014
Function SeleccionaCierreCajaServiciosServiciosHospitalarios(FI As String, FF As String, IdCaja As Long, IdTurno As Long, IdCajero As Long, idTComprobante As Long, idProducto As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaCierreCajaServiciosServiciosHospitalarios"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionaCierreCajaServiciosServiciosHospitalarios = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'modificado por Samuel 17/08/2014
Function SeleccionaCierreCajaServiciosRedondeo(FI As String, FF As String, IdCaja As Long, IdTurno As Long, IdCajero As Long, idTComprobante As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaCierreCajaServiciosRedondeo"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionaCierreCajaServiciosRedondeo = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function
'modificado por Samuel 09/08/2014
Function SeleccionaCierreCajaInsumosRedondeo(FI As String, FF As String, IdCaja As Long, IdTurno As Long, IdCajero As Long, idTComprobante As Long, idFarmacia As Long, idVendedor As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SeleccionaCierreCajaInsumosRedondeo"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IdCaja): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IdTurno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idCajero", adInteger, adParamInput, 0, IdCajero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTComprobante", adInteger, adParamInput, 0, idTComprobante): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFarmacia", adInteger, adParamInput, 0, idFarmacia): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idVendedor", adInteger, adParamInput, 0, idVendedor): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set SeleccionaCierreCajaInsumosRedondeo = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function


'debb-27/05/2015
Function EquivalenciaCPT_SMIseleccionarTodos() As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "EquivalenciaCPT_SMIseleccionarTodos"

        
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set EquivalenciaCPT_SMIseleccionarTodos = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

Function FacturacionServicioDespachoActualizaLAB(lnIdOrden As Long, lnIdProducto As Long, lcLabConfHIS As String, _
                                                 oConexion As Connection) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioDespachoActualizaLAB"
        Set oParameter = .CreateParameter("@IdOrden", adInteger, adParamInput, 0, lnIdOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@labConfHIS", adVarChar, adParamInput, 3, lcLabConfHIS): .Parameters.Append oParameter
        .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
   FacturacionServicioDespachoActualizaLAB = True
   'oConexion.Close
   Set oRecordset = Nothing
   'Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    'MsgBox Err.Description
End Function

'debb-16/07/2015
Function FacturacionCatalogoPaquetesXpaquete(lnIdFactPaquete As Long) As ADODB.Recordset
'On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionCatalogoPaquetesXpaquete"
        Set oParameter = .CreateParameter("@IdFactPaquete", adInteger, adParamInput, 0, lnIdFactPaquete): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set FacturacionCatalogoPaquetesXpaquete = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function

'debb-29/04/2016
Function DevuelveSiElPacienteFallecioOhistoriaPasoPasivo(lnIdPaciente As Long, oConexion As Connection) As String
    Dim mo_ReglasArchivoClinico  As New ReglasArchivoClinico
    Dim mo_ReglasAdmision As New ReglasAdmision
    Dim lnTotalPorPagar As Double, lnTotal As Double
    Dim lcDeuda As String

    Dim oRsTmp1 As New Recordset
'    Dim oCommand As New ADODB.Command
'    Dim oParameter As ADODB.Parameter
    
    Set oRsTmp1 = mo_ReglasAdmision.AtencionesXidPaciente(lnIdPaciente, oConexion)
'    With oCommand
'           .CommandType = adCmdStoredProc
'           Set .ActiveConnection = oConexion
'           .CommandTimeout = 150
'           .CommandText = "AtencionesXidPaciente"
'           Set oParameter = .CreateParameter("@lnIdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
'           Set oRsTmp1 = .Execute
'    End With
    oRsTmp1.Filter = "idcondicionalta=4"
    ms_Texto = ""
    If oRsTmp1.RecordCount > 0 Then
        ms_Texto = "FALLECIDO el " & oRsTmp1!fechaEgreso & "  Cuenta: " & oRsTmp1!IdCuentaAtencion
    End If
    If ms_Texto = "" Then
       If oRsTmp1.State = 1 Then oRsTmp1.Close
       Set oRsTmp1 = mo_ReglasArchivoClinico.HistoriasClinicasXIdPaciente(lnIdPaciente, oConexion)
       If oRsTmp1.RecordCount > 0 Then
          If oRsTmp1!idEstadoHistoria = 2 Or oRsTmp1!idEstadoHistoria = 3 Then
             ms_Texto = "NO está ACTIVA la HISTORIA, verifique ESTADO en opción ARCHIVO CLINICO -> HISTORIA CLINICA  "
          End If
       End If
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
    'Set oCommand = Nothing
    Set mo_ReglasArchivoClinico = Nothing
    Set mo_ReglasAdmision = Nothing
    DevuelveSiElPacienteFallecioOhistoriaPasoPasivo = ms_Texto
End Function

'debb-19/04/2016
Function FacturacionServicioFinanciamientosDespachos(FI As Date, FF As Date, lnIdFuenteFinanciamiento As Long) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FacturacionServicioFinanciamientosDespachos"
        Set oParameter = .CreateParameter("@FechaCobranzaInicial", adDBTimeStamp, adParamInput, 8, FI): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaCobranzaFinal", adDBTimeStamp, adParamInput, 8, FF): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, lnIdFuenteFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set FacturacionServicioFinanciamientosDespachos = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function
'debb-19/04/2016
Function FactCatalogoServiciosXIdProducto(lnIdProducto As Long, oConexion As Connection) As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosXIdProducto"
        Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set FactCatalogoServiciosXIdProducto = oRecordset
    'Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function


