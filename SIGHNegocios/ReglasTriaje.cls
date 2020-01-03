VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasTriaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para TRIAJE o SIGNOS VITALES
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad


Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


Public Function ListaVariableTriajeTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oDatos As New TriajeVariable
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    Set oDatos.Conexion = oConexion
    Set ListaVariableTriajeTodos = oDatos.ListarVariableTodos
    ms_MensajeError = oDatos.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oDatos = Nothing
    
    Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  
End Function

Public Function ListaVariableTriajeAntropometricas() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oDatos As New TriajeVariable
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    Set oDatos.Conexion = oConexion
    Set ListaVariableTriajeAntropometricas = oDatos.ListarVariableSAntropometricas
    ms_MensajeError = oDatos.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oDatos = Nothing
    
    Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  
End Function


Public Function ListarValorNormalesSegunParametros(oDOTriajeValorNormal As DOTriajeValorNormal) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oDatos As New TriajeValorNormal
    Dim oConexion As New ADODB.Connection
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
'    oConexion.Open SIGHEntidades.CadenaConexion
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    
    Set oDatos.Conexion = oConexion
    Set ListarValorNormalesSegunParametros = oDatos.ListarValorNormalesSegunParametros(oDOTriajeValorNormal)
    ms_MensajeError = oDatos.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oDatos = Nothing
    
    Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
  
End Function

Public Function RetornaObjetoValorNormalParaBusqueda(oDoPaciente As doPaciente, _
                    oDoAtencion As DOAtencion, _
                    Optional lEstadoPaciente As sighTriajeEstadoPaciente = NoRequerido) As DOTriajeValorNormal
                    
    Dim oDOTriajeValorNormal As New DOTriajeValorNormal
    Dim lEdadPacienteEnDias As Long
    Dim oFechaHora As New FechaHora
    
    lEdadPacienteEnDias = oFechaHora.EdadActualEnDias(oDoPaciente.FechaNacimiento, oDoAtencion.FechaIngreso)
    
    oDOTriajeValorNormal.EdadInicialEnDia = lEdadPacienteEnDias
    oDOTriajeValorNormal.SexoPaciente = oDoPaciente.IdTipoSexo
    oDOTriajeValorNormal.FechaVigencia = oDoAtencion.FechaIngreso
    oDOTriajeValorNormal.EstadoPaciente = lEstadoPaciente
    
    Set RetornaObjetoValorNormalParaBusqueda = oDOTriajeValorNormal
End Function


Public Function muestraValoresNormalesTriaje(IdTriajeVariable As Long, _
                rsValoresNormalesTriaje As ADODB.Recordset, _
                Optional AddUnidadMedida As Boolean = True, _
                Optional esPresion As Boolean = False, _
                Optional delimitardorPresion As String = "/") As String
                
    Dim sValorNormal As String
    sValorNormal = ""
    
    If esPresion = False And (IdTriajeVariable = sighTriajeVariable.PresArtSistolica _
                        Or IdTriajeVariable = sighTriajeVariable.PresArtDiastolica) Then
                        
        sValorNormal = "Sist/Diast : " & muestraValoresNormalesTriaje(sighTriajeVariable.PresArtSistolica, _
                                rsValoresNormalesTriaje, False, True, delimitardorPresion)
        sValorNormal = sValorNormal & delimitardorPresion & muestraValoresNormalesTriaje(sighTriajeVariable.PresArtDiastolica, _
                            rsValoresNormalesTriaje, False, True, delimitardorPresion)
    Else
        If Not (rsValoresNormalesTriaje.BOF = True And rsValoresNormalesTriaje.EOF = True) Then
            rsValoresNormalesTriaje.MoveFirst
            rsValoresNormalesTriaje.Find "IdTriajeVariable=" & IdTriajeVariable
            If rsValoresNormalesTriaje.EOF = False Then
                If Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) Or Not IsNull(rsValoresNormalesTriaje!ValorNormalMaximo) Then
                    If Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) And _
                                Not IsNull(rsValoresNormalesTriaje!ValorNormalMaximo) Then
                                
                        sValorNormal = rsValoresNormalesTriaje!ValorNormalMinimo & " a " & _
                                        rsValoresNormalesTriaje!ValorNormalMaximo
                                        
                    ElseIf Not IsNull(rsValoresNormalesTriaje!ValorNormalMinimo) Then
                        sValorNormal = "min. " & rsValoresNormalesTriaje!ValorNormalMinimo
                    Else
                        sValorNormal = "max." & rsValoresNormalesTriaje!ValorNormalMaximo
                    End If
                End If
            End If
        End If
        
    End If
    If AddUnidadMedida = True Then
        sValorNormal = sValorNormal & unidadMedidaTriaje(IdTriajeVariable)
    End If
    muestraValoresNormalesTriaje = sValorNormal
miError:
    If Err Then
        MsgBox Err.Number & " " & Err.Description
    End If
End Function

Private Function unidadMedidaTriaje(IdTriajeVariable As Long)
    Select Case IdTriajeVariable
        Case sighTriajeVariable.Temperatura
            unidadMedidaTriaje = " °C"
        Case sighTriajeVariable.Peso
            unidadMedidaTriaje = " Kg"
        Case sighTriajeVariable.Talla
            unidadMedidaTriaje = " Cm"
        Case Else
            unidadMedidaTriaje = ""
    End Select
End Function

Public Function buscarValorNormalVariable(IdTriajeVariable As sighTriajeVariable, rsValoresNormalesTriaje As ADODB.Recordset) As Boolean
    buscarValorNormalVariable = False
    If Not (rsValoresNormalesTriaje Is Nothing) Then
        If Not (rsValoresNormalesTriaje.BOF = True And rsValoresNormalesTriaje.EOF = True) Then
            rsValoresNormalesTriaje.MoveFirst
            rsValoresNormalesTriaje.Find "IdTriajeVariable=" & IdTriajeVariable
            If rsValoresNormalesTriaje.EOF = False Then
                buscarValorNormalVariable = True
            End If
        End If
    End If
End Function

Public Function RetornaTieneFiebre(cTemperatura As String, rsValoresNormalesTriaje As ADODB.Recordset) As Boolean
    Dim TieneFiebre As Boolean
    
    TieneFiebre = False
    If cTemperatura <> "" Then
        If buscarValorNormalVariable(Temperatura, rsValoresNormalesTriaje) = True Then
            If Val(cTemperatura) > rsValoresNormalesTriaje!ValorNormalMaximo Then
                If IsNull(rsValoresNormalesTriaje!ValorCoherenteMaximo) Then
                    TieneFiebre = True
                ElseIf Val(cTemperatura) <= rsValoresNormalesTriaje!ValorCoherenteMaximo Then
                    TieneFiebre = True
                End If
            End If
        End If
    End If
    RetornaTieneFiebre = TieneFiebre
End Function


Public Function OcultarControlesCRED(EsAtencionCRED As Boolean, ParamArray aControl() As Variant) As Boolean
    On Error Resume Next
    Dim oControl As Variant
'    If EsAtencionCRED = False Then
        For Each oControl In aControl
            oControl.Visible = EsAtencionCRED
        Next
'    End If
    OcultarControlesCRED = True
    Err = 0
End Function

'mgaray20141013
Public Function ListaTriajeExcepcionesTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oTriajeExcepciones As New TriajeExcepciones
    Dim oConexion As New ADODB.Connection
    
    
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)

    Set oTriajeExcepciones.Conexion = oConexion
    Set ListaTriajeExcepcionesTodos = oTriajeExcepciones.SeleccionarTodos()
    ms_MensajeError = oTriajeExcepciones.MensajeError
    oConexion.Close
    Set oConexion = Nothing
    Set oTriajeExcepciones = Nothing
    
    Exit Function
ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function

'/*****************************************************************************************/
'/*******************************************INO*******************************************/
'/*****************************************************************************************/
'Function AtencionesRapidaInoAgregar(oDOAtencionesRapidaIno As doatencionesrapidaino) As Boolean
'Dim oConexion As New ADODB.Connection
'Dim oAtencionesRapidaIno As New Atencionesrapidaino
'
'    AtencionesRapidaInoAgregar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionesRapidaIno.Conexion = oConexion
'
'    If oAtencionesRapidaIno.Insertar(oDOAtencionesRapidaIno) Then
'        AtencionesRapidaInoAgregar = True
'
'    Else
'        ms_MensajeError = oAtencionesRapidaIno.MensajeError
'
'        End If
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oAtencionesRapidaIno = Nothing
'
'End Function
'
'
'
'Function AtencionesRapidaInoEliminar(oDOAtencionesRapidaIno As doatencionesrapidaino) As Boolean
'    On Error GoTo ErrCeAgr
'    Dim oConexion As New ADODB.Connection
'    Dim oAtencionesRapidaIno As New Atencionesrapidaino
'    AtencionesRapidaInoEliminar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionesRapidaIno.Conexion = oConexion
'   ' oConexion.BeginTrans
'
'    Set oAtencionesRapidaIno.Conexion = oConexion
'    If oAtencionesRapidaIno.Eliminar(oDOAtencionesRapidaIno) Then
'        '
'       ' Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencionTriaje.IdUsuarioAuditoria, "E", oDoAtencionTriaje.idAtencion, "AtencionesCE", oConexionSIGH, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)            'ListBarItems.idListItem
'        AtencionesRapidaInoEliminar = True
'       ' oConexionSIGH.CommitTrans
'       ' oConexion.CommitTrans
'    Else
'        ms_MensajeError = oAtencionesRapidaIno.MensajeError
''        oConexionSIGH.RollbackTrans
''        oConexion.RollbackTrans
'    End If
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oAtencionesRapidaIno = Nothing
'    Exit Function
'ErrCeAgr:
'End Function
'
'
'
'Function AtencionesRapidaInoModificar(oDOAtencionesRapidaIno As doatencionesrapidaino) As Boolean
'    On Error GoTo ErrCeAgr
'    Dim oConexion As New ADODB.Connection
'
'    Dim oAtencionesRapidaIno As New Atencionesrapidaino
'    AtencionesRapidaInoModificar = False
'
'    oConexion.Open sighentidades.CadenaConexion
'    Set oAtencionesRapidaIno.Conexion = oConexion
'
''    If oDoAtencionTriaje.TriajeFecha = 0 Then
''       oDoAtencionTriaje.TriajeFecha = lcBuscaParametro.RetornaFechaHoraServidorSQL
''    End If
'    If oAtencionesRapidaIno.modificar(oDOAtencionesRapidaIno) Then
'        '
'        'Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoAtencionTriaje.IdUsuarioAuditoria, "M", oDoAtencionTriaje.idAtencion, "AtencionesCE", oConexionSIGH, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)           'ListBarItems.idListItem
'        AtencionesRapidaInoModificar = True
''        oConexionSIGH.CommitTrans
''        oConexion.CommitTrans
'    Else
'        ms_MensajeError = oAtencionesRapidaIno.MensajeError
''        oConexionSIGH.RollbackTrans
''        oConexion.RollbackTrans
'    End If
'    oConexion.Close
'    Set oConexion = Nothing
'    Set oAtencionesRapidaIno = Nothing
'    Exit Function
'ErrCeAgr:
'End Function

Function BuscarTriajeOftalmologico(idAtencion As String) As Recordset

Dim oRsTmp1 As New Recordset
Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "AtencionesRapidaInoSeleccionarPorId"
        Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, idAtencion): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
    End With
    Set BuscarTriajeOftalmologico = oRsTmp1
    Set oCommand = Nothing
End Function

