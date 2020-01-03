VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasAntecedentesPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para reglas de antecedentes del Paciente
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String

Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Public Function ListarPreguntasPorPacienteYGrupo(oGrupoHCPaciente As DOAtenInteGrupoHCPaciente) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenInteGrupoHCPaciente
    Dim rsPregunta As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPregunta = oTabla.ListarPreguntasPorPacienteYGrupo(oGrupoHCPaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set ListarPreguntasPorPacienteYGrupo = rsPregunta
End Function


Public Function ListarRespuestasPorPacienteYGrupo(oGrupoHCPaciente As DOAtenInteGrupoHCPaciente) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenInteHCRespuestaPaciente
    Dim rsRespuestas As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsRespuestas = oTabla.ListarRespuestasPorPacienteYGrupo(oGrupoHCPaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
    Set ListarRespuestasPorPacienteYGrupo = rsRespuestas
End Function

Public Function grabarRespuestasPaciente(cPreguntas As Collection, _
        cRespuestas As Collection) As Boolean
On Error GoTo miError

    Dim ErrorInTransaction As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oGrupoHCPaciente As DOAtenInteGrupoHCPaciente
    Dim oGrupoHcRespuestaPaciente As DOAtenInteHCRespuestaPaciente
    Dim oDOPregunta As New DOAtenIntePregunta
    
    grabarRespuestasPaciente = False
    ErrorInTransaction = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    If Not (cPreguntas Is Nothing) Then
        For Each oGrupoHCPaciente In cPreguntas
            If verificarAgregarPreguntaPaciente(oGrupoHCPaciente, oConexion) = True Then
                For Each oGrupoHcRespuestaPaciente In cRespuestas
                    If oGrupoHCPaciente.idPregunta = oGrupoHcRespuestaPaciente.idPregunta Then
                        Call RespustasPacienteEliminarPorPregunta(oGrupoHCPaciente, oConexion)
                        
                        oGrupoHcRespuestaPaciente.IdGrupoHCPaciente = oGrupoHCPaciente.IdGrupoHCPaciente
                        
                        If verificarAgregarRespuestaPaciente(oGrupoHcRespuestaPaciente, _
                                    oConexion) = False Then
                            ErrorInTransaction = True
                            GoTo EndTransaction
                        End If
                    End If
                Next
                'auditar
            Else
                ErrorInTransaction = True
                GoTo EndTransaction
            End If
        Next
    End If
EndTransaction:
    Dim response As Boolean
    response = False
    If ErrorInTransaction = True Then
        oConexion.RollbackTrans
    Else
        oConexion.CommitTrans
        response = True
    End If
    oConexion.Close
    grabarRespuestasPaciente = response
miError:
    If Err Then
        
        ms_MensajeError = Err.Number & " : " & Err.Description
        If oConexion.State = 1 Then
        
            oConexion.Close
        End If
    End If
    Set oConexion = Nothing
End Function

Private Function verificarAgregarPreguntaPaciente(ByRef oGrupoHCPaciente As DOAtenInteGrupoHCPaciente, _
        oConexion As ADODB.Connection) As Boolean
        
    Dim GrupoHCPacinte As New AtenInteGrupoHCPaciente
    
    Set GrupoHCPacinte.Conexion = oConexion
    verificarAgregarPreguntaPaciente = GrupoHCPacinte.verificarInsertar(oGrupoHCPaciente)
    If verificarAgregarPreguntaPaciente = False Then
        ms_MensajeError = GrupoHCPacinte.MensajeError
    End If
    Set GrupoHCPacinte = Nothing
    
End Function

Private Function verificarAgregarRespuestaPaciente(ByRef oDoGrupoHcRespuestaPaciente As DOAtenInteHCRespuestaPaciente, _
        oConexion As ADODB.Connection) As Boolean
        
    Dim oHCRespuestaPaciente As New AtenInteHCRespuestaPaciente
    
    Set oHCRespuestaPaciente.Conexion = oConexion
    verificarAgregarRespuestaPaciente = oHCRespuestaPaciente.verificaInsertar(oDoGrupoHcRespuestaPaciente)
    If verificarAgregarRespuestaPaciente = False Then
        ms_MensajeError = oHCRespuestaPaciente.MensajeError
    End If
    Set oHCRespuestaPaciente = Nothing
End Function

Function RespustasPacienteEliminarPorPregunta(ByRef oGrupoHCPaciente As DOAtenInteGrupoHCPaciente, _
        oConexion As ADODB.Connection) As Boolean
    
    Dim oHCRespuestaPaciente As New AtenInteHCRespuestaPaciente
    
    Set oHCRespuestaPaciente.Conexion = oConexion
    RespustasPacienteEliminarPorPregunta = oHCRespuestaPaciente.EliminarRespuestaPorPregunta(oGrupoHCPaciente)
    If RespustasPacienteEliminarPorPregunta = False Then
        ms_MensajeError = oHCRespuestaPaciente.MensajeError
    End If
    Set oHCRespuestaPaciente = Nothing
End Function

Public Function PreguntaSeleccionarPorId(idPregunta As Long) As DOAtenIntePregunta
        
    Dim oPregunta As New AtenIntePregunta
    Dim oDOPregunta As New DOAtenIntePregunta
    
    
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set PreguntaSeleccionarPorId = Nothing
    
    Set oPregunta.Conexion = oConexion
    oDOPregunta.idPregunta = idPregunta
    If oPregunta.SeleccionarPorId(oDOPregunta) Then
        Set PreguntaSeleccionarPorId = oDOPregunta
    End If
    ms_MensajeError = oPregunta.MensajeError
    Set oPregunta = Nothing
    Set oDOPregunta = Nothing
    
    oConexion.Close
    Set oConexion = Nothing
End Function
