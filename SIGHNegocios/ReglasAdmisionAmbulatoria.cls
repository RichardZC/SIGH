VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasAdmisionAmb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ms_MensajeError As String
Dim mo_AdminArchivoClinico As New SIGHReglasNegocios.ReglasArchivoClinico

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------------------
'                            REGLAS DE NEGOCIOS DEL MODULO DE PACIENTES
'------------------------------------------------------------------------------------------------

Function PacientesAgregarPacienteEHistoriaClinica(oDOPaciente As doPaciente) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oHistoria As New HistoriasClinicas
Dim oDoHistoriaClinica As New DOHistoriaClinica
    
    PacientesAgregarPacienteEHistoriaClinica = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oPaciente.Conexion = oConexion
    
    oConexion.BeginTrans
    oDOPaciente.NroHistoriaClinica = mo_AdminArchivoClinico.GenerarNroHistoriaClinica(oDOPaciente.IdTipoGeneracion)
    
    If oPaciente.Insertar(oDOPaciente) Then
        Select Case oDOPaciente.IdTipoGeneracion
        Case sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaManual, sghHistoriaDefinitivaReciclada
            oDoHistoriaClinica.IdHistoriaClinica = oDOPaciente.NroHistoriaClinica
            oDoHistoriaClinica.FechaCreacion = Format(Date, "DD/MM/YYYY")
            oDoHistoriaClinica.FechaPasoAPasivo = 0
            oDoHistoriaClinica.IdEstadoHistoria = 1
            oDoHistoriaClinica.IdPaciente = oDOPaciente.IdPaciente
            oDoHistoriaClinica.IdTipoHistoria = 1
            oDoHistoriaClinica.IdTipoGeneracionNroHistoria = oDOPaciente.IdTipoGeneracion
            oDoHistoriaClinica.IdUsuarioAuditoria = oDOPaciente.IdUsuarioAuditoria
            Set oHistoria.Conexion = oConexion
            If oHistoria.Insertar(oDoHistoriaClinica) Then
                PacientesAgregarPacienteEHistoriaClinica = True
                oConexion.CommitTrans
            End If
        Case Else
            PacientesAgregarPacienteEHistoriaClinica = True
            oConexion.CommitTrans
        End Select
    Else
        ms_MensajeError = oPaciente.MensajeError
        oConexion.RollbackTrans
    End If
        
    oConexion.Close
    
End Function

Function PacientesModificarYActualizarHistoriaClinicaDefinitiva(oDOPaciente As doPaciente, oDoHistoriaClinica As DOHistoriaClinica) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oHistoria As New HistoriasClinicas
    
    PacientesModificarYActualizarHistoriaClinicaDefinitiva = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oPaciente.Conexion = oConexion
    
    oConexion.BeginTrans
    Select Case oDOPaciente.IdTipoGeneracion
    Case sghHistoriaDefinitivaManual
        'El nro de Historia y la fecha ya vienen con un valor
    Case sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaReciclada
        'Se crea el nro de historia y se inicializan los valores
        oDOPaciente.NroHistoriaClinica = mo_AdminArchivoClinico.GenerarNroHistoriaClinica(oDOPaciente.IdTipoGeneracion)
        oDoHistoriaClinica.IdHistoriaClinica = oDOPaciente.NroHistoriaClinica
    End Select
    
    If oPaciente.Modificar(oDOPaciente) Then
        oDoHistoriaClinica.IdPaciente = oDOPaciente.IdPaciente
        oDoHistoriaClinica.IdUsuarioAuditoria = oDOPaciente.IdUsuarioAuditoria
        oDoHistoriaClinica.FechaPasoAPasivo = 0
        oDoHistoriaClinica.IdEstadoHistoria = 1
        oDoHistoriaClinica.IdTipoHistoria = 1
        oDoHistoriaClinica.IdTipoGeneracionNroHistoria = oDOPaciente.IdTipoGeneracion
        Set oHistoria.Conexion = oConexion
        If oHistoria.Insertar(oDoHistoriaClinica) Then
            PacientesModificarYActualizarHistoriaClinicaDefinitiva = True
            oConexion.CommitTrans
        Else
            ms_MensajeError = oHistoria.MensajeError
            oConexion.RollbackTrans
        End If
    Else
        ms_MensajeError = oPaciente.MensajeError
        oConexion.RollbackTrans
    End If
        
    oConexion.Close
    
End Function

Function PacientesModificar(oDOPaciente As doPaciente) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
    
    PacientesModificar = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oPaciente.Conexion = oConexion
    If oPaciente.Modificar(oDOPaciente) Then
        PacientesModificar = True
    Else
        ms_MensajeError = oPaciente.MensajeError
    End If
    oConexion.Close
    
End Function

Function PacientesEliminar(oDOPaciente As doPaciente) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
    
    PacientesEliminar = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oPaciente.Conexion = oConexion
    If oPaciente.Eliminar(oDOPaciente) Then
        PacientesEliminar = True
    Else
        ms_MensajeError = oPaciente.MensajeError
    End If
    oConexion.Close
    
End Function

Function PacientesSeleccionarPorId(IdPaciente As Integer) As doPaciente
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oDOPaciente As New doPaciente
    
    Set PacientesSeleccionarPorId = Nothing
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPaciente.Conexion = oConexion
    oDOPaciente.IdPaciente = IdPaciente
    If oPaciente.SeleccionarPorId(oDOPaciente) Then
        Set PacientesSeleccionarPorId = oDOPaciente
    End If
    oConexion.Close

    ms_MensajeError = oPaciente.MensajeError

End Function
Function PacientesSeleccionarPorHistoriaClinicaDefinitiva(IdHistoriaClinica As Long) As doPaciente
Dim oConexion As New ADODB.Connection
Dim oPaciente As New Pacientes
Dim oDOPaciente As New doPaciente
    
    Set PacientesSeleccionarPorHistoriaClinicaDefinitiva = Nothing
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oPaciente.Conexion = oConexion
    oDOPaciente.NroHistoriaClinica = IdHistoriaClinica
    
    If oPaciente.SeleccionarPorHistoriaClinicaDefinitiva(oDOPaciente) Then
        Set PacientesSeleccionarPorHistoriaClinicaDefinitiva = oDOPaciente
    End If
    oConexion.Close

    ms_MensajeError = oPaciente.MensajeError

End Function


Function PacientesFiltrar(ByVal oPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrar = oTabla.Filtrar(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

End Function
Function PacientesFiltrarConHistoriasTemporales(ByVal oPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarConHistoriasTemporales = oTabla.FiltrarConHistoriasTemporales(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

End Function
Function PacientesFiltrarConHistoriasDefinitivas(ByVal oPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesFiltrarConHistoriasDefinitivas = oTabla.FiltrarConHistoriasDefinitivas(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

End Function

Function PacientesObtenerConElMismoNombre(ByVal oPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesObtenerConElMismoNombre = oTabla.ObtenerConElMismoNombre(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

End Function
Function PacientesObtenerConElAutogenerado(ByVal oPaciente As doPaciente) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Pacientes
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set PacientesObtenerConElAutogenerado = oTabla.ObtenerConElMismoAutogenerado(oPaciente)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError

End Function

Function PacienteCrearNroAutogenerado(oPaciente As doPaciente)
Dim P1 As String    'Primer digito del apellido paterno
Dim P4 As String    'Cuarto Digito del apellido paterno
Dim M1 As String    'Primer digito del apellido materno
Dim M4 As String    'Cuarto digito del apellido materno
Dim N11 As String   'Primer digito del primer nombre
Dim N41 As String   'Cuarto digito del primer materno
Dim N12 As String   'Primer digito del Ultimo materno
Dim N42 As String   'Cuarto digito del Ultimo materno
Dim D As String     'Digito de verificacion
Dim DD As String
Dim MM As String
Dim AAA As String
Dim sTemp  As String

        DD = Left(oPaciente.FechaNacimiento, 2)
        MM = Mid(oPaciente.FechaNacimiento, 4, 2)
        AAA = Mid(oPaciente.FechaNacimiento, 8, 3)
        DevuelvePrimeryCuartoCaracter oPaciente.ApellidoPaterno, P1, P4
        DevuelvePrimeryCuartoCaracter oPaciente.ApellidoMaterno, M1, M4
        DevuelvePrimeryCuartoCaracter oPaciente.PrimerNombre, N11, N41
        DevuelvePrimeryCuartoCaracter oPaciente.SegundoNombre, N12, N42
        sTemp = AAA + MM + DD & oPaciente.IdTipoSexo & P1 + P4 + M1 + M4 + N11 + N41 + N12 + N42
        PacienteCrearNroAutogenerado = sTemp & Modulo10(sTemp)
        
End Function
Function Modulo10(sValor As String) As Integer
Dim sTemp As String
Dim i As Integer
Dim k As Integer
Dim iTotal As Integer

    sTemp = ""
    
    For i = 1 To Len(sValor)
        If IsNumeric(Mid(sValor, i, 1)) Then
            sTemp = sTemp + Mid(sValor, i, 1)
        Else
            sTemp = sTemp + DevuelveValorEnNumeros(Mid(sValor, i, 1))
        End If
    Next i

    'Acumula total de digitos
    iTotal = 0
    For i = 1 To Len(sTemp)
        If i Mod 2 <> 0 Then
            k = CInt(Mid(sTemp, i, 1)) * 2
            iTotal = iTotal + (k - (k Mod 10)) / 10 + (k Mod 10)
        Else
            iTotal = iTotal + CInt(Mid(sTemp, i, 1))
        End If
    Next i

    If (iTotal Mod 10) = 0 Then
        Modulo10 = 0
    Else
        Modulo10 = 10 - (iTotal Mod 10)
    End If



End Function
Function DevuelveValorEnNumeros(sCaracter As String) As String

    Select Case sCaracter
    Case "A" To "N"
        DevuelveValorEnNumeros = Asc(sCaracter) - 55
    Case "Ñ"
        DevuelveValorEnNumeros = 24
    Case "O" To "Z"
        DevuelveValorEnNumeros = Asc(sCaracter) - 54
    End Select

End Function


Sub DevuelvePrimeryCuartoCaracter(sPalabra As String, C1 As String, C2 As String)
Dim sTemp As String
        If sPalabra <> "" Then
            sTemp = ObtenerUltimaPalabra(EliminarConjunciones(sPalabra))
            C1 = Left(sTemp, 1)
            C2 = DevuelveCuartoCaracter(sTemp)
        Else
            C1 = "X"
            C2 = "X"
        End If
End Sub
Function DevuelveCuartoCaracter(sPalabra) As String
    If Len(sPalabra) <= 4 Then
        DevuelveCuartoCaracter = Right(sPalabra, 1)
    Else
        DevuelveCuartoCaracter = Mid(sPalabra, 4, 1)
    End If
End Function
Function ObtenerPrimeraPalabra(sTexto As String) As String
Dim p As String
Dim sTemp As String

    sTemp = Trim(sTexto)
    
    p = InStr(sTemp, " ")
    If p > 0 Then
        ObtenerPrimeraPalabra = Left(sTemp, p - 1)
    Else
        ObtenerPrimeraPalabra = sTemp
    End If
End Function
Function ObtenerUltimaPalabra(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimaPalabra = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimaPalabra = sTemp
    End If
End Function
Function ObtenerUltimoNombre(sTexto As String) As String
Dim p As String
Dim iUltBlanco As Integer
Dim sTemp As String


    sTemp = Trim(sTexto)

    p = InStr(sTemp, " ")
    iUltBlanco = 0
    Do While p > 0
        iUltBlanco = p
        p = InStr(p + 1, sTemp, " ")
    Loop
    If iUltBlanco > 0 Then
        ObtenerUltimoNombre = Mid(sTemp, iUltBlanco + 1)
    Else
        ObtenerUltimoNombre = ""
    End If
End Function

Function EliminarConjunciones(sPalabra As String)
Dim sTemp As String

        sTemp = SIGHComun.ReemplazarCadena(sPalabra, " DE ", " ")
        sTemp = SIGHComun.ReemplazarCadena(sTemp, " DEL ", " ")
        sTemp = SIGHComun.ReemplazarCadena(sTemp, " EL ", " ")
        sTemp = SIGHComun.ReemplazarCadena(sTemp, " LA ", " ")
        sTemp = SIGHComun.ReemplazarCadena(sTemp, " LOS ", " ")
        sTemp = SIGHComun.ReemplazarCadena(sTemp, " LAS ", " ")

        EliminarConjunciones = sTemp

End Function

'------------------------------------------------------------------------------------------------
'                            REGLAS DE NEGOCIOS DEL MODULO DE CITAS
'------------------------------------------------------------------------------------------------
Function CitasAgregar(oDOCita As doCita, oDOPaciente As doPaciente, bPacienteNuevo As Boolean) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
Dim oPaciente As New Pacientes
Dim oDOPrestamo As New DOPrestamoHistoriaClinica
Dim oPrestamo As New PrestamosHistoria
Dim ProcesoCitasAgregarOK  As Boolean
    
    CitasAgregar = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oCita.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    Set oPrestamo.Conexion = oConexion
    
    oConexion.BeginTrans
    
    'Si elpaciente es nuevo
    If bPacienteNuevo Then
        'Verifica si cumple condicion de creacion de historia clinica
        'If Val(SIGHComun.GenerarHistoriaClinicaSiempre) Or (oDOCita.Fecha = Date) Then
        '    oDOPAciente.IdTipoGeneracion = sghHistoriaDefinitiva
        'Else
            oDOPaciente.IdTipoGeneracion = sghHistoriaTemporalConsultaExterna
        'End If
        
        'Agrega el paciente y la historia clinica
        If PacientesAgregarPacienteEHistoriaClinica(oDOPaciente) Then
            oDOCita.IdPaciente = oDOPaciente.IdPaciente
            'Agrega la cita
            'If Val(SIGHComun.SolicitarHistoriaEnFormaAutomatica) Then
                'Agrega la solicitud de la historia
            '    InicializarPrestamoAutomatico oDOCita, oDOPAciente, oDOPrestamo, bPacienteNuevo
            '    If oPrestamo.Insertar(oDOPrestamo) Then
            '        oDOCita.IdPrestamo = oDOPrestamo.IdPrestamo
            '        If oCita.Insertar(oDOCita) Then
            '            ProcesoCitasAgregarOK = True
            '        End If
            '    End If
            'Else
                If oCita.Insertar(oDOCita) Then
                    ProcesoCitasAgregarOK = True
                End If
            'End If
        End If
    Else
        'Si elpaciente no es nuevo (Tiene filiacion)
        oDOCita.IdPaciente = oDOPaciente.IdPaciente
        
        'Verifica el tipo de historia clinica que tiene el paciente
        Select Case oDOPaciente.IdTipoGeneracion
        Case sghHistoriaDefinitivaAutomatica, sghHistoriaDefinitivaManual, sghHistoriaDefinitivaReciclada
                'Si tiene historia definitva, solo agrega su cita
                If Val(SIGHComun.SolicitarHistoriaEnFormaAutomatica) Then
                    'Agrega la solicitud de la historia
                    InicializarPrestamoAutomatico oDOCita, oDOPaciente, oDOPrestamo, bPacienteNuevo
                    If oPrestamo.Insertar(oDOPrestamo) Then
                        oDOCita.IdPrestamo = oDOPrestamo.IdPrestamo
                        If oCita.Insertar(oDOCita) Then
                            ProcesoCitasAgregarOK = True
                        Else
                            ms_MensajeError = oCita.MensajeError
                        End If
                    Else
                        ms_MensajeError = oPrestamo.MensajeError
                    End If
                Else
                    If oCita.Insertar(oDOCita) Then
                        ProcesoCitasAgregarOK = True
                    Else
                        ms_MensajeError = oCita.MensajeError
                    End If
                End If
        Case Else
            'Si el paciente no tiene historia, o su historia el tempral de CE o emergencia
            'Se verifica la condicion de generacion de historias
'            If Val(SIGHComun.GenerarHistoriaClinicaSiempre) Or (oDOCita.Fecha = Date) Then
'                'Si se permite la creacion de historias, se actualiza la historia
'                'antigua y se inserta en la tabla de historias clinica
'                oDOPAciente.IdTipoGeneracion = sghHistoriaDefinitiva
'                If PacientesActualizarHistoriaClinicaDefinitiva(oDOPAciente) Then
'                    oDOCita.IdPaciente = oDOPAciente.IdPaciente
'                    'Se graba la cita
'                    If Val(SIGHComun.SolicitarHistoriaEnFormaAutomatica) Then
'                        'Agrega la solicitud de la historia
'                        InicializarPrestamoAutomatico oDOCita, oDOPAciente, oDOPrestamo, bPacienteNuevo
'                        If oPrestamo.Insertar(oDOPrestamo) Then
'                            oDOCita.IdPrestamo = oDOPrestamo.IdPrestamo
'                            If oCita.Insertar(oDOCita) Then
'                                ProcesoCitasAgregarOK = True
'                            End If
'                        End If
'                    Else
'                        If oCita.Insertar(oDOCita) Then
'                            ProcesoCitasAgregarOK = True
'                        End If
'                    End If
'                End If
'            Else
                oDOCita.IdPaciente = oDOPaciente.IdPaciente
                If oCita.Insertar(oDOCita) Then
                    ProcesoCitasAgregarOK = True
                Else
                    ms_MensajeError = oPrestamo.MensajeError
                End If
'            End If
        End Select
    End If
    
    If ProcesoCitasAgregarOK Then
        CitasAgregar = True
        oConexion.CommitTrans
    Else
        CitasAgregar = False
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    
End Function
Sub InicializarPrestamoAutomatico(oDOCita As doCita, oDOPaciente As doPaciente, oDOPrestamo As DOPrestamoHistoriaClinica, bPacienteNuevo As Boolean)
        
        oDOPrestamo.FechaSolicitud = Format(Now, "dd/mm/yyyy")
        oDOPrestamo.HoraSolicitud = Format(Now, "hh:mm")
        
        oDOPrestamo.FechaPrestamoRequerida = oDOCita.Fecha
        oDOPrestamo.HoraPrestamoRequerida = oDOCita.HoraInicio
        
        oDOPrestamo.FechaPrestamoReal = 0
        oDOPrestamo.HoraPrestamoReal = ""
        
        oDOPrestamo.FechaDevolucion = 0
        oDOPrestamo.HoraDevolucion = ""
        
        oDOPrestamo.IdEnvio = 0
        oDOPrestamo.IdEstadoPrestamo = 1
        oDOPrestamo.IdHistoriaClinica = oDOPaciente.NroHistoriaClinica
        oDOPrestamo.IdMotivo = 1
        oDOPrestamo.IdServicio = oDOCita.IdServicio
        oDOPrestamo.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
        oDOPrestamo.Observacion = "CITA CE: " + IIf(bPacienteNuevo, "PACIENTE NUEVO", "") + " Fecha: " & oDOCita.Fecha & " (" & oDOCita.HoraInicio & " - " & oDOCita.HoraFin + ")"
        
End Sub

Function CitasModificar(oDOCita As doCita) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
    
    CitasModificar = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oCita.Conexion = oConexion
    If oCita.Modificar(oDOCita) Then
        CitasModificar = True
    Else
        ms_MensajeError = oCita.MensajeError
    End If
    oConexion.Close
    
End Function

Function CitasEliminar(oDOCita As doCita) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
Dim oDOPrestamo As New DOPrestamoHistoriaClinica
Dim oPrestamo As New PrestamosHistoria
Dim bProcesoCitasEliminarOK  As Boolean
    
    bProcesoCitasEliminarOK = False
    
    oConexion.Open SIGHComun.CadenaConexion
    Set oCita.Conexion = oConexion
    Set oPrestamo.Conexion = oConexion
    
    oConexion.BeginTrans
    
    If oCita.Eliminar(oDOCita) Then
        oDOPrestamo.IdPrestamo = oDOCita.IdPrestamo
        oDOPrestamo.IdUsuarioAuditoria = oDOCita.IdUsuarioAuditoria
        If oPrestamo.Eliminar(oDOPrestamo) Then
            bProcesoCitasEliminarOK = True
        End If
    Else
        ms_MensajeError = oCita.MensajeError
    End If
    
    If bProcesoCitasEliminarOK Then
        oConexion.CommitTrans
        CitasEliminar = True
    Else
        oConexion.RollbackTrans
    End If
    
    
    oConexion.Close
    
End Function

Function CitasSeleccionarPorId(IdCita As Long, oDOCita As doCita, oDOPaciente As doPaciente) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCita As New Citas
Dim oPaciente As New Pacientes
    
    CitasSeleccionarPorId = False
    
    oConexion.Open SIGHComun.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCita.Conexion = oConexion
    Set oPaciente.Conexion = oConexion
    
    oDOCita.IdCita = IdCita
    If oCita.SeleccionarPorId(oDOCita) Then
        oDOPaciente.IdPaciente = oDOCita.IdPaciente
        If oPaciente.SeleccionarPorId(oDOPaciente) Then
            CitasSeleccionarPorId = True
        End If
    End If
    oConexion.Close
    
    ms_MensajeError = oCita.MensajeError

End Function

