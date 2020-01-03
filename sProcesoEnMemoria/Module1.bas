Attribute VB_Name = "Module1"
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Módulo que Pide Usuario y Clave, además de CLAVE ESPECIAL
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public lcNumeroPc As String
Public wxNumMaxHoraLibre As Long            'Nro MINUTOS donde el Cliente tiene su HORA LIBRE
Public wxMinTiempoMinimoEnCabina As Long    'Nro MINUTOS minimos que un Cliente puede alquilar
Public wxMinTiempoMaximoEnCabina As Long    'Nro MINUTOS maximos que un Cliente puede alquilar
Public wxMinUltimoAvisoParaAlarma As Long   'Nro MINUTOS antes que acabe donde la ALARMA aparece
Public wxMinRegaloNumero As Long            'Minutos que se regalan desde wxMinRegaloApartirDe
Public wxMinRegaloApartirDe As Long         'a partir del MINUTO se regalan wxMinRegaloNumero
Public wxTime As String                     'HORA para todas las CABINAS
Public wxDate As Date                       'FECHA para todas las CABINAS
Public lnCliente As Long                    'Codigo del CLIENTE
Public wxUsuarioSist As String              'Usuario que ADMINISTRA LA CABINA
'Public Const wxCadenaConexion As String = "DSN=cabina"    '"Driver=Microsoft Access Driver (*.mdb);DBQ=\\Servidor\dbfs\cabina.mdb;Password=debb"
Public wxCadenaConexion As String
Public wxNombreServidorRed As String        'Nombre de la PC que será el Servidor en la RED
Public wxRutaBaseDatos As String            'Ruta de la Base de Datos en el Servidor
Public wxFechaInicioCompetencia As Date, wxPremio As String 'Para Premios por feriados
Public wxDiasQueSeVeHorasAcumuladasParaRegalo   'Dias en el Mes que se puede ver Horas acumuladas para el PREMIO
Public wxMuestraGrid As String      'HOSPITALIZACION,EMERGENCIA,CE,TODOS,CERRAR   (muestra o no grid con las CUENTAS QUE SERAN CERRADAS)
Public wxNumMinutosGrid As String   'Numero de Minutos en que se mostrará el GRID o CERRARAN las cuentas
Public wxReniecHoraInicio As String       'Hora Inicio para el proceso de comparar pacientes RENIEC vs GALENHOS
Public wxReniecHoraFin As String          'Hora Final para el proceso de comparar pacientes RENIEC vs GALENHOS
Public wxSisAcreditacioHoraInicio As String   'Hora en que comienza a IMPORTAR DATOS DEL SIS - ACREDITACIONES
Public wxSisAcreditacioHoraFinal As String   'Hora en que termina a IMPORTAR DATOS DEL SIS - ACREDITACIONES
'***************************Ocultar Programa cuando se pulsa CTRL-ALT-DEL
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long
Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0
Public Sub MakeMeService()
    Dim pid As Long
    Dim regserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub
Public Sub UnMakeMeService()
    Dim pid As Long
    Dim regserv As Long
    pid = GetCurrentProcessId()
    regserv = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub
'***************************Ocultar Programa cuando se pulsa CTRL-ALT-DEL


Function sGetINI(sIniFile As String, sSection As String, sKey As String, sDefault As String) As String
    Dim sTemp As String * 256
    Dim nLength As Integer
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sIniFile)
    sGetINI = Left$(sTemp, nLength)
End Function


Public Sub CargaHoraFechaEnEsteMomento()
      On Error GoTo ErrCargaFH
      wxTime = Time  'Esto debe cargarlo desde TGENERAL cada minuto
      wxDate = Date
      Dim wxConexionRed As New ADODB.Connection
      wxConexionRed.CommandTimeout = 150
      wxConexionRed.CursorLocation = adUseClient
      wxConexionRed.Open wxCadenaConexion
      Dim wrs_uhc As New ADODB.Recordset
      wrs_uhc.CursorLocation = adUseClient
      wrs_uhc.Open "select * from TEMPRESA", wxConexionRed, adOpenKeyset, adLockOptimistic
      If Not IsNull(wrs_uhc.Fields("horaActual").Value) And Not IsNull(wrs_uhc.Fields("FechaActual").Value) Then
        wxTime = wrs_uhc.Fields("horaActual").Value
        wxDate = wrs_uhc.Fields("FechaActual").Value
      End If
      wxUsuarioSist = IIf(IsNull(wrs_uhc.Fields("usuario").Value), "", wrs_uhc.Fields("usuario").Value)
      wrs_uhc.Close
      Set wrs_uhc = Nothing
      wxConexionRed.Close
      Set wxConexionRed = Nothing
ErrCargaFH:
End Sub

Public Sub GrabaHoraFechaEnEsteMomento()
      On Error GoTo ErrGrabaHF
      wxTime = Time  'Esto debe grabarlo en TGENERAL cada minuto
      wxDate = Date
      Dim wxConexionRed As New ADODB.Connection
      wxConexionRed.CommandTimeout = 150
      wxConexionRed.CursorLocation = adUseClient
      wxConexionRed.Open wxCadenaConexion
      Dim wrs_uhc As New ADODB.Recordset
      wrs_uhc.CursorLocation = adUseClient
      wrs_uhc.Open "select * from TEMPRESA", wxConexionRed, adOpenKeyset, adLockOptimistic
      wrs_uhc.Fields("horaActual").Value = wxTime
      wrs_uhc.Fields("FechaActual").Value = wxDate
      wrs_uhc.Update
      wrs_uhc.Close
      Set wrs_uhc = Nothing
      wxConexionRed.Close
      Set wxConexionRed = Nothing
      Exit Sub
ErrGrabaHF:
      MsgBox "No está Grabando Hora y Fecha ->" & Err.Description
End Sub


Public Function ValidaMinutosEnCabina(lcMinutosEnCabina As String) As Boolean
    ValidaMinutosEnCabina = False
     If Val(lcMinutosEnCabina) < wxMinTiempoMinimoEnCabina And Val(lcMinutosEnCabina) <> 0 Then
       MsgBox "El tiempo mímimo de alquiler de una CABINA es " & wxMinTiempoMinimoEnCabina & " minutos", vbCritical, ""
       Exit Function
     End If
     If Val(lcMinutosEnCabina) > wxMinTiempoMaximoEnCabina And Val(lcMinutosEnCabina) <> 0 Then
       MsgBox "El tiempo máximo de alquiler de una CABINA es " & wxMinTiempoMaximoEnCabina & " minutos", vbCritical, ""
       Exit Function
     End If
     ValidaMinutosEnCabina = True
End Function


Public Function DevuelveMinutosRealesEnCabina(lcMinutosEnCabina As String) As Long
        If Val(lcMinutosEnCabina) = 0 Then
           DevuelveMinutosRealesEnCabina = 0
        ElseIf Val(lcMinutosEnCabina) >= wxMinRegaloApartirDe Then
           DevuelveMinutosRealesEnCabina = Val(lcMinutosEnCabina) + wxMinRegaloNumero
        Else
           DevuelveMinutosRealesEnCabina = Val(lcMinutosEnCabina)
        End If
End Function

Public Sub AgregaBD(oConexionRed As ADODB.Connection, lcMinutosEnCabina As String, lnCliente As Long, lcCliente As String, lnNumPC As Long, lcObservacion As String, lnDebemos As Integer, LcHoraLibre As String, lnNumeroMinutosHl As Long)
    On Error GoTo ErrAgr
    Dim oRsTmp As New ADODB.Recordset
    Dim lnMinutosEnCabina As Long
    Dim lnNumeroMov As Long
    Dim lcTexto As String
    Dim lcDeben As String: Dim lcDebemos As String
    With oRsTmp
        If LcHoraLibre = "S" Then
           lnMinutosEnCabina = Val(lcMinutosEnCabina)
        Else
           lnMinutosEnCabina = DevuelveMinutosRealesEnCabina(lcMinutosEnCabina)
        End If
        
        Dim lcHrInicio24 As String
        Dim lcHrFinal24 As String
        Dim lcHrFinalAmPm As String
        lcHrInicio24 = CambiaFormatoTime24(Format(wxTime, "hh:mm AM/PM"))
        lcHrFinal24 = SumaMinutoTimes(lcHrInicio24, Str(lnMinutosEnCabina))
        lcHrFinalAmPm = CambiaFormatoTimeAMPM(lcHrFinal24)
        
        
        
        .CursorLocation = adUseClient
        .Open "select * from tcabinas where codigo=" & lnNumPC, oConexionRed, adOpenKeyset, adLockOptimistic
        .Fields!horaIcli = Format(wxTime, "hh:mm AM/PM")
        .Fields!horaScli = lcHrFinalAmPm
        .Fields!fechaIcli = wxDate
        .Fields!numMinutosCli = lnMinutosEnCabina
        .Fields!numMinutosQuedan = IIf(Val(lcMinutosEnCabina) <= 0, 0, 9999)
        .Fields!ApagaPC = 0
        .Fields!ReiniciaPC = 0
        If lnCliente > 0 Then
            .Fields!dCliente = lcCliente
            .Fields!Cliente = lnCliente
        End If
        .Fields!numeroMovCli = 0
        .Update
        .Close
        'movimientos
        If Val(lcMinutosEnCabina) > 0 Then
            .CursorLocation = adUseClient
            .Open "select * from movimientos", oConexionRed, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("tipo").Value = "V"
            .Fields("servicio").Value = 1
            .Fields("usuario").Value = wxUsuarioSist
            .Fields("horalibre").Value = LcHoraLibre
            .Fields("MinSobrHL").Value = 0
            .Fields("cliente").Value = lnCliente
            .Fields("fecha").Value = wxDate
            .Fields("hingreso").Value = Format(wxTime, "hh:mm AM/PM")   ' Devuelve "05:04:23 PM"
            .Fields("hsalida").Value = "__:__ AM"
            'oRsMovim.Fields("importe").Value = Val(txtImporte.Text)
            .Fields("NumMinutos").Value = lnMinutosEnCabina
            .Fields("numcabina").Value = Val(lcNumeroPc)
            .Fields("observaciones").Value = lcObservacion
            .Fields("minPedidos").Value = lnMinutosEnCabina
            If LcHoraLibre = "S" Then
               .Fields("minSobrHL").Value = lnNumeroMinutosHl - (wxNumMaxHoraLibre * 60)
            End If
            .Fields("auditoria").Value = Trim(Str(lnCliente)) & "-" & Trim(Str(lnMinutosEnCabina)) & "&" & Format(wxTime, "hh:mm AM/PM") & "/"
            .Update
            lnNumeroMov = .Fields!numeroMov
            .Close
            .CursorLocation = adUseClient
            .Open "select * from tcabinas where codigo=" & lnNumPC, oConexionRed, adOpenKeyset, adLockOptimistic
            .Fields!numeroMovCli = lnNumeroMov
            .Update
            .Close
        End If
        
        .CursorLocation = adUseClient
        .Open "select * from clientes where codigo=" & lnCliente, oConexionRed, adOpenKeyset, adLockOptimistic
        If lnDebemos > 0 Then
            lcTexto = DevuelveDebeDebemosCliente(.Fields("deben").Value, "")
            .Fields("debemos").Value = 0
            .Fields("dni").Value = lcTexto
            .Update
        End If
        lcDeben = IIf(IsNull(.Fields!deben), "", .Fields!deben)
        lcDebemos = IIf(IsNull(.Fields!debemos), "", .Fields!debemos)
        .Close
        'actualiza datos
        .Open "select * from tcabinas where codigo=" & lnNumPC, oConexionRed, adOpenKeyset, adLockOptimistic
        .Fields!deben = lcDeben
        .Fields!debemos = lcDebemos
        .Update
        .Close
        
    End With
    Set oRsTmp = Nothing
ErrAgr:
End Sub


Public Function DevuelveDebeDebemosCliente(lcDeben As String, lcDebemos As String) As String
        DevuelveDebeDebemosCliente = ""
        If Trim(lcDebemos) <> "" Then
           DevuelveDebeDebemosCliente = DevuelveDebeDebemosCliente & "Debemos(min): " & Trim(lcDebemos)
        End If
        If Trim(lcDeben) <> "" Then
           DevuelveDebeDebemosCliente = DevuelveDebeDebemosCliente & "- deben(S/): " & Trim(lcDeben)
        End If
        If Trim(DevuelveDebeDebemosCliente) <> "" Then
           DevuelveDebeDebemosCliente = Left(DevuelveDebeDebemosCliente & "-" & Date, 40)
        End If
End Function

'Resta 2 Tiempos (HH:MM) de 0 a 24 horas
Function RestaTimes(Time1 As String, time2 As String) As String
    RestaTimes = "__:__"
    If Time1 <> "__:__" And time2 <> "__:__" Then
        whora1 = Val(Left(Time1, 2)): wminu1 = Val(Right(Time1, 2))
        whora2 = Val(Left(time2, 2)): wminu2 = Val(Right(time2, 2))
        wdif = ((whora2 * 60) + wminu2) - ((whora1 * 60) + wminu1)
        If wdif >= 0 Then
           whora = Int(wdif / 60)
           wminu = wdif - whora * 60
           RestaTimes = Right("0" & Trim(Str(whora)), 2) & ":" & Right("0" & Trim(Str(wminu)), 2)
        Else
           whora2 = whora2 + 24
           wdif = ((whora2 * 60) + wminu2) - ((whora1 * 60) + wminu1)
           whora = Int(wdif / 60)
           wminu = wdif - whora * 60
           RestaTimes = Right("0" & Trim(Str(whora)), 2) & ":" & Right("0" & Trim(Str(wminu)), 2)
        End If
    End If
End Function


'cambia el Formato de una FECHA (hh:mm PM) hacia 0 a 24 horas
Function CambiaFormatoTime24(TimeAmPm1 As String) As String
    CambiaFormatoTime24 = "__:__"
    If Left(TimeAmPm1, 5) <> "__:__" Then
        whora1 = Val(Left(TimeAmPm1, 2))
        wminu1 = Val(Mid(TimeAmPm1, 4, 2))
        If UCase(Right(TimeAmPm1, 2)) <> "AM" Then
           If whora1 <> 12 Then
              whora1 = whora1 + 12
           End If
        Else
           If whora1 = 12 Then
              whora1 = 0
           End If
        End If
        CambiaFormatoTime24 = Right("0" & Trim(Str(whora1)), 2) & ":" & Right("0" & Trim(Str(wminu1)), 2)
    End If
End Function

' Tiempo (HH:MM) de 0 a 24 horas
Function DevuelveMinutosDeUnaHora(Time1 As String) As Long
    DevuelveMinutosDeUnaHora = 0
    If Time1 <> "__:__" Then
        whora1 = Val(Left(Time1, 2)): wminu1 = Val(Right(Time1, 2))
        DevuelveMinutosDeUnaHora = ((whora1 * 60) + wminu1)
    End If
End Function

Function MinutosRegaladosQueSeDescuentan(lnMinPedidos As Long, lnMinQuedan As Long) As Long
    MinutosRegaladosQueSeDescuentan = 0
    If lnMinPedidos >= wxMinRegaloApartirDe And lnMinQuedan <= wxMinRegaloNumero Then
        MinutosRegaladosQueSeDescuentan = wxMinRegaloNumero
    End If
End Function

Function MinutosConsumidos(lcHrInicio As String, lcHrSalio As String, lnMinPedidos As Long, lnMinQuedan As Long) As Long
    If lnMinPedidos >= wxMinRegaloApartirDe And lnMinQuedan <= wxMinRegaloNumero Then
       MinutosConsumidos = lnMinPedidos - wxMinRegaloNumero
    Else
       MinutosConsumidos = DevuelveMinutosDeUnaHora(RestaTimes(CambiaFormatoTime24(lcHrInicio), CambiaFormatoTime24(Format(lcHrSalio, "hh:mm AM/PM"))))
    End If
End Function

'Suma Minutos(MM) a una Hora (HH:MM de 0 a 24 horas)
Function SumaMinutoTimes(Time1 As String, Minutos1 As String) As String
    SumaMinutoTimes = "__:__"
    If Time1 <> "__:__" And Val(Minutos1) > 0 Then
        whora1 = Val(Left(Time1, 2))
        wminu1 = Val(Right(Time1, 2))
        wminu2 = Val(Minutos1)
        wtotal = (whora1 * 60) + wminu1 + wminu2
        If wtotal >= 0 Then
           whora = Int(wtotal / 60)
           wminu = wtotal - whora * 60
           SumaMinutoTimes = Right("0" & Trim(Str(whora)), 2) & ":" & Right("0" & Trim(Str(wminu)), 2)
        End If
    End If
End Function

'cambia el Formato de una FECHA (desde 0 a 24 horas) hacia (hh:mm PM)
Function CambiaFormatoTimeAMPM(Time24 As String) As String
    CambiaFormatoTimeAMPM = "__:__"
    If Left(Time24, 5) <> "__:__" Then
        whora1 = Val(Left(Time24, 2))
        wminu1 = Val(Mid(Time24, 4, 2))
        If whora1 > 12 Then
           whora1 = whora1 - 12
           CambiaFormatoTimeAMPM = Right("0" & Trim(Str(whora1)), 2) & ":" & Right("0" & Trim(Str(wminu1)), 2) & " PM"
        ElseIf whora1 = 12 Then
           CambiaFormatoTimeAMPM = Right("0" & Trim(Str(whora1)), 2) & ":" & Right("0" & Trim(Str(wminu1)), 2) & " PM"
        Else
           CambiaFormatoTimeAMPM = Right("0" & Trim(Str(whora1)), 2) & ":" & Right("0" & Trim(Str(wminu1)), 2) & " AM"
        End If
    End If
End Function

Sub CargaIni()
 On Error GoTo ErrINI
 Dim sIniFile As String
 sIniFile = App.Path & "\setup.ini"
 If Dir$(sIniFile) <> "" Then
    wxNumMinutosGrid = sGetINI(sIniFile, "Variables", "NumMinutosGrid", "?")
    wxMuestraGrid = sGetINI(sIniFile, "Variables", "MuestraGrid", "?")
    wxReniecHoraInicio = sGetINI(sIniFile, "Variables", "ReniecHoraInicio", "?")
    wxReniecHoraFin = sGetINI(sIniFile, "Variables", "ReniecHoraFin", "?")
    wxSisAcreditacioHoraInicio = sGetINI(sIniFile, "Variables", "SisHoraInicio", "?")
    wxSisAcreditacioHoraFinal = sGetINI(sIniFile, "Variables", "SisHoraFinal", "?")
 Else
    MsgBox "Se ha borrado el archivo de Configuracion", vbCritical, ""
 End If
 Exit Sub
ErrINI:
  End
End Sub


Sub main()
    CargaIni
    If wxMuestraGrid = "CuposLibres" Or wxMuestraGrid = "ATENCIONCE" Then
       MonitoresTV.Show 1
    Else
       Procesos.Show
    End If
End Sub

