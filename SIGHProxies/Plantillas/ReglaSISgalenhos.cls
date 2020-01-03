VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasSISgalenhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas del SIS
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReglasComunes As New ReglasComunes
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_SisConsumoWeb As New SIGHNegocios.SisConsumoWeb
Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim lnValorActualProgress As Long
Public Event ProgressActualizaValor(ByRef lnValorActual As Long, ByRef lnValorTotal As Long, ByRef lcTablaActualizar As String)

Property Let ValorActualProgress(lValue As Long)
   lnValorActualProgress = lValue
End Property
Property Get ValorActualProgress() As Long
   ValorActualProgress = lnValorActualProgress
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'debb
'Actualizado frank 1109
Sub GeneraTXT(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, orst As Recordset, lnNroEnvio As String, lnMesEnvio As String, lnAnioEnvio As String)
        On Error GoTo ErrGenTXT
        Dim sw As Boolean 'Variable para verificar palotes en lclinea
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long, lcDNI As String
        Const lcSeparadorCampo As String = "|"
        Dim lcDisa As String, lcFormato As String, lcNumero As String, lcSql As String, lnCuenta As Long
        Dim oConexionExterna As New Connection
        Dim oRsAfiliadosSIS As New Recordset
        Dim oDoSisFuaAtencion As New DoSisFuaAtencion, oSisFuaAtencion As New SisFuaAtencion
        Dim oDoSisFuaAtencionDIA As New DoSisFuaAtencionDIA, oSisFuaAtencionDIA As New SisFuaAtencionDIA
        Dim oDoSisFuaAtencionINS As New DoSisFuaAtencionINS, oSisFuaAtencionINS As New SisFuaAtencionINS
        Dim oDoSisFuaAtencionMED As New DoSisFuaAtencionMED, oSisFuaAtencionMED As New SisFuaAtencionMED
        Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO, oSisFuaAtencionPRO As New SisFuaAtencionPRO
        Dim oDoSisFuaAtencionSMI As New DoSisFuaAtencionSMI, oSisFuaAtencionSMI As New SisFuaAtencionSMI
        Dim oDoCatalogoServicio As New DOCatalogoServicio, oCatalogoServicio As New CatalogoServicios
        Dim rstCatalogo As ADODB.Recordset
        
        Dim fso As Object 'crea objeto para manejar archivos
        Dim act As Object 'crea objeto para manejar archivos
        
        If Right(lcArchivoTXT, 13) = "Pacientes.txt" Or Right(lcArchivoTXT, 12) = "ATENCION.txt" Or Right(lcArchivoTXT, 7) = "DIA.txt" Or Right(lcArchivoTXT, 7) = "INS.txt" Or Right(lcArchivoTXT, 7) = "MED.txt" Or Right(lcArchivoTXT, 7) = "PRO.txt" Or Right(lcArchivoTXT, 7) = "SMI.txt" Then
            oConexionExterna.CommandTimeout = 300
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            oConexionExterna.CursorLocation = adUseClient
            If Right(lcArchivoTXT, 12) = "ATENCION.txt" Then
               Set oSisFuaAtencion.Conexion = oConexionExterna
            ElseIf Right(lcArchivoTXT, 7) = "DIA.txt" Then
               Set oSisFuaAtencionDIA.Conexion = oConexionExterna
            ElseIf Right(lcArchivoTXT, 7) = "INS.txt" Then
               Set oSisFuaAtencionINS.Conexion = oConexionExterna
            ElseIf Right(lcArchivoTXT, 7) = "MED.txt" Then
               Set oSisFuaAtencionMED.Conexion = oConexionExterna
            ElseIf Right(lcArchivoTXT, 7) = "PRO.txt" Then
               Set oSisFuaAtencionPRO.Conexion = oConexionExterna
            ElseIf Right(lcArchivoTXT, 7) = "SMI.txt" Then
               Set oSisFuaAtencionSMI.Conexion = oConexionExterna
            End If
        End If
        Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
        Set act = fso.CreateTextFile(lcArchivoTXT, True) 'creo el archivo con nombre asignado por parametro
        lnUltimoCampo = oRsTmp1.Fields.Count
        lnContador = 1
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           DoEvents: lnContador = lnContador + 1
           '
           If Right(lcArchivoTXT, 12) = "ATENCION.txt" Then
              lnCuenta = oRsTmp1!IdCuentaAtencion
           ElseIf Right(lcArchivoTXT, 7) = "DIA.txt" Or Right(lcArchivoTXT, 7) = "INS.txt" Or Right(lcArchivoTXT, 7) = "MED.txt" Or Right(lcArchivoTXT, 7) = "PRO.txt" Or Right(lcArchivoTXT, 7) = "SMI.txt" Then
              lnCuenta = oRsTmp1!id
           End If
           lcLinea = ""
           lcDNI = ""
           sw = True 'Inicailizo sw en true
           orst.MoveFirst   'Mueve el registro al primero para iniciar la comparacion
           For lnCont = 0 To lnUltimoCampo - 1
'********************************************************************modificacion para reducir campos
                If (orst.Fields!Estado.Value = True) Then 'Genera el filtro para solo los activos de recordset de estadostrama
                    If UCase(oRsTmp1.Fields(lnCont).Name) = "CABNROENVIOALSIS" Then
                       If Right(lcArchivoTXT, 12) = "ATENCION.txt" Then
                             oDoSisFuaAtencion.IdCuentaAtencion = lnCuenta
                             If Not oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) Then
                                MsgBox oSisFuaAtencion.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencion.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencion.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencion.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencion.Modificar(oDoSisFuaAtencion) Then
                                MsgBox oSisFuaAtencion.MensajeError: GoTo ErrGenTXT
                             End If
                       ElseIf Right(lcArchivoTXT, 7) = "DIA.txt" Then
                             oDoSisFuaAtencionDIA.id = lnCuenta
                             If Not oSisFuaAtencionDIA.SeleccionarPorId(oDoSisFuaAtencionDIA) Then
                                MsgBox oSisFuaAtencionDIA.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencionDIA.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencionDIA.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencionDIA.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencionDIA.Modificar(oDoSisFuaAtencionDIA) Then
                                MsgBox oSisFuaAtencionDIA.MensajeError: GoTo ErrGenTXT
                             End If
                       ElseIf Right(lcArchivoTXT, 7) = "INS.txt" Then
                             oDoSisFuaAtencionINS.id = lnCuenta
                             If Not oSisFuaAtencionINS.SeleccionarPorId(oDoSisFuaAtencionINS) Then
                                MsgBox oSisFuaAtencionINS.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencionINS.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencionINS.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencionINS.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencionINS.Modificar(oDoSisFuaAtencionINS) Then
                                MsgBox oSisFuaAtencionINS.MensajeError: GoTo ErrGenTXT
                             End If
                       ElseIf Right(lcArchivoTXT, 7) = "MED.txt" Then
                             oDoSisFuaAtencionMED.id = lnCuenta
                             If Not oSisFuaAtencionMED.SeleccionarPorId(oDoSisFuaAtencionMED) Then
                                MsgBox oSisFuaAtencionMED.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencionMED.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencionMED.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencionMED.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencionMED.Modificar(oDoSisFuaAtencionMED) Then
                                MsgBox oSisFuaAtencionMED.MensajeError: GoTo ErrGenTXT
                             End If
                       ElseIf Right(lcArchivoTXT, 7) = "PRO.txt" Then
                             oDoSisFuaAtencionPRO.id = lnCuenta
                             If Not oSisFuaAtencionPRO.SeleccionarPorId(oDoSisFuaAtencionPRO) Then
                                MsgBox oSisFuaAtencionPRO.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencionPRO.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencionPRO.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencionPRO.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencionPRO.Modificar(oDoSisFuaAtencionPRO) Then
                                MsgBox oSisFuaAtencionPRO.MensajeError: GoTo ErrGenTXT
                             End If
                       ElseIf Right(lcArchivoTXT, 7) = "SMI.txt" Then
                             oDoSisFuaAtencionSMI.id = lnCuenta
                             If Not oSisFuaAtencionSMI.SeleccionarPorId(oDoSisFuaAtencionSMI) Then
                                MsgBox oSisFuaAtencionSMI.MensajeError: GoTo ErrGenTXT
                             End If
                             If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                                oDoSisFuaAtencionSMI.CabNroEnvioAlSIS = 1
                             Else
                                oDoSisFuaAtencionSMI.CabNroEnvioAlSIS = Trim(Str(Val(oDoSisFuaAtencionSMI.CabNroEnvioAlSIS) + 1))
                             End If
                             If Not oSisFuaAtencionSMI.Modificar(oDoSisFuaAtencionSMI) Then
                                MsgBox oSisFuaAtencionSMI.MensajeError: GoTo ErrGenTXT
                             End If
                       End If
                    End If
                    lcCampoValor = ""
                    
                    If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                         Select Case oRsTmp1.Fields(lnCont).Type
                         Case adDate
                              lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                         Case adChar, adVarChar
                              lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                              If lcCampoValor = lcSeparadorCampo Then
                                 lcCampoValor = "_"
                              End If
                         Case adDouble, adInteger
                              lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                         Case Else
                              lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                         End Select
                    End If
'***********************************modficacion para modelar los datos de salida segun trama********
                    If lcCampoValor = "__/__/____" Then lcCampoValor = ""
                    If (oRsTmp1.Fields(lnCont).Name = "Mes") Then lcCampoValor = Format(lnMesEnvio, "00")
                    If (oRsTmp1.Fields(lnCont).Name = "CabNroEnvioAlSIS") Then lcCampoValor = Format(lnNroEnvio, "00")
                    If (oRsTmp1.Fields(lnCont).Name = "FuaNumero") Then lcCampoValor = Mid(Format(lcCampoValor, "00000000"), 1, 8)
                    If (oRsTmp1.Fields(lnCont).Name = "Valor") Then lcCampoValor = Mid(Replace(lcCampoValor, "_", ""), 1, 8) ' Frank 1109
                    If (oRsTmp1.Fields(lnCont).Name = "CabOrigenDelRegistro") Then lcCampoValor = Format(lcCampoValor, "0000000000")
                    If (oRsTmp1.Fields(lnCont).Name = "DxTipoDPR") Then
                        If lcCampoValor = "D" Then lcCampoValor = "1"
                        If lcCampoValor = "P" Then lcCampoValor = "2"
                        If lcCampoValor = "C" Then lcCampoValor = "3"
                        If lcCampoValor = "R" Then lcCampoValor = "4"
                    End If
                    ''''''averiguar si ya hay procedimientode fechas
                    If (oRsTmp1.Fields(lnCont).Name = "EstablecimientoCodigoRenaes") Then lcCampoValor = Format(lcCampoValor, "0000000000")
                    If (oRsTmp1.Fields(lnCont).Name = "FuaAtencionFecha") Or (oRsTmp1.Fields(lnCont).Name = "FuaHospitalizadoFingreso") Then
                        If (IsNull(lcCampoValor)) Or lcCampoValor = "" Then lcCampoValor = "" Else lcCampoValor = forDDMMYYYY(lcCampoValor)
                    End If
                    If (oRsTmp1.Fields(lnCont).Name = "FuaHospitalizadoFalta") Or (oRsTmp1.Fields(lnCont).Name = "fnacimiento") Then
                        If (IsNull(lcCampoValor)) Or lcCampoValor = "" Then lcCampoValor = "" Else lcCampoValor = forDDMMYYYY(lcCampoValor)
                    End If
                    'salida de codigo de catalogo de servicios, cambiando por el codigosis que se modifico en tabla
                    If Right(lcArchivoTXT, 7) = "PRO.txt" And oRsTmp1.Fields(lnCont).Name = "Codigo" Then
                        Set rstCatalogo = oCatalogoServicio.SeleccionarPorCodigo(lcCampoValor)
                        If Not IsNull(rstCatalogo.Fields!CodigoSIS) Then
                            lcCampoValor = rstCatalogo.Fields!CodigoSIS
                        Else
                            lnBarraTotal = lnBarraTotal - 1
                            GoTo siguiente_registro
                        End If
                    End If
                    
                    '*******************************fin modficacion para modelar los datos de salida segun trama********
                    If sw Then
                       lcLinea = lcLinea & lcCampoValor
                       sw = False
                    Else
                       lcLinea = lcLinea & lcSeparadorCampo & lcCampoValor
                    End If
                    If Right(lcArchivoTXT, 13) = "Pacientes.txt" And UCase(oRsTmp1.Fields(lnCont).Name) = "NRODOCUMENTO" Then
                       lcDNI = lcCampoValor
                    End If
                End If 'cierra el if de filtro
                orst.MoveNext 'mueve el recorset de estado de trama
           Next
           '
           If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
                lcDisa = "": lcFormato = "": lcNumero = ""
                If Len(lcDNI) = 8 Then
                     Set oRsAfiliadosSIS = SisFiliacionesXdni(lcDNI, oConexionExterna)
                     If oRsAfiliadosSIS.RecordCount > 0 Then
                        lcDisa = Trim(oRsAfiliadosSIS.Fields!AfiliacionDisa)
                        lcFormato = Trim(oRsAfiliadosSIS.Fields!AfiliacionTipoFormato)
                        lcNumero = Trim(oRsAfiliadosSIS.Fields!AfiliacionNroFormato)
                     End If
                     oRsAfiliadosSIS.Close
                End If
                lcLinea = lcLinea & lcDisa & lcSeparadorCampo & lcFormato & lcSeparadorCampo & lcNumero & lcSeparadorCampo
           End If
            lcLinea = lcLinea + Chr(10) 'inserta el caracter cl
            act.Write (lcLinea) 'graba la linea

siguiente_registro:
           oRsTmp1.MoveNext
        Loop
        act.Close
        Set oConexionExterna = Nothing
        Set oRsAfiliadosSIS = Nothing
        Set oDoSisFuaAtencion = Nothing
        Set oSisFuaAtencion = Nothing
        Set oDoSisFuaAtencionDIA = Nothing
        Set oSisFuaAtencionDIA = Nothing
        Set oDoSisFuaAtencionINS = Nothing
        Set oSisFuaAtencionINS = Nothing
        Set oDoSisFuaAtencionMED = Nothing
        Set oSisFuaAtencionMED = Nothing
        Set oDoSisFuaAtencionPRO = Nothing
        Set oSisFuaAtencionPRO = Nothing
        Set oDoSisFuaAtencionSMI = Nothing
        Set oSisFuaAtencionSMI = Nothing
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
       ' Resume
End Sub

Function LeerArchivoTexto(nombreFichero As String) As String
    Dim numlib As Integer, isOpen As Boolean
    Dim LongitudArchivo As Long
    Dim lcLeerArchivoTexto As String
    Dim VarTexto As String
    On Error GoTo Manejador_Error
    ' Obtengo el siguiente número libre de archivo
    numlib = FreeFile()
    Open nombreFichero For Input As #numlib
    ' Se ha abierto el fichero sin problemas
    isOpen = True
    ' Leo todo el contenido en una única operación
    LongitudArchivo = LOF(numlib)
    If LongitudArchivo <= 32767 Then
        lcLeerArchivoTexto = Input(LOF(numlib), #numlib)
    Else
        Do Until EOF(numlib)
            Line Input #numlib, VarTexto
            lcLeerArchivoTexto = lcLeerArchivoTexto & VarTexto & vbCrLf
        Loop
    End If
    LeerArchivoTexto = lcLeerArchivoTexto
'   LeerArchivoTexto = Input(1, #numlib)
    ' Cierro el archivo
Manejador_Error:
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

Sub DOS_CreandoRutaEnDisco(lcRutaTemporal As String)
    Dim lcSql As String
    On Error Resume Next
    lcSql = "c:"
    sighentidades.ejecutarComando lcSql
    lcSql = "cd \"
    sighentidades.ejecutarComando lcSql
    lcSql = "md " & lcRutaTemporal
    sighentidades.ejecutarComando lcSql
    lcSql = "cd " & lcRutaTemporal
    sighentidades.ejecutarComando lcSql
    lcSql = "del *.*"
    sighentidades.ejecutarComando lcSql
End Sub



Sub ImportaAcreditadosSIScursor(lcArchivoAcreditadosZIP As String)
     On Error GoTo ErrImporSIS
     Dim lcSql As String
     Dim lcArchivoTmp As String, lcArchivo As String
     Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
     Dim mo_ReglasComunes As New ReglasComunes
     Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single
     Dim lnIdSiaSis As Long, lcCodigo As String, lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String
     Dim lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, lcDocumentoTipo As String
     Dim lcCodigoEstablAdscripcion As String, ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String
     Dim lcPnombre As String, lcOnombres As String, lcGenero As String, ldFnacimiento As Date
     Dim lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, lcDocumentoNumero As String
     Dim lcMotivoBaja As String, lcfecha As String, lcHoraInicioProceso As String, lcHoraFinalProceso As String
     Dim oRsTmp1 As New Recordset, oConexion As New Connection
     Dim lineas() As String, lnRegistros As Long, lnContador As Long, lcRename As String
     Dim oCrypKey As New CrypKey.Util
     Dim oCommand As New ADODB.Command
     ms_MensajeError = ""
     lcArchivoTmp = "c:\debbSis"
     DOS_CreandoRutaEnDisco lcArchivoTmp
     sighentidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(318)), lcArchivoAcreditadosZIP, lcArchivoTmp, False
     lcArchivo = Dir(lcArchivoTmp & "\*.*")
     lcRename = lcArchivoTmp & "\" & lcArchivo
     lcSql = lcArchivoTmp & "\filiacionesSIS.txt"
     FileCopy lcRename, lcSql
     '
     oConexion.CommandTimeout = 300
     oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexion.CursorLocation = adUseClient
     With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "SisFiliacionesImportaAcreditados"
       .Execute
     End With
     '
     mo_ReglasComunes.ParametrosActualizaValorTextoXid 312, "Afiliados SIS: " & lcArchivo
     '
     Set oRsTmp1 = Nothing
     oConexion.Close
     Set oConexion = Nothing
     Set mo_ReglasSISgalenhos = Nothing
     On Error Resume Next
     lcArchivoTmp = "DEL " & lcArchivoTmp
     sighentidades.ejecutarComando lcArchivoTmp
     
     Exit Sub
ErrImporSIS:
    ms_MensajeError = Err.Description
    MsgBox Err.Description
    
    'Resume

End Sub

Sub ImportaAcreditadosSIS(lcArchivoAcreditadosZIP As String)
     On Error GoTo ErrImporSIS
     Dim lcArchivoTmp As String, lcArchivo As String, lcSql As String
     Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
     Dim mo_ReglasComunes As New ReglasComunes
     Dim Fichero As Integer, Registro As String, Campos() As String, Columna As Single
     Dim lnIdSiaSis As Long, lcCodigo As String, lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String
     Dim lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, lcDocumentoTipo As String
     Dim lcCodigoEstablAdscripcion As String, ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String
     Dim lcPnombre As String, lcOnombres As String, lcGenero As String, ldFnacimiento As Date
     Dim lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, lcDocumentoNumero As String
     Dim lcMotivoBaja As String, lcfecha As String, lcHoraInicioProceso As String, lcHoraFinalProceso As String
     Dim oRsTmp1 As New Recordset, oConexion As New Connection
     Dim lineas() As String, lnRegistros As Long, lnContador As Long, lcMensajeError As String
     Dim oCrypKey As New CrypKey.Util
     Dim fila As Long 'Fila As Single
     ms_MensajeError = ""
     lcArchivoTmp = "c:\debbSis"
     
     'Frank cambio 14112014
     EliminarArchivosDirectorio lcArchivoTmp
     ''''''''''''''''''''''
     
     DOS_CreandoRutaEnDisco lcArchivoTmp
     sighentidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(318)), lcArchivoAcreditadosZIP, lcArchivoTmp, False
     lcArchivo = Dir(lcArchivoTmp & "\*.*")
     '
'     lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
'     lnRegistros = UBound(lineas)

     lnRegistros = TotalRegistrosArchivoTexto(lcArchivoTmp & "\" & lcArchivo)
     
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     
     Fichero = FreeFile
     Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
     lnContador = 0
     While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, ",")
        lnIdSiaSis = 0
        lcCodigo = ""
        lcAfiliacionDisa = ""
        lcAfiliacionTipoFormato = ""
        lcAfiliacionNroFormato = ""
        lcAfiliacionNroIntegrante = ""
        lcDocumentoTipo = ""
        lcCodigoEstablAdscripcion = ""
        ldAfiliacionFecha = 0
        lcPaterno = ""
        lcMaterno = ""
        lcPnombre = ""
        lcOnombres = ""
        lcGenero = ""
        ldFnacimiento = 0
        lcIdDistritoDomicilio = ""
        lcEstado = ""
        ldFbaja = 0 'Actualizado 16102014
        lcDocumentoNumero = ""
        lcMotivoBaja = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                 lnIdSiaSis = Val(Campos(Columna))
            Case 1
                lcCodigo = Campos(Columna)
            Case 2
                lcAfiliacionDisa = Campos(Columna)
            Case 3
                lcAfiliacionTipoFormato = Campos(Columna)
            Case 4
                lcAfiliacionNroFormato = Campos(Columna)
            Case 5
                lcAfiliacionNroIntegrante = Campos(Columna)
            Case 6
                lcDocumentoTipo = Campos(Columna)
            Case 7
                lcCodigoEstablAdscripcion = Campos(Columna)
            Case 8
                lcfecha = Campos(Columna)
                lcfecha = Right(lcfecha, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Left(lcfecha, 4)
                If sighentidades.EsFecha(lcfecha, "DD/MM/AAAA") Then
                   ldAfiliacionFecha = CDate(lcfecha)
                End If
            Case 9
                lcPaterno = Campos(Columna)
            Case 10
                lcMaterno = Campos(Columna)
            Case 11
                lcPnombre = Campos(Columna)
            Case 12
                lcOnombres = Campos(Columna)
            Case 13
                lcGenero = Campos(Columna)
            Case 14
                lcfecha = Campos(Columna)
                lcfecha = Right(lcfecha, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Left(lcfecha, 4)
                If sighentidades.EsFecha(lcfecha, "DD/MM/AAAA") Then
                    ldFnacimiento = CDate(lcfecha)
                End If
            Case 15
                lcIdDistritoDomicilio = Campos(Columna)
            Case 16
                lcEstado = Campos(Columna)
            Case 17
                lcfecha = Campos(Columna)
                lcfecha = Right(lcfecha, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Left(lcfecha, 4)
                If sighentidades.EsFecha(lcfecha, "DD/MM/AAAA") Then
                    ldFbaja = lcfecha
                 End If
            Case 18
                lcDocumentoNumero = Campos(Columna)
            Case 19
                lcMotivoBaja = Campos(Columna)
            End Select
        Next
        lcMensajeError = SisFiliacionesBuscaYactualizaDatosXafiliado(oConexion, lnIdSiaSis, lcCodigo, lcAfiliacionDisa, _
                                                    lcAfiliacionTipoFormato, lcAfiliacionNroFormato, _
                                                    lcAfiliacionNroIntegrante, lcDocumentoTipo, _
                                                    lcCodigoEstablAdscripcion, ldAfiliacionFecha, lcPaterno, _
                                                    lcMaterno, lcPnombre, lcOnombres, lcGenero, ldFnacimiento, _
                                                    lcIdDistritoDomicilio, lcEstado, ldFbaja, lcDocumentoNumero, _
                                                    lcMotivoBaja)
        ms_MensajeError = ms_MensajeError & Chr(13) & lcMensajeError
        
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "")
        
        fila = fila + 1
    Wend
    Close #Fichero
    '
    mo_ReglasComunes.ParametrosActualizaValorTextoXid 312, "Afiliados SIS: " & lcArchivo
    
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set mo_ReglasSISgalenhos = Nothing
    On Error Resume Next
    lcArchivoTmp = "DEL " & lcArchivoTmp
    sighentidades.ejecutarComando lcArchivoTmp
    Exit Sub
ErrImporSIS:
    If ms_MensajeError = "" Then
       ms_MensajeError = Err.Description
    End If
    MsgBox ms_MensajeError
    'Resume
End Sub

Function SisFiliacionesBuscaYactualizaDatosXafiliado(oConexion As Connection, lnIdSiaSis As Long, lcCodigo As String, _
                                          lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String, _
                                          lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, _
                                          lcDocumentoTipo As String, lcCodigoEstablAdscripcion As String, _
                                          ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String, _
                                          lcPnombre As String, _
                                          lcOnombres As String, lcGenero As String, ldFnacimiento As Date, _
                                          lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, _
                                          lcDocumentoNumero As String, lcMotivoBaja As String) As String
        Dim oSisFiliaciones As New SisFiliaciones, oDoSisFiliaciones As New DoSisFiliaciones
        Dim lbEsNuevo As Boolean
        SisFiliacionesBuscaYactualizaDatosXafiliado = ""
        Set oSisFiliaciones.Conexion = oConexion
        oDoSisFiliaciones.idSiasis = lnIdSiaSis
        oDoSisFiliaciones.codigo = lcCodigo
        lbEsNuevo = True
        oDoSisFiliaciones.idSiasis = lnIdSiaSis
        oDoSisFiliaciones.codigo = lcCodigo
        oDoSisFiliaciones.AfiliacionDisa = lcAfiliacionDisa
        oDoSisFiliaciones.AfiliacionTipoFormato = lcAfiliacionTipoFormato
        oDoSisFiliaciones.AfiliacionNroFormato = lcAfiliacionNroFormato
        oDoSisFiliaciones.AfiliacionNroIntegrante = lcAfiliacionNroIntegrante
        oDoSisFiliaciones.DocumentoTipo = lcDocumentoTipo
        oDoSisFiliaciones.CodigoEstablAdscripcion = lcCodigoEstablAdscripcion
        If ldAfiliacionFecha <> 0 Then
           oDoSisFiliaciones.AfiliacionFecha = ldAfiliacionFecha
        End If
        oDoSisFiliaciones.Paterno = lcPaterno
        oDoSisFiliaciones.Materno = lcMaterno
        oDoSisFiliaciones.Pnombre = lcPnombre
        oDoSisFiliaciones.Onombres = lcOnombres
        oDoSisFiliaciones.Genero = lcGenero
        If ldFnacimiento <> 0 Then
           oDoSisFiliaciones.Fnacimiento = ldFnacimiento
        End If
        oDoSisFiliaciones.IdDistritoDomicilio = lcIdDistritoDomicilio
        oDoSisFiliaciones.Estado = lcEstado
        oDoSisFiliaciones.Fbaja = IIf(ldFbaja = 0, "", IIf(IsNull(ldFbaja), "", Format(ldFbaja, sighentidades.DevuelveFechaSoloFormato_DMY)))  'Actualizado 15102014
        oDoSisFiliaciones.DocumentoNumero = lcDocumentoNumero
        oDoSisFiliaciones.MotivoBaja = lcMotivoBaja
        If oSisFiliaciones.Insertar(oDoSisFiliaciones, True) = False Then
           SisFiliacionesBuscaYactualizaDatosXafiliado = oSisFiliaciones.MensajeError
        End If
        Set oSisFiliaciones = Nothing
        Set oDoSisFiliaciones = Nothing
End Function

'debb
Function ChequeaCuentasSinFUAgrabado(ByVal mrs_Tmp As Recordset, oConexionExterna As Connection, lcFechaInicio As String, _
                                     lcFechaFinal As String) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lnTotalReg As Long, lnIdCuentaInicial As Long, lbSeImprime As Boolean, lcSql As String
        On Error GoTo ErrorProceso1
        '
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
             .CommandText = "AtencionesSeleccionaSoloSISxFechas"
             Set oParameter = .CreateParameter("@FechaEgrIni", adDBTimeStamp, adParamInput, 0, CDate(lcFechaInicio)): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@FechaEgrFin", adDBTimeStamp, adParamInput, 0, CDate(lcFechaFinal)): .Parameters.Append oParameter
             Set oRsTmp1 = .Execute
             Set oRsTmp1.ActiveConnection = Nothing
        End With
        '
        lnTotalReg = oRsTmp1.RecordCount
        'Chequea si hay Cuentas Observadas (No se grabaron en tablas SIS FUA)
        If lnTotalReg > 0 Then
           lnIdCuentaInicial = 1
           With mrs_Tmp
                  .Fields.Append "idCuentaAtencion", adInteger
                  .Fields.Append "ApellidoPaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "apellidoMaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "PrimerNombre", adVarChar, 50, adFldIsNullable
                  .Fields.Append "nroHistoriaClinica", adInteger, 0, adFldIsNullable
                  .Fields.Append "tipoServicio", adVarChar, 50, adFldIsNullable
                  .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                  .Fields.Append "FechaIngreso", adDate, 8, adFldIsNullable
                  .Fields.Append "FechaEgreso", adDate, 8, adFldIsNullable
                  .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
                  .LockType = adLockOptimistic
                  .Open
           End With
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              Set oRsTmp2 = SisFuaAtencionSeleccionarPorCuenta(oRsTmp1.Fields!IdCuentaAtencion, oConexionExterna)
              If oRsTmp2.RecordCount = 0 Then
                 lbSeImprime = False
                 If oRsTmp1.Fields!IdTipoServicio = 1 Then
                    lbSeImprime = True
                    lcSql = "Cuenta de Consultorios externos, no se ha impreso FICHA FUA"
                 Else
                    If Not IsNull(oRsTmp1.Fields!FechaEgreso) Then
                       lbSeImprime = True
                       lcSql = "Cuenta de Hospitalización/emergencia, tiene ALTA MEDICA pero no se ha impreso FICHA FUA"
                    End If
                 End If
                 If lbSeImprime = True Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!IdCuentaAtencion = oRsTmp1.Fields!IdCuentaAtencion
                    mrs_Tmp.Fields!ApellidoPaterno = oRsTmp1.Fields!ApellidoPaterno
                    mrs_Tmp.Fields!ApellidoMaterno = oRsTmp1.Fields!ApellidoMaterno
                    mrs_Tmp.Fields!PrimerNombre = oRsTmp1.Fields!PrimerNombre
                    mrs_Tmp.Fields!NroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!TipoServicio = oRsTmp1.Fields!TipoServicio
                    mrs_Tmp.Fields!Servicio = oRsTmp1.Fields!Servicio
                    mrs_Tmp.Fields!FechaIngreso = oRsTmp1.Fields!FechaIngreso
                    mrs_Tmp.Fields!FechaEgreso = oRsTmp1.Fields!FechaEgreso
                    mrs_Tmp.Fields!observaciones = lcSql
                    mrs_Tmp.Update
                 End If
              End If
              oRsTmp2.Close
              oRsTmp1.MoveNext
           Loop
        End If
        Set ChequeaCuentasSinFUAgrabado = oRsTmp1
        Exit Function
ErrorProceso1:
        MsgBox Err.Description
        'Resume
End Function


Sub CrearReporteObservaciones_excel(mrs_Tmp As Recordset)
'Dim oExcel As Excel.Application
'Dim oWorkBookPlantilla As Workbook
'Dim oWorkBook As Workbook
'Dim oWorkSheet As Worksheet
'Dim mo_ReporteUtil As New SIGHEntidades.ReporteUtil
'Dim iFila As Long
'Dim lnTotal As Double
'        On Error GoTo ManejadorErrorExcel
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        'Crea nueva hoja
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
'        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'        iFila = 2
'        oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
'        iFila = 4
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
'        oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
'        oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
'        oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
'        oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
'        oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
'        oWorkSheet.Cells(iFila, 8).Value = "Servicio"
'        oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
'        oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
'        oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        iFila = 6: lnTotal = 0
'        mrs_Tmp.MoveFirst
'        Do While Not mrs_Tmp.EOF
'           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
'           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
'           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
'           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
'           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
'           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
'           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
'           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
'           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("FechaEgreso").Value
'           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("observaciones").Value
'           iFila = iFila + 1
'           mrs_Tmp.MoveNext
'        Loop
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
'        oExcel.Visible = True
'        oWorkSheet.PrintPreview
'        'oWorkSheet.PrintOut
'        Set oWorkSheet = Nothing
'        Set oExcel = Nothing
'        Exit Sub
'ManejadorErrorExcel:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'        Resume
'    End Select
'    Exit Sub
Dim mo_ReporteUtil As New sighentidades.ReporteUtil
Dim iFila As Long
Dim lnTotal As Double
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
Dim lnHwnd As Long
lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If

        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
            oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        iFila = 2
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Lista de Cuentas Observadas (No fueron grabados sus FUAS)")
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
        End If
        iFila = 4
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Cuenta")
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Apellido Paterno")
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Apellido Materno")
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula("Primer Nombre")
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula("N° Historia Clinica")
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula("Tipo Servicio")
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula("Servicio")
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Fecha Ingreso")
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula("Fecha Alta")
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula("Observaciones")
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
            oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
            oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
            oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
            oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
            oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
            oWorkSheet.Cells(iFila, 8).Value = "Servicio"
            oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
            oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
            oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
        End If
        iFila = 6: lnTotal = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("idCuentaAtencion").Value)
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("ApellidoPaterno").Value)
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("apellidoMaterno").Value)
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("PrimerNombre").Value)
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp.Fields("tipoServicio").Value)
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("Servicio").Value)
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp.Fields("FechaIngreso").Value)
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula((IIf(IsNull(mrs_Tmp.Fields("FechaEgreso").Value), "", mrs_Tmp.Fields("FechaEgreso").Value)))
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula((IIf(IsNull(mrs_Tmp.Fields("observaciones").Value), "", mrs_Tmp.Fields("observaciones").Value)))
        Else
            oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
            oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
            oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
            oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
            oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
            oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
            oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
            oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
            oWorkSheet.Cells(iFila, 10).Value = (IIf(IsNull(mrs_Tmp.Fields("FechaEgreso").Value), "", mrs_Tmp.Fields("FechaEgreso").Value))
            oWorkSheet.Cells(iFila, 11).Value = (IIf(IsNull(mrs_Tmp.Fields("observaciones").Value), "", mrs_Tmp.Fields("observaciones").Value))
        End If
        iFila = iFila + 1
        mrs_Tmp.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
            oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
        End If
        If lbEsOpenOffice = True Then
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
        End If
            If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            Set oWorkSheet = Nothing
            Set oExcel = Nothing
        End If
        Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select
    Exit Sub
End Sub

'debb
Sub ExportaSIS(lcFechaInicio As String, lcFechaFinal As String, lnHwnd As Long)
    On Error GoTo ErrorProceso

    Dim oDoSisFuaResumen As New SIGHCOMUN.DoSisFuaResumen: Dim oSisFuaResumen As New SIGHDatos.SisFuaResumen
    Dim oDOSisFuaEstadosTrama As New SIGHCOMUN.DoSisFuaEstadosTrama: Dim oSisFuaEstadosTrama As New SIGHDatos.SisFuaEstadosTrama
    Dim wrst As New ADODB.Recordset
    Dim oDoSisFuaUsuario As New SIGHCOMUN.DoSisFuaUsuario: Dim oSisFuaUsuario As New SIGHDatos.SisFuaUsuario
    Dim mo_ReglasComunes As New ReglasComunes
    Dim rstEmpleado As New ADODB.Recordset
    Dim rstUsuario As New ADODB.Recordset
    Dim lniFreeFile As Integer
    Dim lcMes As String: Dim lcAnio As String
    Dim lcTempo As Object
    Dim lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, lcSql As String
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim mrs_Tmp As New Recordset
    Dim lnTotalReg As Long, lbProcesoOK As Boolean
    Dim oConexionExterna As New Connection
    Dim oConexion As New Connection
    Dim lbSeImprime As Boolean, lcRutaExportar As String, lcArchivoExpZip As String
    Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim oCrypKey As New CrypKey.Util
    Dim archivo As String
    
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CursorLocation = adUseClient
    
    Set oRsTmp1 = ChequeaCuentasSinFUAgrabado(mrs_Tmp, oConexionExterna, lcFechaInicio, lcFechaFinal)
    lbSeImprime = True
    lbProcesoOK = False
    If oRsTmp1.RecordCount = 0 Then
        MsgBox "No existen Datos en ese rango de Fechas", vbInformation, "Mensaje"
        lbSeImprime = False
        oRsTmp1.Close
        mrs_Tmp.Close
    ElseIf mrs_Tmp.RecordCount > 0 Then
        CrearReporteObservaciones_excel mrs_Tmp
        If MsgBox("Existen Cuentas que no se han grabado en FUA, desea exportar FUA de todas maneras ?", vbQuestion + vbYesNo, "Mensaje") <> vbYes Then
           lbSeImprime = False
           oRsTmp1.Close
           mrs_Tmp.Close
        End If
    End If
    If lbSeImprime = True Then
        lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
        Set lcTempo = CreateObject("Scripting.FileSystemObject")
        'verificar si existe la carpeta y borrarla
        'If (GetAttr(lcRutaExportar & "EnviosFua") And vbDirectory) = vbDirectory Then
           'lcTempo.deleteFolder lcRutaExportar & "EnviosFua"
        'End If
        lcTempo.CreateFolder lcRutaExportar & "EnviosFua"
        '
        oRsTmp1.MoveFirst
        lnIdCuentaInicial = oRsTmp1.Fields!IdCuentaAtencion
        oRsTmp1.MoveLast
        lnIdCuentaFinal = oRsTmp1.Fields!IdCuentaAtencion
        oRsTmp1.Close
        'SisFuaAtencion
        Set oRsTmp1 = sisFuaAtencionXrangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)
        lnTotalReg = oRsTmp1.RecordCount
        
        If lnTotalReg > 0 Then
            lcMes = Mid(lcBuscaParametro.RetornaFechaServidorSQL, 4, 2) 'obtener mes del servidor
            lcAnio = Mid(lcBuscaParametro.RetornaFechaServidorSQL, 7, 4) 'obtener año del servidor
            oDoSisFuaResumen.CantFilATE = lnTotalReg 'asigno cantidad de atenciones para el resumen
            oDoSisFuaResumen.Anio = lcAnio 'asigno el año del resumen
            oDoSisFuaResumen.Mes = lcMes 'asigno el mes del resumen
            oDoSisFuaResumen.VersionGTI = lcBuscaParametro.SeleccionaFilaParametro(332) 'sacar de parametros la version de gvti
            oDoSisFuaUsuario.Periodo = lcAnio 'asigno el año tomado del servidor al usuario
            oDoSisFuaUsuario.Mes = lcMes 'asigno el mes tomado del servidor al usuario
            oDoSisFuaUsuario.CodigoEstablecimiento = oRsTmp1.Fields!CabOrigenDelRegistro 'asigno el codigo de establecimiento tomado de atenciones
            Set rstEmpleado = mo_ReglasComunes.EmpleadosSeleccionarPorIdEmpleado(ml_idUsuario, oConexion) 'Obtener el dni de usuario desde tabla empleados
            If rstEmpleado.RecordCount > 0 Then
                oDoSisFuaUsuario.DNI = Trim(rstEmpleado.Fields!DNI) 'asigno el dni
                oDoSisFuaUsuario.ApellidoPat = rstEmpleado.Fields!ApellidoPaterno 'asigno el apellido paterno
                oDoSisFuaUsuario.ApellidoMat = rstEmpleado.Fields!ApellidoMaterno 'asigno el apellido materno
                oDoSisFuaUsuario.TipoDoc = rstEmpleado.Fields!idTipoDocumento 'asigno el tipo de documento
                oDoSisFuaUsuario.PrimerNombre = sighentidades.RetornaPrimerNombre(rstEmpleado.Fields!Nombres)    'obtengo y asigno el primer nombre
                oDoSisFuaUsuario.SegundoNombre = sighentidades.RetornaSegundoNombre(rstEmpleado.Fields!Nombres) 'obtengo y asigno el segundo nombre
                
                'jala el nro de envio para el odosisfuausuario.nroenvio
                Set oSisFuaUsuario.Conexion = oConexionExterna 'asigno a conexion
                Set rstUsuario = oSisFuaUsuario.SeleccionarPorDNI(oDoSisFuaUsuario.DNI) 'genero la seleccion de usuarios
                If rstUsuario.RecordCount > 0 Then
                    rstUsuario.MoveLast
                    If Format(rstUsuario.Fields!Mes, "00") = lcMes Then 'comparacion de mes
                        oDoSisFuaUsuario.NroEnvio = rstUsuario.Fields!NroEnvio + 1 'envio del mismo mes, aumenta
                    Else
                        oDoSisFuaUsuario.NroEnvio = 1 'envio en mes diferente vuelve a empezar
                    End If
                Else
                    oDoSisFuaUsuario.NroEnvio = 1 'si es su primera vez
                End If
                oDoSisFuaResumen.NroEnvio = oDoSisFuaUsuario.NroEnvio 'asigno nro de envio obtenido
                'concateno y asigno nombre del paquete (archivo zip)
                oDoSisFuaResumen.NomPaquete = oDoSisFuaUsuario.Periodo + Format(oDoSisFuaUsuario.Mes, "00") + Format(oDoSisFuaUsuario.NroEnvio, "00") + RetornarCodigoUDR + RetornarCodigoPPDD + ".zip"
                
                GeneraTXTUsuario oDoSisFuaUsuario, lcRutaExportar & "EnviosFua\USUARIOS.txt", RetornarCodigoUDR, lcBuscaParametro.SeleccionaFilaParametro(304)   'llamo al procedimiento que hace archivo de texto
                Set oSisFuaUsuario.Conexion = oConexionExterna 'asigno conexion
                If oSisFuaUsuario.Insertar(oDoSisFuaUsuario) = True Then 'inserto Usuario
                    Set rstUsuario = Nothing 'limpio recordset
                End If
            End If
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencion", oConexionExterna) 'cargo recordset para filtro de campos
            'genero archivo de texto pasando el recorset de filtro y el nro de envio
            GeneraTXTAtencion oRsTmp1, lcRutaExportar & "EnviosFua\ATENCION.txt", lnTotalReg, 1, wrst, Format(oDoSisFuaUsuario.NroEnvio, "00"), Format(oDoSisFuaUsuario.Mes, "00"), oDoSisFuaUsuario.Periodo
            wrst.Close
        Else
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCION.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close
        'Llamada a los componentes
        'SisFuaAtencionDIA
        Set oRsTmp1 = sisFuaAtencionDIAxRangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)  'carga de datos de atencionDIA
        lnTotalReg = oRsTmp1.RecordCount
        If lnTotalReg > 0 Then
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencionDIA", oConexionExterna) 'cargo recordset para filtro de campos
            GeneraTXT oRsTmp1, lcRutaExportar & "EnviosFua\ATENCIONDIA.txt", lnTotalReg, 2, wrst, Format(oDoSisFuaResumen.NroEnvio, "00"), Format(oDoSisFuaResumen.Mes, "00"), oDoSisFuaResumen.Anio   'genero texto
            wrst.Close
            oDoSisFuaResumen.CantFilDIA = lnTotalReg 'asigno para resumen la cantidad de registros de atencionDIA
        Else 'si no hay registros creo archivo de texto vacio
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCIONDIA.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close

        'SisFuaAtencionINS
        Set oRsTmp1 = sisFuaAtencionINSxRangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)
        lnTotalReg = oRsTmp1.RecordCount
        If lnTotalReg > 0 Then
            oDoSisFuaResumen.CantFilINS = lnTotalReg
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencionINS", oConexionExterna)
            GeneraTXT oRsTmp1, lcRutaExportar & "EnviosFua\ATENCIONINS.txt", lnTotalReg, 3, wrst, Format(oDoSisFuaResumen.NroEnvio, "00"), Format(oDoSisFuaResumen.Mes, "00"), oDoSisFuaResumen.Anio
            wrst.Close
        Else
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCIONINS.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close
        'SisFuaAtencionMED
        Set oRsTmp1 = sisFuaAtencionMEDxRangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)
        lnTotalReg = oRsTmp1.RecordCount
        If lnTotalReg > 0 Then
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencionMED", oConexionExterna)
            GeneraTXT oRsTmp1, lcRutaExportar & "EnviosFua\ATENCIONMED.txt", lnTotalReg, 4, wrst, Format(oDoSisFuaResumen.NroEnvio, "00"), Format(oDoSisFuaResumen.Mes, "00"), oDoSisFuaResumen.Anio
            wrst.Close
            oDoSisFuaResumen.CantFilMED = lnTotalReg
        Else
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCIONMED.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close
        'SisFuaAtencionPRO
        Set oRsTmp1 = sisFuaAtencionPROxRangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)
        lnTotalReg = oRsTmp1.RecordCount
        If lnTotalReg > 0 Then
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencionPRO", oConexionExterna)
            GeneraTXT oRsTmp1, lcRutaExportar & "EnviosFua\ATENCIONPRO.txt", lnTotalReg, 5, wrst, Format(oDoSisFuaResumen.NroEnvio, "00"), Format(oDoSisFuaResumen.Mes, "00"), oDoSisFuaResumen.Anio
            wrst.Close
            oDoSisFuaResumen.CantFilPRO = lnTotalReg
        Else
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCIONPRO.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close
        'SisFuaAtencionSMI
        Set oRsTmp1 = sisFuaAtencionSMIxRangoDeCuentas(lnIdCuentaInicial, lnIdCuentaFinal, oConexionExterna)
        lnTotalReg = oRsTmp1.RecordCount
        If lnTotalReg > 0 Then
            Set wrst = oSisFuaEstadosTrama.SeleccionarPorTabla("sisFuaAtencionSMI", oConexionExterna)
            GeneraTXT oRsTmp1, lcRutaExportar & "EnviosFua\ATENCIONSMI.txt", lnTotalReg, 6, wrst, Format(oDoSisFuaResumen.NroEnvio, "00"), Format(oDoSisFuaResumen.Mes, "00"), oDoSisFuaResumen.Anio
            wrst.Close
            oDoSisFuaResumen.CantFilSMI = lnTotalReg
        Else
            lniFreeFile = FreeFile
            Open lcRutaExportar & "EnviosFua\ATENCIONSMI.txt" For Output As #lniFreeFile
            Close #lniFreeFile
        End If
        oRsTmp1.Close
        
        'AFILIACION
        lniFreeFile = FreeFile
        Open lcRutaExportar & "EnviosFua\AFILIACION.txt" For Output As #lniFreeFile
        Close #lniFreeFile
        
        'FILIACION
        lniFreeFile = FreeFile
        Open lcRutaExportar & "EnviosFua\FILIACION.txt" For Output As #lniFreeFile
        Close #lniFreeFile
        
        'INSCRIPCION
        lniFreeFile = FreeFile
        Open lcRutaExportar & "EnviosFua\INSCRIPCION.txt" For Output As #lniFreeFile
        Close #lniFreeFile
        
        'RESPAPLICACION
        lniFreeFile = FreeFile
        Open lcRutaExportar & "EnviosFua\RESPAPLICACION.txt" For Output As #lniFreeFile
        Close #lniFreeFile
        
        
      ' fin de llamada a componentes
        Set oSisFuaResumen.Conexion = oConexionExterna 'asigno conexion
        oDoSisFuaResumen.CantFilUSU = 1 'asigno cantidad de usuarios para resumen
        If oSisFuaResumen.Insertar(oDoSisFuaResumen) Then 'inserto resumen
            Set rstEmpleado = Nothing 'limpio recordset
        End If
        Call GeneraTXTResumen(oDoSisFuaResumen, lcRutaExportar & "EnviosFua\RESUMEN.txt", RetornarCodigoUDR, lcBuscaParametro.SeleccionaFilaParametro(304))  'genero archivo de texto
        
        'AQUI EL CODIGO DE LAS DEMAS
        lbProcesoOK = True
        Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "FUA", oConexion, 500 + 180, mo_lcNombrePc, "Exportar al SIS: " & lcFechaInicio & " al " & lcFechaFinal)
    End If
    Set oRsTmp1 = Nothing
    Set mrs_Tmp = Nothing
    Set oConexionExterna = Nothing
    If lbProcesoOK = True Then
        lcArchivoExpZip = lcRutaExportar + oDoSisFuaResumen.NomPaquete 'modifico el nombre del archivo zip
        sighentidades.ComprimeArchivoZip lcArchivoExpZip, oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(319)), lcRutaExportar & "EnviosFua\*.txt"
        MsgBox "Se creó el archivo: " & lcArchivoExpZip
        lcTempo.deleteFolder lcRutaExportar & "EnviosFua" 'elimino carpeta de creacion de textos
    End If

    Set oCrypKey = Nothing
Exit Sub
ErrorProceso:
MsgBox Err.Description

End Sub

Function FuaAgregar(oDoSisFuaAtencion As DoSisFuaAtencion, oRsVacunasSp As Recordset, _
                    oRsCpt As Recordset, oRsFarmacia As Recordset, oRsDx As Recordset, _
                    lcInsumo As String, lcMedicamento As String, _
                    lnIdUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                    lcObservacion As String, lcParametro320 As String, lnIdAtencion As Long, ByRef lbNroFuaRepetido As Boolean) As Boolean
    On Error GoTo ErrFuaAgr
    Dim oConexionExterna As New ADODB.Connection
    Dim oConexion As New ADODB.Connection
    Dim oSisFuaAtencion As New SisFuaAtencion
    Dim oDoSisFuaAtencionDIA As New DoSisFuaAtencionDIA: Dim oSisFuaAtencionDIA As New SisFuaAtencionDIA
    Dim oDoSisFuaAtencionINS As New DoSisFuaAtencionINS: Dim oSisFuaAtencionINS As New SisFuaAtencionINS
    Dim oDoSisFuaAtencionMED As New DoSisFuaAtencionMED: Dim oSisFuaAtencionMED As New SisFuaAtencionMED
    Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO: Dim oSisFuaAtencionPRO As New SisFuaAtencionPRO
    Dim oDoSisFuaAtencionSMI As New DoSisFuaAtencionSMI: Dim oSisFuaAtencionSMI As New SisFuaAtencionSMI
    Dim oDoAtencionDatosAdicionales As New DoAtencionDatosAdicionales, oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
    Dim oDoSisFua As New DoSisFua, oSisFuaCl As New SisFuaCl
    Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset, lcSql As String
    FuaAgregar = False
    '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    '
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.BeginTrans
    '
    Set oSisFuaAtencion.Conexion = oConexionExterna
    Set oSisFuaAtencionDIA.Conexion = oConexionExterna
    Set oSisFuaAtencionINS.Conexion = oConexionExterna
    Set oSisFuaAtencionMED.Conexion = oConexionExterna
    Set oSisFuaAtencionPRO.Conexion = oConexionExterna
    Set oSisFuaAtencionSMI.Conexion = oConexionExterna
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    Set oSisFuaCl.Conexion = oConexionExterna
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "A", oDoSisFuaAtencion.IdCuentaAtencion, "FUA", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)
    '
    If Not oSisFuaCl.SeleccionarTodos(oDoSisFua) Then
       ms_MensajeError = "Correlativos FUA: " & oSisFuaCl.MensajeError: GoTo ErrFuaAgr
    End If
'    If lcParametro320 = sghFuaTipo.sghFuaTipoAutomatico And lnIdAtencion > 0 Then
'       '*************************FUA automatica********************
'       Dim lcCodigoRenaes As String, lcfecha As String, lcCorrelativo As String
'       lcCodigoRenaes = Right("000000" & lcBuscaParametro.SeleccionaFilaParametro(280), 6)
'       lcfecha = lcBuscaParametro.RetornaFechaServidorSQL_AAMMDD
'       lcCorrelativo = "0001"
'       If lcfecha = Mid(oDoSisFua.FuaUltimoGenerado, 7, 6) Then
'          lcCorrelativo = Right("000" & Trim(Str(Val(Right(oDoSisFua.FuaUltimoGenerado, 4)) + 1)), 4)
'       End If
'       oDoSisFua.FuaUltimoGenerado = lcCodigoRenaes & lcfecha & lcCorrelativo
'       If Not oSisFuaCl.Modificar(oDoSisFua) Then
'           ms_MensajeError = oSisFuaCl.MensajeError: GoTo ErrFuaAgr
'       End If
'       oDoSisFuaAtencion.FuaNumero = oDoSisFua.FuaUltimoGenerado
'       If oDoSisFuaAtencion.idCuentaAtencion = 999999999 Then
'           oDoSisFuaAtencion.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'           oDoSisFuaAtencionDIA.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'           oDoSisFuaAtencionINS.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'           oDoSisFuaAtencionMED.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'           oDoSisFuaAtencionSMI.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'           oDoSisFuaAtencionPRO.idCuentaAtencion = Val(Right(oDoSisFua.FuaUltimoGenerado, 9))
'        End If
'    Else
        '*************************FUA manual********************
        If Not (Val(oDoSisFuaAtencion.FuaNumero) >= Val(oDoSisFua.FuaNumeroInicial) And Val(oDoSisFuaAtencion.FuaNumero) <= Val(oDoSisFua.FuaNumeroFinal)) Then
           ms_MensajeError = "El N° Formato FUA registrado debe estar en el RANGO de fuas asignado por la DISA " & Chr(13) & "Se le ha generado un N° Formato FUA disponible, si desea editar el RANGO use la opción 'Herramientas --> Exporta/importa SIS'": GoTo ErrFuaAgr
        Else
            'Frank SIS FUA - 23 07
            oDoSisFua.FuaUltimoGenerado = oDoSisFuaAtencion.FuaNumero
            If Not oSisFuaCl.Modificar(oDoSisFua) Then
                ms_MensajeError = oSisFuaCl.MensajeError: GoTo ErrFuaAgr
            End If
            If lnIdAtencion > 0 Then
                oDoSisFuaAtencion.FuaNumero = Right("00000000" & oDoSisFua.FuaUltimoGenerado, 8)
                If oDoSisFuaAtencion.IdCuentaAtencion = 999999999 Then
                    oDoSisFuaAtencion.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                    oDoSisFuaAtencionDIA.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                    oDoSisFuaAtencionINS.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                    oDoSisFuaAtencionMED.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                    oDoSisFuaAtencionSMI.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                    oDoSisFuaAtencionPRO.IdCuentaAtencion = Val(Right("000000000" & oDoSisFua.FuaUltimoGenerado, 9))
                End If
            End If
         End If
    Set oRsTmp2 = sisFuaAtencionXnroAfiliacion(oDoSisFuaAtencion.FuaDisa, oDoSisFuaAtencion.FuaLote, oDoSisFuaAtencion.FuaNumero, oConexionExterna)
    If oRsTmp2.RecordCount > 0 Then
       oRsTmp2.Close
       lbNroFuaRepetido = True
       ms_MensajeError = "El N° Formato FUA ya existe (" & oDoSisFuaAtencion.FuaDisa & "-" & _
                         oDoSisFuaAtencion.FuaLote & "-" & oDoSisFuaAtencion.FuaNumero & ")" & Chr(13) & _
                         "Puede usar el siguiente Nº Formato FUA ó" & Chr(13) & _
                         "Usar la opción 'Herramientas --> Exporta/importa SIS'": GoTo ErrFuaAgr
    Else
       oRsTmp2.Close
    End If
    '
    If Not oSisFuaAtencion.Insertar(oDoSisFuaAtencion) Then
         ms_MensajeError = "Cabecera: " & oSisFuaAtencion.MensajeError: GoTo ErrFuaAgr
    End If
    
    'Dx,Insumos,Medicamentos,Vacunas, Dx
    FUAgrabaServIntermEnTablas oDoSisFuaAtencion, oSisFuaAtencionDIA, oDoSisFuaAtencionDIA, oRsDx, _
                          oSisFuaAtencionINS, oDoSisFuaAtencionINS, oRsFarmacia, oSisFuaAtencionMED, oDoSisFuaAtencionMED, _
                          oSisFuaAtencionSMI, oDoSisFuaAtencionSMI, oRsVacunasSp, _
                          oSisFuaAtencionPRO, oDoSisFuaAtencionPRO, oRsCpt, sghAgregar, lcInsumo, lcMedicamento
                          
    'Actualiza DAtos Adicionales
    If lnIdAtencion > 0 Then
        oDoAtencionDatosAdicionales.idAtencion = lnIdAtencion
        oDoAtencionDatosAdicionales.IdUsuarioAuditoria = oDoSisFuaAtencion.CabDniUsuarioRegistra
        If Not oAtencionesDatosAdicionales.SeleccionarPorId(oDoAtencionDatosAdicionales) Then
           ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaAgr
        End If
        oDoAtencionDatosAdicionales.FuaCodigoPrestacion = oDoSisFuaAtencion.FuaCodigoPrestacion
        If oDoAtencionDatosAdicionales.idAtencion = 0 Then
           If Not oAtencionesDatosAdicionales.Insertar(oDoAtencionDatosAdicionales) Then
              ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaAgr
           End If
        Else
           If Not oAtencionesDatosAdicionales.Modificar(oDoAtencionDatosAdicionales) Then
              ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaAgr
           End If
        End If
    End If
    '
    FuaAgregar = True
ErrFuaAgr:
'Resume
    If FuaAgregar = True Then
       oConexionExterna.CommitTrans
       oConexion.CommitTrans
    Else
'Resume
       oConexionExterna.RollbackTrans
       oConexion.RollbackTrans
    End If
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oSisFuaAtencion = Nothing
    Set oDoSisFuaAtencionDIA = Nothing: Set oSisFuaAtencionDIA = Nothing
    Set oDoSisFuaAtencionINS = Nothing: Set oSisFuaAtencionINS = Nothing
    Set oDoSisFuaAtencionMED = Nothing: Set oSisFuaAtencionMED = Nothing
    Set oDoSisFuaAtencionPRO = Nothing: Set oSisFuaAtencionPRO = Nothing
    Set oDoSisFuaAtencionSMI = Nothing: Set oSisFuaAtencionSMI = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oDoAtencionDatosAdicionales = Nothing: Set oAtencionesDatosAdicionales = Nothing
End Function

Function FuaModificar(oDoSisFuaAtencion As DoSisFuaAtencion, oRsVacunasSp As Recordset, _
                    oRsCpt As Recordset, oRsFarmacia As Recordset, oRsDx As Recordset, _
                    lcInsumo As String, lcMedicamento As String, _
                    lnIdUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                    lcObservacion As String, lnIdAtencion As Long) As Boolean
    On Error GoTo ErrFuaMod
    Dim oConexionExterna As New ADODB.Connection
    Dim oConexion As New ADODB.Connection
    Dim oSisFuaAtencion As New SisFuaAtencion
    Dim oDoSisFuaAtencionDIA As New DoSisFuaAtencionDIA: Dim oSisFuaAtencionDIA As New SisFuaAtencionDIA
    Dim oDoSisFuaAtencionINS As New DoSisFuaAtencionINS: Dim oSisFuaAtencionINS As New SisFuaAtencionINS
    Dim oDoSisFuaAtencionMED As New DoSisFuaAtencionMED: Dim oSisFuaAtencionMED As New SisFuaAtencionMED
    Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO: Dim oSisFuaAtencionPRO As New SisFuaAtencionPRO
    Dim oDoSisFuaAtencionSMI As New DoSisFuaAtencionSMI: Dim oSisFuaAtencionSMI As New SisFuaAtencionSMI
    Dim oDoAtencionDatosAdicionales As New DoAtencionDatosAdicionales, oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
    FuaModificar = False
    '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    '
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.BeginTrans
    '
    Set oSisFuaAtencion.Conexion = oConexionExterna
    Set oSisFuaAtencionDIA.Conexion = oConexionExterna
    Set oSisFuaAtencionINS.Conexion = oConexionExterna
    Set oSisFuaAtencionMED.Conexion = oConexionExterna
    Set oSisFuaAtencionPRO.Conexion = oConexionExterna
    Set oSisFuaAtencionSMI.Conexion = oConexionExterna
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "M", oDoSisFuaAtencion.IdCuentaAtencion, "FUA", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)
    '
    If Not oSisFuaAtencion.Modificar(oDoSisFuaAtencion) Then
         ms_MensajeError = "Cabecera: " & oSisFuaAtencion.MensajeError: GoTo ErrFuaMod
    End If
    'Dx,Insumos,Medicamentos,Vacunas, Dx
    FUAgrabaServIntermEnTablas oDoSisFuaAtencion, oSisFuaAtencionDIA, oDoSisFuaAtencionDIA, oRsDx, _
                          oSisFuaAtencionINS, oDoSisFuaAtencionINS, oRsFarmacia, oSisFuaAtencionMED, oDoSisFuaAtencionMED, _
                          oSisFuaAtencionSMI, oDoSisFuaAtencionSMI, oRsVacunasSp, _
                          oSisFuaAtencionPRO, oDoSisFuaAtencionPRO, oRsCpt, sghModificar, lcInsumo, lcMedicamento
    'Actualiza DAtos Adicionales
    If lnIdAtencion > 0 Then
        oDoAtencionDatosAdicionales.idAtencion = lnIdAtencion
        oDoAtencionDatosAdicionales.IdUsuarioAuditoria = oDoSisFuaAtencion.CabDniUsuarioRegistra
        If Not oAtencionesDatosAdicionales.SeleccionarPorId(oDoAtencionDatosAdicionales) Then
           ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaMod
        End If
        oDoAtencionDatosAdicionales.FuaCodigoPrestacion = oDoSisFuaAtencion.FuaCodigoPrestacion
        If oDoAtencionDatosAdicionales.idAtencion = 0 Then
           If Not oAtencionesDatosAdicionales.Insertar(oDoAtencionDatosAdicionales) Then
              ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaMod
           End If
        Else
           If Not oAtencionesDatosAdicionales.Modificar(oDoAtencionDatosAdicionales) Then
              ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaMod
           End If
        End If
    End If
    '
    FuaModificar = True
ErrFuaMod:
'Resume
    If FuaModificar = True Then
       oConexionExterna.CommitTrans
       oConexion.CommitTrans
    Else
       oConexionExterna.RollbackTrans
       oConexion.RollbackTrans
    End If
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oSisFuaAtencion = Nothing
    Set oDoSisFuaAtencionDIA = Nothing: Set oSisFuaAtencionDIA = Nothing
    Set oDoSisFuaAtencionINS = Nothing: Set oSisFuaAtencionINS = Nothing
    Set oDoSisFuaAtencionMED = Nothing: Set oSisFuaAtencionMED = Nothing
    Set oDoSisFuaAtencionPRO = Nothing: Set oSisFuaAtencionPRO = Nothing
    Set oDoSisFuaAtencionSMI = Nothing: Set oSisFuaAtencionSMI = Nothing
    Set oDoAtencionDatosAdicionales = Nothing: Set oAtencionesDatosAdicionales = Nothing
End Function

Function FuaEliminar(oDoSisFuaAtencion, _
                    lnIdUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, _
                    lcObservacion As String, lnIdAtencion As Long) As Boolean
    On Error GoTo ErrFuaEli
    Dim oConexionExterna As New ADODB.Connection
    Dim oConexion As New ADODB.Connection
    Dim oSisFuaAtencion As New SisFuaAtencion
    Dim oDoSisFuaAtencionDIA As New DoSisFuaAtencionDIA: Dim oSisFuaAtencionDIA As New SisFuaAtencionDIA
    Dim oDoSisFuaAtencionINS As New DoSisFuaAtencionINS: Dim oSisFuaAtencionINS As New SisFuaAtencionINS
    Dim oDoSisFuaAtencionMED As New DoSisFuaAtencionMED: Dim oSisFuaAtencionMED As New SisFuaAtencionMED
    Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO: Dim oSisFuaAtencionPRO As New SisFuaAtencionPRO
    Dim oDoSisFuaAtencionSMI As New DoSisFuaAtencionSMI: Dim oSisFuaAtencionSMI As New SisFuaAtencionSMI
    Dim oDoAtencionDatosAdicionales As New DoAtencionDatosAdicionales, oAtencionesDatosAdicionales As New AtencionesDatosAdicionales
    FuaEliminar = False
    '
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    '
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.BeginTrans
    '
    Set oSisFuaAtencion.Conexion = oConexionExterna
    Set oSisFuaAtencionDIA.Conexion = oConexionExterna
    Set oSisFuaAtencionINS.Conexion = oConexionExterna
    Set oSisFuaAtencionMED.Conexion = oConexionExterna
    Set oSisFuaAtencionPRO.Conexion = oConexionExterna
    Set oSisFuaAtencionSMI.Conexion = oConexionExterna
    Set oAtencionesDatosAdicionales.Conexion = oConexion
    '
    
    Call mo_ReglasSeguridad.AuditoriaAgregarV(lnIdUsuario, "E", oDoSisFuaAtencion.IdCuentaAtencion, "FUA", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcObservacion)    '
    'Dx
    
    If oSisFuaAtencionDIA.SisFuaAtencionDIAeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
       ms_MensajeError = "DIA: " & oSisFuaAtencionDIA.MensajeError: GoTo ErrFuaEli
    End If
    'Insumos
    If oSisFuaAtencionINS.SisFuaAtencionINSeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
       ms_MensajeError = "INS: " & oSisFuaAtencionINS.MensajeError: GoTo ErrFuaEli
    End If
    'Medicamentos
    If oSisFuaAtencionMED.SisFuaAtencionMEDeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
       ms_MensajeError = "INS: " & oSisFuaAtencionMED.MensajeError
       GoTo ErrFuaEli
    End If
    'SMI - Vacunas y  SP
    If oSisFuaAtencionSMI.SisFuaAtencionSMIeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
       ms_MensajeError = "SMI: " & oSisFuaAtencionSMI.MensajeError: GoTo ErrFuaEli
    End If
    'Cpt
    If oSisFuaAtencionPRO.SisFuaAtencionPROeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
       ms_MensajeError = "CPT: " & oSisFuaAtencionPRO.MensajeError: GoTo ErrFuaEli
    End If
    '
    If Not oSisFuaAtencion.Eliminar(oDoSisFuaAtencion) Then
         ms_MensajeError = "Cabecera: " & oSisFuaAtencion.MensajeError: GoTo ErrFuaEli
    End If
    'Actualiza DAtos Adicionales
    If lnIdAtencion > 0 Then
        oDoAtencionDatosAdicionales.idAtencion = lnIdAtencion
        oDoAtencionDatosAdicionales.IdUsuarioAuditoria = oDoSisFuaAtencion.IdUsuarioAuditoria
        If Not oAtencionesDatosAdicionales.SeleccionarPorId(oDoAtencionDatosAdicionales) Then
           ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaEli
        End If
        oDoAtencionDatosAdicionales.FuaCodigoPrestacion = ""
        If Not oAtencionesDatosAdicionales.Modificar(oDoAtencionDatosAdicionales) Then
           ms_MensajeError = "ATENCIONES DATOS ADICIONALES: " & oAtencionesDatosAdicionales.MensajeError: GoTo ErrFuaEli
        End If
    End If
    '
    FuaEliminar = True
ErrFuaEli:
'Resume
    If FuaEliminar = True Then
       oConexionExterna.CommitTrans
       oConexion.CommitTrans
    Else
       oConexionExterna.RollbackTrans
       oConexion.RollbackTrans
    End If
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oSisFuaAtencion = Nothing
    Set oDoSisFuaAtencionDIA = Nothing: Set oSisFuaAtencionDIA = Nothing
    Set oDoSisFuaAtencionINS = Nothing: Set oSisFuaAtencionINS = Nothing
    Set oDoSisFuaAtencionMED = Nothing: Set oSisFuaAtencionMED = Nothing
    Set oDoSisFuaAtencionPRO = Nothing: Set oSisFuaAtencionPRO = Nothing
    Set oDoSisFuaAtencionSMI = Nothing: Set oSisFuaAtencionSMI = Nothing
    Set oDoAtencionDatosAdicionales = Nothing: Set oAtencionesDatosAdicionales = Nothing
End Function





Function Sis_ValidaSiEsAfiliadoActualDelSIS(oRecordset As Recordset, ldFechaActual As Date) As Boolean
        Dim lcSql As String
        Sis_ValidaSiEsAfiliadoActualDelSIS = False
        
        If IsNull(oRecordset.Fields!fbajaOK) Then
            If Val(oRecordset.Fields!estadoSis) = 0 Then
                Sis_ValidaSiEsAfiliadoActualDelSIS = True
            Else
                lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
                        "Motivo de Baja: " & IIf(IsNull(oRecordset.Fields!MotivoBaja), "", oRecordset.Fields!MotivoBaja) & Chr(13) & _
                        "Estado: " & IIf(oRecordset.Fields!estadoSis = 0, "Activo", "Inactivo")
                MsgBox lcSql, vbExclamation, "SIS"
            End If
        Else
            If ldFechaActual <= oRecordset.Fields!fbajaOK Then
                If Val(oRecordset.Fields!estadoSis) = 0 Then
                    Sis_ValidaSiEsAfiliadoActualDelSIS = True
                Else
                    lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
                            "Motivo de Baja: " & IIf(IsNull(oRecordset.Fields!MotivoBaja), "", oRecordset.Fields!MotivoBaja) & Chr(13) & _
                            "Fecha Baja: " & oRecordset.Fields!fbajaOK & Chr(13) & _
                            "Estado: " & IIf(oRecordset.Fields!estadoSis = 0, "Activo", "Inactivo")
                    MsgBox lcSql, vbExclamation, "SIS"
                End If
            Else
                lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
                        "Motivo de Baja: " & IIf(IsNull(oRecordset.Fields!MotivoBaja), "", oRecordset.Fields!MotivoBaja) & Chr(13) & _
                        "Fecha Baja: " & oRecordset.Fields!fbajaOK & Chr(13) & _
                        "Estado: " & IIf(oRecordset.Fields!estadoSis = 0, "Activo", "Inactivo")
                MsgBox lcSql, vbExclamation, "SIS"
            End If
        End If
        
        
'        If (IsNull(oRecordset.Fields!fbajaOK) Or (ldFechaActual <= oRecordset.Fields!fbajaOK)) And Val(oRecordset.Fields!estadoSis) = 0 Then
'           Sis_ValidaSiEsAfiliadoActualDelSIS = True
'        Else
'            lcSql = "La afiliación de este paciente tiene problemas: " & Chr(13) & Chr(13) & _
'                    "Motivo de Baja: " & IIf(IsNull(oRecordset.Fields!MotivoBaja), "", oRecordset.Fields!MotivoBaja) & Chr(13) & _
'                    "Fecha Baja: " & oRecordset.Fields!fbajaOK & Chr(13) & _
'                    "Estado: " & IIf(oRecordset.Fields!estadoSis = 0, "Activo", "Inactivo")
'            MsgBox lcSql, vbExclamation, "SIS"
'        End If
End Function



'debb
Function Sis_a_categoriaeessDevuelveNivel(lcCodigoCategoriaSis As String) As String

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "a_categoriaeessSeleccionarPorId"
            Set oParameter = .CreateParameter("@lcCodigoCategoriaSis", adVarChar, adParamInput, 2, lcCodigoCategoriaSis): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Sis_a_categoriaeessDevuelveNivel = ""
    If oRecordset.RecordCount > 0 Then
       Sis_a_categoriaeessDevuelveNivel = Trim(oRecordset.Fields!cat_nivel)
    End If
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

'debb
Sub Sis_FuasRegistradasXpacienteSegunCodigoPrestacion(ldFechaAtencionFuaActual As Date, lnIdCuentaAtencion As Long, _
                                                      lcCodigoPrestacion As String, lcNroHistoriaClinica As String, _
                                                      ByRef lnFuasDelDia As Integer, ByRef lnFuasDelMes As Integer, _
                                                      ByRef lnFuasDelAnio As Integer)
    lnFuasDelDia = 0: lnFuasDelMes = 0: lnFuasDelAnio = 0
    Dim lcSql As String, oRsTmp1 As New Recordset, ldFechaFuaHistorica As Date
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionXcuentaHistoriaPrestacion"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FuaNroHistoria", adVarChar, adParamInput, 20, Trim(lcNroHistoriaClinica)):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FuaCodigoPrestacion", adVarChar, adParamInput, 3, lcCodigoPrestacion):   .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
    End With
    '
    If oRsTmp1.RecordCount > 0 Then
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          ldFechaFuaHistorica = CDate(DevuelveFechaSegunFormato_YMD_SIS(oRsTmp1.Fields!FuaAtencionFecha))
          If ldFechaFuaHistorica = ldFechaAtencionFuaActual Then
             lnFuasDelDia = lnFuasDelDia + 1
             lnFuasDelMes = lnFuasDelMes + 1
             lnFuasDelAnio = lnFuasDelAnio + 1
          ElseIf Month(ldFechaFuaHistorica) = Month(ldFechaAtencionFuaActual) And Year(ldFechaFuaHistorica) = Year(ldFechaAtencionFuaActual) Then
             lnFuasDelMes = lnFuasDelMes + 1
             lnFuasDelAnio = lnFuasDelAnio + 1
          ElseIf Year(ldFechaFuaHistorica) = Year(ldFechaAtencionFuaActual) Then
             lnFuasDelAnio = lnFuasDelAnio + 1
          End If
          oRsTmp1.MoveNext
       Loop
    End If
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
End Sub

Function Sis_SMIregistradasXpaciente(lcCodigoSMI As String, lcNroHistoriaClinica As String, lnIdCuentaAtencion As Long) As Integer

On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SMIregistradasXpaciente"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FuaNroHistoria", adVarChar, adParamInput, 20, lcNroHistoriaClinica):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IntervencionesPreventivas", adVarChar, adParamInput, 3, lcCodigoSMI):   .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

'debb
Function SisFuaAtencionDIAxIdCuentaAtencion(ml_IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionDIAxIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionDIAxIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
  End Function


Function SisFuaAtencionSMIxIdCuentaAtencion(ml_IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionSMIxIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionSMIxIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

Function SisFuaAtencionMEDxIdCuentaAtencion(ml_IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionMEDxIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionMEDxIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'debb
Function SisFuaAtencionINSxIdCuentaAtencion(ml_IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionINSxIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionINSxIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

 End Function

'debb
Function SisFuaAtencionPROxIdCuentaAtencion(ml_IdCuentaAtencion As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionPROxIdCuentaAtencion"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, ml_IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionPROxIdCuentaAtencion = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

Function SisFiliacionesSeleccionarPorIdSiaSis(lnIdSiaSis As Long, lcCodigo As String, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set SisFiliacionesSeleccionarPorIdSiaSis = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFiliacionesSeleccionarPorIdSiaSis"
        Set oParameter = .CreateParameter("@idSiasis", adInteger, adParamInput, 0, lnIdSiaSis): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 2, lcCodigo):   .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set SisFiliacionesSeleccionarPorIdSiaSis = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


'debb-24/08/2015
Function SisFuaActualizaDatos(lcDisa As String, lcLote As String, lcNumeroInicial As String, _
                             lcNumeroFinal As String, lcNumeroUltimoGenerado As String, _
                             lnIdUsoEnGalenhos As Integer, lcTieneInternet As String, _
                             lcNuevoFormatoFua As String, lcFUA2015anexo As String) As Boolean
    On Error GoTo ErrAD
    ms_MensajeError = ""
    SisFuaActualizaDatos = False
    Dim lcSql As String, oRsTmp1 As New Recordset, lnFor As Long
    Dim oConexion As New Connection
    Dim oConexion1 As New Connection
    Dim oDoSisFua As New DoSisFua, oSisFuaCl As New SisFuaCl
    Dim oDOPArametro As New DOPArametro, oParametros As New Parametros
    Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
    Dim mo_Servicios  As New Servicios, mo_doServicio As New doServicio
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexion.CursorLocation = adUseClient
    Set oSisFuaCl.Conexion = oConexion
    If Not oSisFuaCl.SeleccionarTodos(oDoSisFua) Then
       ms_MensajeError = oSisFuaCl.MensajeError: GoTo ErrAD
    End If
    oDoSisFua.FuaDisa = lcDisa
    oDoSisFua.FuaLote = lcLote
    oDoSisFua.FuaNumeroInicial = lcNumeroInicial
    oDoSisFua.FuaNumeroFinal = lcNumeroFinal
    
    'Frank 29102014
    oDoSisFua.FuaUltimoGenerado = lcNumeroUltimoGenerado
'    If lnIdUsoEnGalenhos = 2 Then
'       oDoSisFua.FuaUltimoGenerado = mo_ReglasSISgalenhos.sisFuaAtencionUltimoCorrelativoAutomatico(oConexion)
'    Else
'       oDoSisFua.FuaUltimoGenerado = lcNumeroUltimoGenerado
'    End If

    If Not oSisFuaCl.Modificar(oDoSisFua) Then
       ms_MensajeError = oSisFuaCl.MensajeError: GoTo ErrAD
    End If
    '
    oConexion1.CommandTimeout = 300
    oConexion1.Open sighentidades.CadenaConexion
    oConexion1.CursorLocation = adUseClient
    Set oParametros.Conexion = oConexion1
    '
    oDOPArametro.IdParametro = 320  'Nro fua: manual/automatico
    If Not oParametros.SeleccionarPorId(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    oDOPArametro.ValorTexto = lnIdUsoEnGalenhos
    If Not oParametros.Modificar(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    '
    oDOPArametro.IdParametro = 322  'tiene WEB para SIS
    If Not oParametros.SeleccionarPorId(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    oDOPArametro.ValorTexto = lcTieneInternet
    If Not oParametros.Modificar(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    '
    oDOPArametro.IdParametro = 358     'formato FUA a usar: nuevo2015 o actual
    If Not oParametros.SeleccionarPorId(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    oDOPArametro.ValorTexto = lcNuevoFormatoFua
    If Not oParametros.Modificar(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    '
    oDOPArametro.IdParametro = 359      'Tipo Anexo Formato 2015
    If Not oParametros.SeleccionarPorId(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    oDOPArametro.ValorTexto = lcFUA2015anexo
    If Not oParametros.Modificar(oDOPArametro) Then
       ms_MensajeError = oParametros.MensajeError: GoTo ErrAD
    End If
    '
    Set mo_Servicios.Conexion = oConexion1
    
    For lnFor = 1 To 3
        Set oRsTmp1 = mo_Servicios.SeleccionarPorTipo(lnFor)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
                mo_doServicio.IdServicio = oRsTmp1!IdServicio
                If mo_Servicios.SeleccionarPorId(mo_doServicio) = True Then
                    'debb-21/09/2015 (inicio)
'                   If lcNuevoFormatoFua = "A" Then
'                       mo_DOServicio.FuaTipoAnexo2015 = 0
'                   ElseIf Val(lcFUA2015anexo) = 3 Then
'                       mo_DOServicio.FuaTipoAnexo2015 = 1
'                   Else
'                       mo_DOServicio.FuaTipoAnexo2015 = Val(lcFUA2015anexo)
'                   End If
                   If lcNuevoFormatoFua = "A" Then
                       mo_doServicio.FuaTipoAnexo2015 = 0
                   Else
                       mo_doServicio.FuaTipoAnexo2015 = Val(lcFUA2015anexo)
                   End If
                   'debb-21/09/2015 (fin)
                   If mo_Servicios.Modificar(mo_doServicio) = True Then
                   End If
                End If
                oRsTmp1.MoveNext
           Loop
       End If
       oRsTmp1.Close
    Next
    '
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    oConexion1.Close
    Set oConexion1 = Nothing
    SisFuaActualizaDatos = True
    Set mo_ReglasSISgalenhos = Nothing
    Set mo_Servicios = Nothing
    Set mo_doServicio = Nothing
    Exit Function
ErrAD:
    ms_MensajeError = Err.Description
End Function

'debb
Function SisTiposDocumentosSeleccionarTodos() As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "a_tipoDocumentoSeleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set SisTiposDocumentosSeleccionarTodos = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

'debb
Function SisServiciosSeleccionarPorFiltro(lcFiltro As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim orstmp As New Recordset, lcSql As String, oRsTmp1 As New Recordset
    Dim lcDescripcion As String, lcServicio As String
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "SisServiciosSeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set orstmp = .Execute
            Set orstmp.ActiveConnection = Nothing
    End With
    With oRsTmp1
          .Fields.Append "ser_IdServicio", adVarChar, 3, adFldIsNullable
          .Fields.Append "ser_Descripcion", adVarChar, 100, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    If orstmp.RecordCount > 0 Then
       orstmp.MoveFirst
       Do While Not orstmp.EOF
          lcServicio = orstmp.Fields!dServicioCodigo
          lcDescripcion = orstmp.Fields!dservicio
          Do While Not orstmp.EOF And lcServicio = orstmp.Fields!dServicioCodigo
             orstmp.MoveNext
             If orstmp.EOF Then
                Exit Do
             End If
          Loop
          oRsTmp1.AddNew
          oRsTmp1.Fields!ser_idServicio = lcServicio
          oRsTmp1.Fields!ser_descripcion = lcDescripcion
          oRsTmp1.Update
       Loop
       oRsTmp1.Sort = "ser_Descripcion"
    End If
    orstmp.Close
    Set SisServiciosSeleccionarPorFiltro = oRsTmp1
    Set orstmp = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function
'debb
Function SisServiciosDevuelveDescripcion(lcIdServicio As String) As String
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_serviciosXcodigo"
            Set oParameter = .CreateParameter("@lcIdServicio", adVarChar, adParamInput, 3, lcIdServicio): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    SisServiciosDevuelveDescripcion = ""
    If oRecordset.RecordCount > 0 Then
       SisServiciosDevuelveDescripcion = oRecordset.Fields!ser_descripcion
    End If
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

Function SisConceptoPrestacionalSeleccionarTodos(lbEsVersion2 As Boolean) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "a_modalidadAtencionSeleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    If lbEsVersion2 = True Then
        oRecordset.Filter = "mod_IdEstado=0"
    End If
    Set SisConceptoPrestacionalSeleccionarTodos = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

'debb
Function SisDestinoAtencionSeleccionarTodos() As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "a_destinoAseguradoSeleccionarTodos"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set SisDestinoAtencionSeleccionarTodos = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function SisFiltraPacientesAfiliados(lcWhereOrder As String, wxParametroJAMO As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFiltraPacientesAfiliados = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFiltraPacientesAfiliados"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 2000, lcWhereOrder):   .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFiltraPacientesAfiliados = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function SisFiltraPacientesAtendidos(lcWhereOrder As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFiltraPacientesAtendidos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFiltraPacientesAtendidos"
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 1000, lcWhereOrder):   .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFiltraPacientesAtendidos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function





Function Rc01SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc01SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc01SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

Function Rc03SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc03SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc03SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function Rc04SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc04SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set Rc04SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

Function Rc05SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc05SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set Rc05SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function Rc06SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset


    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc06SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set Rc06SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function


Function Rc09SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc09SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc09SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function


Function Rc12SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc12SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc12SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function


Function Rc13SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc13SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc13SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function


Function Rc14SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc14SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc14SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

Function Rc15SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc15SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc15SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function Rc16SeleccionarPorFiltro(lcFiltro As String, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "Rc16SeleccionarPorFiltro"
            Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, lcFiltro): .Parameters.Append oParameter
            Set oRecordset = .Execute
    End With
    Set Rc16SeleccionarPorFiltro = oRecordset
    Set oCommand = Nothing
    Set oParameter = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description
End Function

Function TiposDestinoAtencionSeleccionarSoloSIS(lnIdTipoServicio As Long, oRsTipoDestinoSIS As Recordset) As ADODB.Recordset
    Dim lcSql As String, oRsTmp1 As New Recordset, orstmp As New Recordset
    With orstmp
          .Fields.Append "IdDestinoAtencion", adInteger
          .Fields.Append "DescripcionLarga", adVarChar, 100
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    If oRsTipoDestinoSIS.RecordCount > 0 Then
        Set oRsTmp1 = mo_ReglasAdmision.TiposDestinoAtencionSoloSIS(lnIdTipoServicio)
        If oRsTmp1.RecordCount > 0 Then
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              oRsTipoDestinoSIS.MoveFirst
              oRsTipoDestinoSIS.Find "rc04_idDestinoAsegurado='" & oRsTmp1.Fields!id_DestinoAseguradoSIS & "'"
              If Not oRsTipoDestinoSIS.EOF Then
                 orstmp.AddNew
                 orstmp.Fields!IdDestinoAtencion = oRsTmp1.Fields!IdDestinoAtencion
                 orstmp.Fields!DescripcionLarga = Left(oRsTmp1.Fields!codigo & " = " & oRsTmp1.Fields!Descripcion, 100)
                 orstmp.Update
              End If
              oRsTmp1.MoveNext
           Loop
        End If
        oRsTmp1.Close
    End If
    Set TiposDestinoAtencionSeleccionarSoloSIS = orstmp
End Function


Function ReglasDeConsistenciaSISestanOK(lnIdOpcionGalenHos As sghOpcionGalenHos, _
                                        wxParametro302 As String, lnIdFuenteFinanciamiento As Long, ByRef lcDevuelveMensaje As String, _
                                        lnIdCuentaAtencion As Long, lcCodigoPrestacion As String, wxParametro326 As String, _
                                        ml_TipoServicio As sghTipoServicio, _
                                        mi_opcion As sghOpciones, lbDesdeValidacionDatosObligatorios As Boolean, _
                                        lbDespuesDeElegirCodigoPrestacion As Boolean, lnNroHistoriaClinica As Long, _
                                        lcFechaIngreso As String, lcHoraIngreso As String, _
                                        Optional oRsFarmacia As Recordset, _
                                        Optional ByRef oRsDestinosHospEmerg As Recordset) As Boolean
         ReglasDeConsistenciaSISestanOK = True

         If wxParametro302 = "S" Then
            Dim oRsTmp1 As New Recordset
            Dim oRsTmp2 As New Recordset
            Dim orsTmp3 As New Recordset
            Dim rsReporte As New Recordset
            Dim oConexionExterna As New Connection
            Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
            Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
            Dim lcMensaje As String, lcCodigosCpts As String, lcCpt As String, lcCodigosFarmacia As String, lcFarmacia As String
            Dim lbAmbulatoria As Boolean, lbReferencia As Boolean, lbEmergencia As Boolean, lnPesoKg As Double
            Dim lnFor As Integer, lnPosActual As Integer, lbContinuar As Boolean, lcFiltro As String
            Dim lcCeroItemsMedicamentos As Boolean, lcCeroItemsCpt As Boolean, lbEsNuevo As Boolean
            oConexionExterna.CommandTimeout = 300
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
            oConexionExterna.CursorLocation = adUseClient
            Select Case lnIdOpcionGalenHos
            Case sghVentasFarmacia
                 If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And lnIdCuentaAtencion > 0 Then
                       Set oRsTmp1 = mo_AdminAdmision.AtencionesDatosAdicionalesSeleccionarPorIdCuenta(lnIdCuentaAtencion)
                       If oRsTmp1.RecordCount > 0 Then
                          lcCodigoPrestacion = IIf(IsNull(oRsTmp1.Fields!FuaCodigoPrestacion), vbNullString, oRsTmp1.Fields!FuaCodigoPrestacion)
                       End If
                       oRsTmp1.Close
                       If lcCodigoPrestacion <> vbNullString Then
                           lcFiltro = " and rc12_idServicio='" & lcCodigoPrestacion & "'"
                           Set oRsTmp1 = Rc12SeleccionarPorFiltro(lcFiltro, oConexionExterna)
                           If oRsTmp1.RecordCount > 0 Then
                              lcCeroItemsMedicamentos = False
                              lcCeroItemsCpt = False
                              'Mínimo 1 item en FArmacia - rc12
                              If oRsTmp1.Fields!rc12_MedIns = 1 Then
                                  If oRsFarmacia.RecordCount = 0 Then
                                     lcCeroItemsMedicamentos = True
                                  End If
                              End If
                              Select Case oRsTmp1.Fields!rc12_ControlCantidad
                              Case 1    'No grabar si falta Item
                                   If lcCeroItemsMedicamentos = True Then
                                      lcDevuelveMensaje = lcDevuelveMensaje & "Para la PRESTACION elejida, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                                   End If
                              Case 2    'No grabar si falta registrar Farmacia y Cie10
                                    If lcCeroItemsMedicamentos = True Then
                                       lcDevuelveMensaje = lcDevuelveMensaje & "Para la PRESTACION elejida, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                                    End If
                              Case 3    'Grabar si hay al menos un item de Farmacia o un item de Cie10
                    '                If lcCeroItemsMedicamentos = True Then
                    '                   lcDevuelveMensaje = lcDevuelveMensaje & "Para la PRESTACION elejida, debe registrar al menos 1 MEDICAMENTO/INSUMO (rc12)" & Chr(13)
                    '                End If
                              End Select
                           End If
                           oRsTmp1.Close
                           If oRsFarmacia.RecordCount > 0 Then
                              lcDevuelveMensaje = lcDevuelveMensaje & ReglasDeConsistenciaSISsoloFarmacia(oRsFarmacia, _
                                                                      lnIdCuentaAtencion, lcCodigoPrestacion, _
                                                                      oConexionExterna, lcFechaIngreso, ml_TipoServicio, _
                                                                      lnIdOpcionGalenHos, mi_opcion)
                           End If
                        Else
                           'lcDevuelveMensaje = lcDevuelveMensaje & "No se ha elegido CODIGO DE PRESTACION - FUA (SIS)" & Chr(13)
                           If oRsFarmacia.RecordCount > 0 Then
                              lcDevuelveMensaje = lcDevuelveMensaje & ReglasDeConsistenciaSISsoloFarmacia(oRsFarmacia, _
                                                                      lnIdCuentaAtencion, lcCodigoPrestacion, _
                                                                      oConexionExterna, lcFechaIngreso, ml_TipoServicio, _
                                                                      lnIdOpcionGalenHos, mi_opcion)
                           End If
                        End If
                     End If
            Case sghAdmisionEmergencia, sghAdmisionHospitalizacion
                If lbDesdeValidacionDatosObligatorios = True Then
                    If wxParametro302 = "S" And lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
                        If ml_TipoServicio = sghHospitalizacion Then
                           If mi_opcion = sghModificar And lcCodigoPrestacion = vbNullString Then
                              lcDevuelveMensaje = lcDevuelveMensaje + "Por favor elija 'Código de Prestación (SIS)'       (Ficha 2.1)" + Chr(13)
                              ReglasDeConsistenciaSISestanOK = False
                           End If
                        Else
                           If mi_opcion = sghAgregar And lcCodigoPrestacion = vbNullString Then
                              lcDevuelveMensaje = lcDevuelveMensaje + "Por favor elija 'Código de Prestación (SIS)'       (Ficha 2.1)" + Chr(13)
                              ReglasDeConsistenciaSISestanOK = False
                           End If
                        End If
                    End If
                ElseIf lbDespuesDeElegirCodigoPrestacion = True Then
                    If lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS And lcCodigoPrestacion <> vbNullString Then
                        lbContinuar = True
                        'Maximos FUA por dia,mes,año -  (será chequeado al Agregar/modificar) - rc13
                        lcFiltro = " and rc13_idServicio='" & lcCodigoPrestacion & "'"
                        Set oRsTmp1 = Rc13SeleccionarPorFiltro(lcFiltro, oConexionExterna)
                        If oRsTmp1.RecordCount > 0 Then
                           If Not (IsNull(oRsTmp1.Fields!rc13_TopeDia) And IsNull(oRsTmp1.Fields!rc13_topeMes) And IsNull(oRsTmp1.Fields!rc13_topeAnio)) Then
                                Dim lnFuasDelDia As Integer, lnFuasDelMes As Integer, lnFuasDelAnio As Integer
                                Sis_FuasRegistradasXpacienteSegunCodigoPrestacion CDate(lcFechaIngreso & " " & lcHoraIngreso), _
                                                     99999999, lcCodigoPrestacion, Trim(Str(lnNroHistoriaClinica)), _
                                                     lnFuasDelDia, lnFuasDelMes, lnFuasDelAnio
                                If Not (oRsTmp1.Fields!rc13_TopeDia >= lnFuasDelDia) And (oRsTmp1.Fields!rc13_topeMes >= lnFuasDelMes) And _
                                       (oRsTmp1.Fields!rc13_topeAnio >= lnFuasDelAnio) Then
                                    MsgBox "Para la PRESTACION elejida, excede en N°Fuas (rc13), (Fuas registradas Día: " & lnFuasDelDia & _
                                          ", Mes:" & lnFuasDelMes & ", Año: " & lnFuasDelAnio & ") (Topes: " & oRsTmp1.Fields!rc13_TopeDia & _
                                          "," & oRsTmp1.Fields!rc13_topeMes & "," & oRsTmp1.Fields!rc13_topeAnio & ")", vbInformation, "Reglas de Consistencias - SIS"
                                    ReglasDeConsistenciaSISestanOK = False
                                End If
                           End If
                        End If
                        oRsTmp1.Close
                        'Destino del Asegurado -  (será chequeado cuando se dé el ALTA MEDICA) - rc4
                        If ReglasDeConsistenciaSISestanOK = True And mi_opcion = sghModificar And lcCodigoPrestacion <> vbNullString Then
                           lcFiltro = " and rc04_idServicio='" & lcCodigoPrestacion & "'"
                           Set oRsTmp1 = Rc04SeleccionarPorFiltro(lcFiltro, oConexionExterna)
                           Set oRsDestinosHospEmerg = TiposDestinoAtencionSeleccionarSoloSIS(ml_TipoServicio, oRsTmp1)
                           oRsTmp1.Close
                        End If
                        
                    End If
                End If
            Case sghRegistroCitaCE
            End Select
            oConexionExterna.Close
            Set oConexionExterna = Nothing
            Set oRsTmp1 = Nothing
            Set mo_AdminAdmision = Nothing
            Set oRsTmp2 = Nothing
            Set mo_ReglasFarmacia = Nothing
            Set rsReporte = Nothing
         End If
End Function

'debb
Function ReglasDeConsistenciaSISsoloFarmacia(oRsFarmacia As Recordset, lnIdCuentaAtencion As Long, _
                                             lcCodigoPrestacion As String, oConexionExterna As Connection, _
                                             lcFechaIngreso As String, ml_TipoServicio As Long, _
                                             mo_lnIdTablaLISTBARITEMS As sghOpcionGalenHos, _
                                             mi_opcion As sghOpciones) As String
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim orsTmp3 As New Recordset
       Dim rsReporte As New Recordset
       Dim oParameter As ADODB.Parameter
       Dim oConexion As New Connection
       Dim oConexionSIS As New Connection
       Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
       Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
       Dim lbEsNuevo As Boolean, lcFiltro As String, lcDevuelveMensaje As String, lcPrestacion As String, lnCantidad As Long
       lcDevuelveMensaje = ""
       oConexionSIS.CommandTimeout = 300
       oConexionSIS.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
       oConexionSIS.CursorLocation = adUseClient
       lcPrestacion = vbNullString '""
       '

       Dim oCommand As New ADODB.Command
       With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexionSIS
            .CommandTimeout = 150
            .CommandText = "m_serviciosXcodSer"
            Set oParameter = .CreateParameter("@codSer", adVarChar, adParamInput, 3, lcFiltro): .Parameters.Append oParameter
            Set rsReporte = .Execute
            Set rsReporte.ActiveConnection = Nothing
       End With
       If rsReporte.RecordCount > 0 Then
          lcPrestacion = " (Cod.Prestación SIS: " & lcCodigoPrestacion & "-" & Trim(rsReporte.Fields!ser_descripcion) & ")"
       End If
       rsReporte.Close
       Set oCommand = Nothing
       '
       oConexion.CommandTimeout = 300
       oConexion.Open sighentidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       '******************Consumos - INICIO ***************************************
       With oRsTmp2
               .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
               .Fields.Append "Cantidad", adInteger
               .Fields.Append "CantidadDia", adInteger
               .Fields.Append "CantidadTotal", adInteger
               .CursorType = adOpenDynamic
               .LockType = adLockOptimistic
               .Open
       End With
       With orsTmp3
               .Fields.Append "formaF", adVarChar, 20, adFldIsNullable
               .Fields.Append "Cantidad", adInteger
               .Fields.Append "CantidadTotal", adInteger
               .CursorType = adOpenDynamic
               .LockType = adLockOptimistic
               .Open
       End With
       'Consumo de ese momento
       oRsFarmacia.MoveFirst
       Do While Not oRsFarmacia.EOF
            If mo_lnIdTablaLISTBARITEMS = sghRegistroAtencionCE Then
               lnCantidad = oRsFarmacia!Recetado
            Else
               lnCantidad = oRsFarmacia!Cantidad
            End If
            'ff
            'debb-27/08/2015 (inicio)
            If UCase(Left(oRsFarmacia!tipo, 1)) = "M" Then   'solo MEDICAMENTOS
                lbEsNuevo = True
                If orsTmp3.RecordCount > 0 Then
                   orsTmp3.MoveFirst
                   orsTmp3.Find "formaF='" & oRsFarmacia.Fields!formaF & "'"
                   If Not orsTmp3.EOF Then
                      lbEsNuevo = False
                   End If
                End If
                If lbEsNuevo = True Then
                   orsTmp3.AddNew
                   orsTmp3.Fields!formaF = oRsFarmacia!formaF
                   orsTmp3.Fields!Cantidad = lnCantidad
                   If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghAgregar Then
                      orsTmp3.Fields!cantidadTotal = lnCantidad
                   ElseIf mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghModificar Then
                      orsTmp3.Fields!cantidadTotal = lnCantidad - oRsFarmacia!CantidadSinEditar
                   End If
                Else
                   orsTmp3.Fields!Cantidad = orsTmp3.Fields!Cantidad + lnCantidad
                   If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghAgregar Then
                      orsTmp3.Fields!cantidadTotal = orsTmp3.Fields!cantidadTotal + lnCantidad
                   ElseIf mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghModificar Then
                      orsTmp3.Fields!cantidadTotal = orsTmp3.Fields!cantidadTotal - (lnCantidad - oRsFarmacia!CantidadSinEditar)
                   End If
                End If
                orsTmp3.Update
            End If
            'debb-27/08/2015 (fin)
            'codigo
            lbEsNuevo = True
            If orsTmp3.RecordCount > 0 Then
               orsTmp3.MoveFirst
               orsTmp3.Find "formaF='" & oRsFarmacia.Fields!formaF & "'"
               If Not orsTmp3.EOF Then
                  lbEsNuevo = False
               End If
            End If
            If lbEsNuevo = True Then
               orsTmp3.AddNew
               orsTmp3.Fields!formaF = oRsFarmacia!formaF
               orsTmp3.Fields!Cantidad = lnCantidad
               If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghAgregar Then
                  orsTmp3.Fields!cantidadTotal = lnCantidad
               ElseIf mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghModificar Then
                  orsTmp3.Fields!cantidadTotal = lnCantidad - oRsFarmacia!CantidadSinEditar
               End If
            Else
               orsTmp3.Fields!Cantidad = orsTmp3.Fields!Cantidad + lnCantidad
               If mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghAgregar Then
                  orsTmp3.Fields!cantidadTotal = orsTmp3.Fields!cantidadTotal + lnCantidad
               ElseIf mo_lnIdTablaLISTBARITEMS = sghOpcionGalenHos.sghVentasFarmacia And mi_opcion = sghModificar Then
                  orsTmp3.Fields!cantidadTotal = orsTmp3.Fields!cantidadTotal - (lnCantidad - oRsFarmacia!CantidadSinEditar)
               End If
            End If
            orsTmp3.Update
            'codigo
            lbEsNuevo = True
            If oRsTmp2.RecordCount > 0 Then
               oRsTmp2.MoveFirst
               oRsTmp2.Find "codigo='" & oRsFarmacia.Fields!codigo & "'"
               If Not oRsTmp2.EOF Then
                  lbEsNuevo = False
               End If
            End If
            If lbEsNuevo = True Then
               oRsTmp2.AddNew
               oRsTmp2.Fields!codigo = oRsFarmacia!codigo
               oRsTmp2.Fields!Cantidad = lnCantidad
               oRsTmp2.Fields!cantidadDia = lnCantidad
               oRsTmp2.Fields!cantidadTotal = lnCantidad
            Else
               oRsTmp2.Fields!Cantidad = oRsTmp2.Fields!Cantidad + lnCantidad
               oRsTmp2.Fields!cantidadDia = oRsTmp2.Fields!cantidadDia + lnCantidad
               oRsTmp2.Fields!cantidadTotal = oRsTmp2.Fields!cantidadTotal + lnCantidad
            End If
            oRsTmp2.Update
            oRsFarmacia.MoveNext
       Loop
       'Consumo total y del día
       Set rsReporte = SISDespachosFarmaciaIncluyeDevoluciones(lnIdCuentaAtencion)
       If rsReporte.RecordCount > 0 Then
          rsReporte.MoveFirst
          Do While Not rsReporte.EOF
             lbEsNuevo = True
             If oRsTmp2.RecordCount > 0 Then
               oRsTmp2.MoveFirst
               oRsTmp2.Find "codigo='" & rsReporte.Fields!codigo & "'"
               If Not oRsTmp2.EOF Then
                  lbEsNuevo = False
               End If
             End If
             If lbEsNuevo = True Then
                oRsTmp2.AddNew
                oRsTmp2.Fields!codigo = rsReporte!codigo
                oRsTmp2.Fields!Cantidad = 0
                If Format(rsReporte!FechaCreacion, "dd/mm/yyyy") = lcFechaIngreso Then
                   oRsTmp2.Fields!cantidadDia = rsReporte!Cantidad
                End If
                oRsTmp2.Fields!cantidadTotal = rsReporte!Cantidad
             Else
                'oRsTmp2.Fields!cantidad = oRsTmp2.Fields!cantidad + rsReporte!cantidad
                If Format(rsReporte!FechaCreacion, "dd/mm/yyyy") = lcFechaIngreso Then
                   oRsTmp2.Fields!cantidadDia = oRsTmp2.Fields!cantidadDia + rsReporte!Cantidad
                End If
                oRsTmp2.Fields!cantidadTotal = oRsTmp2.Fields!cantidadTotal + rsReporte!Cantidad
             End If
             oRsTmp2.Update
             If orsTmp3.RecordCount > 0 Then
                orsTmp3.MoveFirst
                orsTmp3.Find "formaF='" & rsReporte!FormaFarmaceutica & "'"
                If Not orsTmp3.EOF Then
                   orsTmp3.Fields!cantidadTotal = orsTmp3.Fields!cantidadTotal + rsReporte!Cantidad
                End If

             End If
             rsReporte.MoveNext
          Loop
       End If
       rsReporte.Close
       '******************Consumos - FINAL ***************************************       '
       'Topes por Unidades de Medida de FARMACIA - rc06
       If InStr("/065/067/068", lcCodigoPrestacion) = 0 Then 'La regla de consistencia 6, no aplica para codigo de prestacion 065,067,068
            If orsTmp3.RecordCount > 0 Then
                 orsTmp3.MoveFirst
                 Do While Not orsTmp3.EOF
                     lcFiltro = " and rc06_AbreviaturaFF='" & Trim(orsTmp3.Fields!formaF) & "'"
                     Set oRsTmp1 = Rc06SeleccionarPorFiltro(lcFiltro, oConexionExterna)
                     If oRsTmp1.RecordCount > 0 Then
                          If ml_TipoServicio <> sghHospitalizacion Then
                             If Not (orsTmp3.Fields!cantidadTotal >= oRsTmp1.Fields!rc06_topeMinimo And orsTmp3.Fields!cantidadTotal <= oRsTmp1.Fields!rc06_topeNoHosp) Then
                                lcDevuelveMensaje = lcDevuelveMensaje & "Para la FORMA FARMACEUTICA: " & orsTmp3!formaF & ", la 'cantidad total entregada' pasa de: " & _
                                            oRsTmp1.Fields!rc06_topeMinimo & " , " & oRsTmp1.Fields!rc06_topeNoHosp & " (RC06)" & Chr(13)
                             End If
                          Else
                             If Not (orsTmp3.Fields!cantidadTotal >= oRsTmp1.Fields!rc06_topeMinimo And orsTmp3.Fields!cantidadTotal <= oRsTmp1.Fields!rc06_topeHosp) Then
                                lcDevuelveMensaje = lcDevuelveMensaje & "Para la FORMA FARMACEUTICA: " & orsTmp3!formaF & ", la 'cantidad total entregada' pasa de: " & _
                                            oRsTmp1.Fields!rc06_topeMinimo & " , " & oRsTmp1.Fields!rc06_topeHosp & " (RC06)" & Chr(13)
                             End If
                          End If
                     End If
                     oRsTmp1.Close
                     orsTmp3.MoveNext
                  Loop
             End If
        End If
        'Topes de Oxígeno y Bomba de Infusión FARMACIA - rc09
        lcFiltro = " rc09_idServicio='" & lcCodigoPrestacion & "'"
        Set oRsTmp1 = Rc09SeleccionarPorFiltro(lcFiltro, oConexionExterna)
        If oRsTmp1.RecordCount > 0 Then
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
                  oRsTmp2.MoveFirst
                  If IsNull(oRsTmp1.Fields!rc09_codMed) Then
                     oRsTmp2.Find "codigo='" & oRsTmp1.Fields!rc09_codIns & "'"
                     If Not oRsTmp2.EOF Then
                        If Not (oRsTmp2.Fields!cantidadTotal >= oRsTmp1.Fields!rc09_topeMinimo And oRsTmp2.Fields!cantidadTotal <= oRsTmp1.Fields!rc09_topeMaximo) Then
                           lcDevuelveMensaje = lcDevuelveMensaje & "Para el Insumo: " & oRsTmp2!codigo & ", la 'cantidad total entregada' pasa de: " & oRsTmp1.Fields!rc09_topeMinimo & _
                                       " , " & oRsTmp1.Fields!rc09_topeMaximo & lcPrestacion & " (RC09)" & Chr(13)
                        End If
                     End If
                  Else
                     oRsTmp2.Find "codigo='" & oRsTmp1.Fields!rc09_codMed & "'"
                     If Not oRsTmp2.EOF Then
                        If Not (oRsTmp2.Fields!cantidadDia >= oRsTmp1.Fields!rc09_topeMinimo And oRsTmp2.Fields!cantidadDia <= oRsTmp1.Fields!rc09_topeMaximo) Then
                           lcDevuelveMensaje = lcDevuelveMensaje & "Para el Medicamento: " & oRsTmp2!codigo & ", la 'cantidad total entregada' pasa de: " & oRsTmp1.Fields!rc09_topeMinimo & _
                                       " , " & oRsTmp1.Fields!rc09_topeMaximo & lcPrestacion & " (RC09)" & Chr(13)
                        End If
                     End If
                  End If
                  oRsTmp1.MoveNext
            Loop
        End If
        oRsTmp1.Close
        oConexion.Close
        oConexionSIS.Close
        Set mo_AdminAdmision = Nothing
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set oConexion = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set rsReporte = Nothing
        Set oConexionSIS = Nothing
        ReglasDeConsistenciaSISsoloFarmacia = lcDevuelveMensaje
End Function


Function SisFuaAtencionSeleccionarPorCuenta(lnIdCuentaAtencion As Long, Optional oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFuaAtencionSeleccionarPorCuenta = Nothing
    ms_MensajeError = ""
    If oConexionExterna Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionExterna Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionExterna
        End If
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexionExterna Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set SisFuaAtencionSeleccionarPorCuenta = oRecordset
   
   If oConexionExterna Is Nothing Then
      oConexion.Close
      Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function SisFuaSeleccionarTodos() As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFuaSeleccionarTodos = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = " SisFuaSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaSeleccionarTodos = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

Sub FuaActualizaDespachosEnFarmacia(lnIdCuentaAtencion As Long, wxParametro302 As String, lnIdTipoServicio As sghTipoServicio, lnIdFuenteFinanciamiento As Long)
    If lnIdCuentaAtencion > 0 And wxParametro302 = "S" And lnIdTipoServicio = sghConsultaExterna And lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
       On Error GoTo ErrFFua
       Dim oConexionExterna As New Connection
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim orsTmp3 As New Recordset
       Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
       Dim lcSql As String
       Dim oSisFuaAtencionINS As New SisFuaAtencionINS, oDoSisFuaAtencionINS As New DoSisFuaAtencionINS
       Dim oSisFuaAtencionMED As New SisFuaAtencionMED, oDoSisFuaAtencionMED As New DoSisFuaAtencionMED
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       oConexionExterna.CursorLocation = adUseClient
       SisFuaAtencionINS_MED_limpiaCantidadEntregada lnIdCuentaAtencion, oConexionExterna
       Set oRsTmp1 = SisFuaAtencionINSxIdCuentaAtencion(lnIdCuentaAtencion)
       Set orsTmp3 = SisFuaAtencionMEDxIdCuentaAtencion(lnIdCuentaAtencion)
       Set oRsTmp2 = SISDespachosFarmaciaIncluyeDevoluciones(lnIdCuentaAtencion)
       
       oConexionExterna.BeginTrans
       Set oSisFuaAtencionINS.Conexion = oConexionExterna
       Set oSisFuaAtencionMED.Conexion = oConexionExterna
       '
       '
       'despachos y devoluciones
       If oRsTmp2.RecordCount > 0 Then
           oRsTmp2.MoveFirst
           Do While Not oRsTmp2.EOF
              If oRsTmp2.Fields!TipoProducto = 1 Then
                    If oRsTmp1.RecordCount > 0 Then
                          oRsTmp1.MoveFirst
                          oRsTmp1.Find "codigo='" & oRsTmp2.Fields!codigo & "'"
                          If Not oRsTmp1.EOF Then
                             oDoSisFuaAtencionINS.id = oRsTmp1.Fields!id
                             If Not oSisFuaAtencionINS.SeleccionarPorId(oDoSisFuaAtencionINS) Then
                                MsgBox oSisFuaAtencionINS.MensajeError: GoTo ErrFFua
                             End If
                             oDoSisFuaAtencionINS.CantidadEntregada = oDoSisFuaAtencionINS.CantidadEntregada + oRsTmp2.Fields!Cantidad
                             If Not oSisFuaAtencionINS.Modificar(oDoSisFuaAtencionINS) Then
                                MsgBox oSisFuaAtencionINS.MensajeError: GoTo ErrFFua
                             End If
                          End If
                    End If
              Else
                    If orsTmp3.RecordCount > 0 Then
                          orsTmp3.MoveFirst
                          orsTmp3.Find "codigo='" & oRsTmp2.Fields!codigo & "'"
                          If Not orsTmp3.EOF Then
                             oDoSisFuaAtencionMED.id = orsTmp3.Fields!id
                             If Not oSisFuaAtencionMED.SeleccionarPorId(oDoSisFuaAtencionMED) Then
                                MsgBox oSisFuaAtencionMED.MensajeError: GoTo ErrFFua
                             End If
                             oDoSisFuaAtencionMED.CantidadEntregada = oDoSisFuaAtencionMED.CantidadEntregada + oRsTmp2.Fields!Cantidad
                             If Not oSisFuaAtencionMED.Modificar(oDoSisFuaAtencionMED) Then
                                MsgBox oSisFuaAtencionMED.MensajeError: GoTo ErrFFua
                             End If
                          End If
                    End If
              End If
              oRsTmp2.MoveNext
           Loop
       End If
       oConexionExterna.CommitTrans
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set orsTmp3 = Nothing
       Set oSisFuaAtencionINS = Nothing
       Set oDoSisFuaAtencionINS = Nothing
       Set oSisFuaAtencionMED = Nothing
       Set oDoSisFuaAtencionMED = Nothing
    End If
    Exit Sub
ErrFFua:
Resume
    oConexionExterna.RollbackTrans
    Set oConexionExterna = Nothing
    Set oRsTmp1 = Nothing
End Sub

Sub FuaActualizaDespachosEnServicios(lnIdCuentaAtencion As Long, wxParametro302 As String, _
                                     lnIdTipoServicio As Long, lnIdFuenteFinanciamiento As Long)
    If lnIdCuentaAtencion > 0 And wxParametro302 = "S" And lnIdTipoServicio = sghConsultaExterna And lnIdFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFSIS Then
       On Error GoTo ErrFFuaS
       Dim oConexionExterna As New Connection
       Dim oRsTmp1 As New Recordset
       Dim oRsTmp2 As New Recordset
       Dim orsTmp3 As New Recordset
       Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
       Dim lcSql As String
       Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO, oSisFuaAtencionPRO As New SisFuaAtencionPRO
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       oConexionExterna.CursorLocation = adUseClient
       SisFuaAtencionPRO_limpiaCantidadEjecutada lnIdCuentaAtencion, oConexionExterna
       Set oRsTmp1 = SisFuaAtencionPROxIdCuentaAtencion(lnIdCuentaAtencion)
       oConexionExterna.BeginTrans
       Set oSisFuaAtencionPRO.Conexion = oConexionExterna
       'despachos
       Set oRsTmp2 = SISDespachosServiciosIncluyeDevoluciones(lnIdCuentaAtencion)
       If oRsTmp2.RecordCount > 0 Then
           oRsTmp2.MoveFirst
           Do While Not oRsTmp2.EOF
              If oRsTmp1.RecordCount > 0 Then
                  oRsTmp1.MoveFirst
                  oRsTmp1.Find "codigo='" & oRsTmp2.Fields!codigo & "'"
                  If Not oRsTmp1.EOF Then
                        oDoSisFuaAtencionPRO.id = oRsTmp1.Fields!id
                        If Not oSisFuaAtencionPRO.SeleccionarPorId(oDoSisFuaAtencionPRO) Then
                           MsgBox oSisFuaAtencionPRO.MensajeError: GoTo ErrFFuaS
                        End If
                        oDoSisFuaAtencionPRO.CantidadEjecutada = oDoSisFuaAtencionPRO.CantidadEjecutada + oRsTmp2.Fields!Cantidad
                        If Not oSisFuaAtencionPRO.Modificar(oDoSisFuaAtencionPRO) Then
                           MsgBox oSisFuaAtencionPRO.MensajeError: GoTo ErrFFuaS
                        End If
                  End If
              End If
              oRsTmp2.MoveNext
           Loop
       End If
       oConexionExterna.CommitTrans
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set orsTmp3 = Nothing
       Set oDoSisFuaAtencionPRO = Nothing
       Set oSisFuaAtencionPRO = Nothing
    End If
    Exit Sub
ErrFFuaS:
    oConexionExterna.RollbackTrans
    Set oConexionExterna = Nothing
    Set oRsTmp1 = Nothing
End Sub

Function SISDespachosFarmaciaIncluyeDevoluciones(lnIdCuentaAtencion As Long) As Recordset
    Dim oRsTmp2 As New Recordset, oRsTmpRetorno As New Recordset
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    With oRsTmpRetorno
          .Fields.Append "codigo", adVarChar, 7
          .Fields.Append "cantidad", adInteger
          .Fields.Append "TipoProducto", adInteger
          .Fields.Append "fechaCreacion", adDate
          .Fields.Append "precio", adDouble
          .Fields.Append "nombre", adVarChar, 300
          .Fields.Append "FormaFarmaceutica", adVarChar, 10
          .Fields.Append "DocumentoNumero", adVarChar, 20
          .Fields.Append "idProducto", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    'despachos
    Set oRsTmp2 = mo_ReglasFarmacia.farmMovimientoVentasDetalleSeleccionarXidCuenta(lnIdCuentaAtencion)
    If oRsTmp2.RecordCount > 0 Then
        oRsTmp2.MoveFirst
        Do While Not oRsTmp2.EOF
           If oRsTmp2.Fields!Precio > 0 Then
                oRsTmpRetorno.AddNew
                oRsTmpRetorno.Fields!codigo = oRsTmp2.Fields!codigo
                oRsTmpRetorno.Fields!Cantidad = oRsTmp2.Fields!Cantidad
                oRsTmpRetorno.Fields!TipoProducto = oRsTmp2.Fields!TipoProducto
                oRsTmpRetorno.Fields!FechaCreacion = oRsTmp2.Fields!FechaCreacion
                oRsTmpRetorno.Fields!Precio = oRsTmp2.Fields!Precio
                oRsTmpRetorno.Fields!nombre = oRsTmp2.Fields!nombre
                oRsTmpRetorno.Fields!FormaFarmaceutica = oRsTmp2.Fields!FormaFarmaceutica
                oRsTmpRetorno.Fields!IdProducto = oRsTmp2.Fields!IdProducto
                oRsTmpRetorno.Fields!DocumentoNumero = oRsTmp2.Fields!DocumentoNumero
           End If
           oRsTmp2.MoveNext
        Loop
    End If
    'devoluciones
    If oRsTmpRetorno.RecordCount > 0 Then
        Set oRsTmp2 = mo_ReglasFarmacia.FarmMovimientoNotaIngresoSeleccionarPorCuenta(lnIdCuentaAtencion)
        oRsTmp2.Filter = "idEstadoMovimiento=1"
        If oRsTmp2.RecordCount > 0 Then
            oRsTmp2.MoveFirst
            Do While Not oRsTmp2.EOF
               oRsTmpRetorno.MoveFirst
               oRsTmpRetorno.Find "codigo='" & oRsTmp2.Fields!codigo & "'"
               If Not oRsTmpRetorno.EOF Then
                    oRsTmpRetorno.Fields!Cantidad = oRsTmpRetorno.Fields!Cantidad - oRsTmp2.Fields!Cantidad
                    oRsTmpRetorno.Update
               End If
               oRsTmp2.MoveNext
            Loop
        End If
    End If
    Set oRsTmp2 = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set SISDespachosFarmaciaIncluyeDevoluciones = oRsTmpRetorno
End Function

Function SISDespachosServiciosIncluyeDevoluciones(lnIdCuentaAtencion As Long) As Recordset
    Dim oRsTmp2 As New Recordset, oRsTmpRetorno As New Recordset
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    With oRsTmpRetorno
          .Fields.Append "codigo", adVarChar, 5
          .Fields.Append "cantidad", adInteger
          .Fields.Append "idTipoProducto", adInteger
          .Fields.Append "precio", adDouble
          .Fields.Append "nombre", adVarChar, 300
          .Fields.Append "idProducto", adInteger
          .Fields.Append "idPuntoCarga", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    'despachos
    Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorCuenta(lnIdCuentaAtencion)
    oRsTmp2.Filter = "idEstadoFacturacion<>9"
    If oRsTmp2.RecordCount > 0 Then
        oRsTmp2.MoveFirst
        Do While Not oRsTmp2.EOF
           If Len(oRsTmp2.Fields!codigo) <= 5 Then
                oRsTmpRetorno.AddNew
                oRsTmpRetorno.Fields!codigo = oRsTmp2.Fields!codigo
                oRsTmpRetorno.Fields!Cantidad = oRsTmp2.Fields!Cantidad
                oRsTmpRetorno.Fields!Precio = oRsTmp2.Fields!Precio
                oRsTmpRetorno.Fields!nombre = oRsTmp2.Fields!nombre
                oRsTmpRetorno.Fields!IdProducto = oRsTmp2.Fields!IdProducto
                oRsTmpRetorno.Fields!idPuntoCarga = oRsTmp2.Fields!idPuntoCarga
           End If
           oRsTmp2.MoveNext
        Loop
    End If
    Set oRsTmp2 = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set SISDespachosServiciosIncluyeDevoluciones = oRsTmpRetorno
End Function
'debb
Function SisA_LotesSeleccionarXCodigo(lcCodigo As String) As Recordset
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  Dim sSQL1 As String
  Set SisA_LotesSeleccionarXCodigo = Nothing
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "a_lotesXidTablaSiaSis"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 2, lcCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
  End With

  Set SisA_LotesSeleccionarXCodigo = oRecordset
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Sub SisFiliacionesDevuelveKEY(ByRef lnIdSiaSis As Long, ByRef lcSIScodigo As String, _
                              lcApellPaterno As String, lcApellMaterno As String, _
                              lcPrimerNombre As String, ldFechaNacimiento As Date, _
                              ByRef lcCodigoEstablecimientoAdscripcionSIS As String)
    On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRsTmp1 As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFiliacionesDevuelveKEY"
        Set oParameter = .CreateParameter("@Paterno", adVarChar, adParamInput, 40, lcApellPaterno):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Materno", adVarChar, adParamInput, 40, lcApellMaterno):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Pnombre", adVarChar, adParamInput, 70, lcPrimerNombre):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Fnacimiento", adDBTimeStamp, adParamInput, 0, ldFechaNacimiento): .Parameters.Append oParameter
        Set oRsTmp1 = .Execute
        Set oRsTmp1.ActiveConnection = Nothing
   End With
   lnIdSiaSis = 0
   lcSIScodigo = ""
   lcCodigoEstablecimientoAdscripcionSIS = ""
   If oRsTmp1.RecordCount > 0 Then
       lnIdSiaSis = oRsTmp1.Fields!idSiasis
       lcSIScodigo = oRsTmp1.Fields!codigo
       lcCodigoEstablecimientoAdscripcionSIS = IIf(IsNull(oRsTmp1.Fields!CodigoEstablAdscripcion), "", oRsTmp1.Fields!CodigoEstablAdscripcion)
   End If
   oConexion.Close
   Set oRsTmp1 = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Sub
ManejadorDeError:
   MsgBox Err.Description

End Sub



Function SisFUAyaFueEnviadoAlSisLIMA(lnIdCuentaAtencion As Long, lnIdTipoFinanciamientoActual As Long, _
                                     wxParametro302 As String) As Boolean
    SisFUAyaFueEnviadoAlSisLIMA = False
    If wxParametro302 = "S" And lnIdTipoFinanciamientoActual = sghTipoFinanciamiento.sghSis Then
        Dim lcSql As String, oRsTmp1 As New Recordset
        Set oRsTmp1 = sisFuaAtencionXCuenta(lnIdCuentaAtencion)
        If oRsTmp1.RecordCount > 0 Then
           If Not IsNull(oRsTmp1.Fields!CabNroEnvioAlSIS) Then
              If Val(oRsTmp1.Fields!CabNroEnvioAlSIS) > 0 Then
                 SisFUAyaFueEnviadoAlSisLIMA = True
                 MsgBox "El Formato FUA de la Cuenta ya fué enviada al SIS CENTRAL - LIMA" & Chr(13) & _
                        "FUA N°: " & oRsTmp1!FuaDisa & "-" & oRsTmp1!FuaLote & "-" & oRsTmp1!FuaNumero, vbInformation, "Facturación"
              End If
           End If
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
    End If
End Function

Sub SisFuaAtencionActualizaDatosDesdeHospEmegCE(ml_IdCuentaAtencion As Long, ml_TipoServicio As Long, _
                                                ml_idAtencion As Long, mo_lnIdTablaLISTBARITEMS As sghOpcionGalenHos, _
                                                ml_idUsuario As Long)
    Dim lcSql As String
    Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
    Dim oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
    Dim oRsPatologia As New Recordset
    Dim oRsFarmacia As New Recordset
    Dim oRsDx As New Recordset, oRsVacunasSp As New Recordset
    Dim lcDxPrincipal As String, lcDxPrincipalNro As Long
    Dim oConexion As New Connection, oConexionExterna As New Connection
    Dim oDoSisFuaAtencion As New DoSisFuaAtencion, oSisFuaAtencion As New SisFuaAtencion
    Dim oDoSisFuaAtencionDIA As New DoSisFuaAtencionDIA: Dim oSisFuaAtencionDIA As New SisFuaAtencionDIA
    Dim oDoSisFuaAtencionINS As New DoSisFuaAtencionINS: Dim oSisFuaAtencionINS As New SisFuaAtencionINS
    Dim oDoSisFuaAtencionMED As New DoSisFuaAtencionMED: Dim oSisFuaAtencionMED As New SisFuaAtencionMED
    Dim oDoSisFuaAtencionPRO As New DoSisFuaAtencionPRO: Dim oSisFuaAtencionPRO As New SisFuaAtencionPRO
    Dim oDoSisFuaAtencionSMI As New DoSisFuaAtencionSMI: Dim oSisFuaAtencionSMI As New SisFuaAtencionSMI
    Dim oDOEstablecimiento As New DOEstablecimiento, oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
    Dim lcTxtSpPeso As String, lcTxtSPtalla As String, lcTxtSPpa As String, lcTxtObservaciones As String
    Dim oDoServicio As New doServicio
    
    Dim lnRegistros As Long, lbNuevo As Boolean, lbTodoOK As Boolean
    Dim lcCodigoSis As String, lcDescripcion As String
    Const lcInsumo As String = "Insumo": Const lcMedicamento As String = "Medicamento"
    Const lcOtros As String = "Otros": Const lcLaboratorio As String = "Laboratorio"
    Const lcImagenes As String = "Imágenes"
    On Error GoTo ErrActHECE
    lbTodoOK = False
    oConexionExterna.CommandTimeout = 300
    oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.BeginTrans
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oRsTmp1 = sisFuaAtencionXCuenta(ml_IdCuentaAtencion)
    If oRsTmp1.RecordCount > 0 And IsNull(oRsTmp1.Fields!CabNroEnvioAlSIS) Then
        FUAcreaTemporales oRsDx, oRsPatologia, oRsFarmacia, oRsVacunasSp
        FuaCargaDxDesdeGAlenHos oRsDx, oConexion, ml_idAtencion, ml_TipoServicio, lcDxPrincipal, lcDxPrincipalNro
        FuaCargaSIDesdeGAlenHos oRsFarmacia, oRsPatologia, mo_lnIdTablaLISTBARITEMS, ml_IdCuentaAtencion, lcInsumo, _
                                lcMedicamento, lcDxPrincipal, lcDxPrincipalNro, lcOtros, lcLaboratorio, lcImagenes
        FuaCargaTriajeVacunasDesdeGAlenHos lcTxtSpPeso, lcTxtSPtalla, lcTxtSPpa, lcTxtObservaciones, ml_idAtencion, _
                                           oConexionExterna
        'vacunas
        If Val(lcTxtSpPeso) <> 0 Then
            oRsVacunasSp.AddNew
            oRsVacunasSp.Fields!intervencionP = "003"
            oRsVacunasSp.Fields!Valor = lcTxtSpPeso
            oRsVacunasSp.Update
        End If
        If Val(lcTxtSPtalla) <> 0 Then
            oRsVacunasSp.AddNew
            oRsVacunasSp.Fields!intervencionP = "004"
            oRsVacunasSp.Fields!Valor = lcTxtSPtalla
            oRsVacunasSp.Update
        End If
        If lcTxtSPpa <> sighentidades.PresionDevuelveVacia Then
            oRsVacunasSp.AddNew
            oRsVacunasSp.Fields!intervencionP = "901"
            oRsVacunasSp.Fields!Valor = Left(lcTxtSPpa, InStr(lcTxtSPpa, "/") - 1)
            oRsVacunasSp.Update
            oRsVacunasSp.AddNew
            oRsVacunasSp.Fields!intervencionP = "301"
            oRsVacunasSp.Fields!Valor = Trim(Mid(lcTxtSPpa, InStr(lcTxtSPpa, "/") + 1, 10))
            oRsVacunasSp.Update
        End If
        'Graba FuaAtencion
        Set oSisFuaAtencion.Conexion = oConexionExterna
        oDoSisFuaAtencion.IdCuentaAtencion = ml_IdCuentaAtencion
        If oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) = False Then
           MsgBox oSisFuaAtencion.MensajeError: Exit Sub
        End If
        Set oRsTmp2 = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(ml_IdCuentaAtencion, oConexion)
        If Not IsNull(oRsTmp2.Fields!idEstablecimientoOrigen) Then
             Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp2.Fields!idEstablecimientoOrigen)
             Sis_m_ee_ssDevuelveCodigoSIS oDOEstablecimiento.codigo, lcCodigoSis, lcDescripcion
             'oDoSisFuaAtencion.FuaReferidoOrigenCodigoRENAES = lcCodigoSis
             oDoSisFuaAtencion.FuaReferidoOrigenCodigoRENAES = Right("0000000000" & oDOEstablecimiento.codigo, 10)
             oDoSisFuaAtencion.FuaReferidoOrigenNreferencia = IIf(IsNull(oRsTmp2.Fields!nroReferenciaOrigen), "", oRsTmp2.Fields!nroReferenciaOrigen)
        ElseIf Not IsNull(oRsTmp2.Fields!IdEstablecimientoNoMinsaOrigen) Then
             Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosNoMinsaSeleccionarPorId(oRsTmp2.Fields!IdEstablecimientoNoMinsaOrigen)
             Sis_m_ee_ssDevuelveCodigoSIS oDOEstablecimientoNoMinsa.codigo, lcCodigoSis, lcDescripcion
'             oDoSisFuaAtencion.FuaReferidoOrigenCodigoRENAES = lcCodigoSis
             oDoSisFuaAtencion.FuaReferidoOrigenCodigoRENAES = Right("0000000000" & oDOEstablecimientoNoMinsa.codigo, 10)
             oDoSisFuaAtencion.FuaReferidoOrigenNreferencia = IIf(IsNull(oRsTmp2.Fields!nroReferenciaOrigen), "", oRsTmp2.Fields!nroReferenciaOrigen)
        End If
        'oRsTmp1.Fields!FuaDestino = mo_AdminAdmision.TiposDestinoAtencionDevuelveIdSis(IIf(IsNull(oRsTmp2.Fields!IdDestinoAtencion), 0, oRsTmp2.Fields!IdDestinoAtencion))
        oDoSisFuaAtencion.FuaDestino = mo_AdminAdmision.TiposDestinoAtencionDevuelveIdSis(IIf(IsNull(oRsTmp2.Fields!IdDestinoAtencion), 1, oRsTmp2.Fields!IdDestinoAtencion))
        If Not IsNull(oRsTmp2.Fields!idEstablecimientoDestino) Then
             Set oDOEstablecimiento = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp2.Fields!idEstablecimientoDestino)
             Sis_m_ee_ssDevuelveCodigoSIS oDOEstablecimiento.codigo, lcCodigoSis, lcDescripcion
             'oRsTmp1.Fields!FuaReferidoDestinoCodigoRENAES = lcCodigoSis
             'oRsTmp1.Fields!FuaReferidoDestinoNreferencia = IIf(IsNull(oRsTmp2.Fields!NroReferenciaDestino), "", oRsTmp2.Fields!NroReferenciaDestino)
             oDoSisFuaAtencion.FuaReferidoDestinoCodigoRENAES = lcCodigoSis
             oDoSisFuaAtencion.FuaReferidoDestinoNreferencia = IIf(IsNull(oRsTmp2.Fields!NroReferenciaDestino), "", oRsTmp2.Fields!NroReferenciaDestino)
        ElseIf Not IsNull(oRsTmp2.Fields!IdEstablecimientoNoMinsaDestino) Then
             Set oDOEstablecimientoNoMinsa = mo_AdminServiciosComunes.EstablecimientosSeleccionarPorId(oRsTmp2.Fields!IdEstablecimientoNoMinsaDestino)
             Sis_m_ee_ssDevuelveCodigoSIS oDOEstablecimientoNoMinsa.codigo, lcCodigoSis, lcDescripcion
             'oRsTmp1.Fields!FuaReferidoDestinoCodigoRENAES = lcCodigoSis
             'oRsTmp1.Fields!FuaReferidoDestinoNreferencia = IIf(IsNull(oRsTmp2.Fields!NroReferenciaDestino), "", oRsTmp2.Fields!NroReferenciaDestino)
             oDoSisFuaAtencion.FuaReferidoDestinoCodigoRENAES = lcCodigoSis
             oDoSisFuaAtencion.FuaReferidoDestinoNreferencia = IIf(IsNull(oRsTmp2.Fields!NroReferenciaDestino), "", oRsTmp2.Fields!NroReferenciaDestino)
        End If
        oDoSisFuaAtencion.Apaterno = oRsTmp2.Fields!ApellidoPaterno
        oDoSisFuaAtencion.Amaterno = oRsTmp2.Fields!ApellidoMaterno
        oDoSisFuaAtencion.Pnombre = oRsTmp2.Fields!PrimerNombre
        oDoSisFuaAtencion.Onombre = IIf(IsNull(oRsTmp2.Fields!SegundoNombre), "", oRsTmp2.Fields!SegundoNombre)
        oDoSisFuaAtencion.Fuaetnia = oRsTmp2.Fields!IdEtnia 'FRANK SIS
        oDoSisFuaAtencion.Fnacimiento = Format(oRsTmp2.Fields!FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_YMD_SIS)
        
        If ml_TipoServicio = sghConsultaExterna Then
             oDoSisFuaAtencion.FuaObservaciones = lcTxtObservaciones
             
             Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(oRsTmp2.Fields!IdServicioIngreso, oConexion)
             oDoSisFuaAtencion.FuaUPS = oDoServicio.codigoServicioFUA
             Set oDoServicio = Nothing
        Else
             oDoSisFuaAtencion.FuaHospitalizadoFingreso = Format(oRsTmp2.Fields!FechaIngreso, sighentidades.DevuelveFechaSoloFormato_YMD_SIS)
             oDoSisFuaAtencion.FuaHospitalizadoFalta = Format(oRsTmp2.Fields!FechaEgreso, sighentidades.DevuelveFechaSoloFormato_YMD_SIS)
             Set oRsTmp2 = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(oRsTmp2.Fields!IdMedicoEgreso)
             If oRsTmp2.RecordCount > 0 Then
                oDoSisFuaAtencion.FuaMedicoDNI = Left(oRsTmp2.Fields!DNI, 9)
                oDoSisFuaAtencion.FuaMedico = Trim(oRsTmp2.Fields!ApellidoPaterno) & " " & Trim(oRsTmp2.Fields!ApellidoMaterno) & " " & oRsTmp2.Fields!Nombres
                oDoSisFuaAtencion.FuaMedicoTipo = oRsTmp2.Fields!TipoEmpleadoSIS
                
                Set oDoServicio = mo_ReglasServiciosHosp.ServiciosSeleccionarPorId(oRsTmp2.Fields!IdServicioEgreso, oConexion)
                oDoSisFuaAtencion.FuaUPS = oDoServicio.codigoServicioFUA
                Set oDoServicio = Nothing
             End If
        End If
        If oSisFuaAtencion.Modificar(oDoSisFuaAtencion) = True Then
            Set oSisFuaAtencion.Conexion = oConexionExterna
            Set oSisFuaAtencionDIA.Conexion = oConexionExterna
            Set oSisFuaAtencionINS.Conexion = oConexionExterna
            Set oSisFuaAtencionMED.Conexion = oConexionExterna
            Set oSisFuaAtencionPRO.Conexion = oConexionExterna
            Set oSisFuaAtencionSMI.Conexion = oConexionExterna
            oDoSisFuaAtencion.IdCuentaAtencion = ml_IdCuentaAtencion
            oDoSisFuaAtencion.IdUsuarioAuditoria = ml_idUsuario
            FUAgrabaServIntermEnTablas oDoSisFuaAtencion, oSisFuaAtencionDIA, oDoSisFuaAtencionDIA, oRsDx, oSisFuaAtencionINS, _
                                       oDoSisFuaAtencionINS, oRsFarmacia, oSisFuaAtencionMED, oDoSisFuaAtencionMED, _
                                       oSisFuaAtencionSMI, oDoSisFuaAtencionSMI, oRsVacunasSp, oSisFuaAtencionPRO, _
                                       oDoSisFuaAtencionPRO, oRsPatologia, sghModificar, lcInsumo, lcMedicamento
        End If
    End If
    oRsTmp1.Close
    lbTodoOK = True
ErrActHECE:
    If lbTodoOK = True Then
       oConexionExterna.CommitTrans
    Else
       oConexionExterna.RollbackTrans
       
    End If
    oConexion.Close
    oConexionExterna.Close
    Set oConexionExterna = Nothing
    Set oConexion = Nothing
    Set mo_AdminServiciosComunes = Nothing
    Set mo_AdminAdmision = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_ReglasDeProgMedica = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
    Set oRsPatologia = Nothing
    Set oRsFarmacia = Nothing
    Set oRsDx = Nothing
    Set oRsVacunasSp = Nothing
    Set oDoSisFuaAtencion = Nothing: Set oSisFuaAtencion = Nothing
    Set oDoSisFuaAtencionDIA = Nothing: Set oSisFuaAtencionDIA = Nothing
    Set oDoSisFuaAtencionINS = Nothing: Set oSisFuaAtencionINS = Nothing
    Set oDoSisFuaAtencionMED = Nothing: Set oSisFuaAtencionMED = Nothing
    Set oDoSisFuaAtencionPRO = Nothing: Set oSisFuaAtencionPRO = Nothing
    Set oDoSisFuaAtencionSMI = Nothing: Set oSisFuaAtencionSMI = Nothing

End Sub

Sub FUAcreaTemporales(ByRef oRsDx As Recordset, ByRef oRsPatologia As Recordset, ByRef oRsFarmacia As Recordset, _
                      ByRef oRsVacunasSp As Recordset)
    On Error GoTo ErrCrea
    With oRsDx
          .Fields.Append "Id", adInteger
          .Fields.Append "DxNro", adInteger
          .Fields.Append "Descripcion", adVarChar, 255, adFldIsNullable
          .Fields.Append "DxIngresoPresuntivo", adBoolean, adFldIsNullable
          .Fields.Append "DxIngresoDefinitivo", adBoolean, adFldIsNullable
          .Fields.Append "DxIngresoRepetido", adBoolean, adFldIsNullable
          .Fields.Append "DxIngreso", adVarChar, 20, adFldIsNullable
          .Fields.Append "DxEgreso", adVarChar, 20, adFldIsNullable
          .Fields.Append "DxEgresoDefinitivo", adBoolean, adFldIsNullable
          .Fields.Append "DxEgresoRepetido", adBoolean, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
          .AddNew
          .Fields!DxNro = 1
          .AddNew
          .Fields!DxNro = 2
          .AddNew
          .Fields!DxNro = 3
          .AddNew
          .Fields!DxNro = 4
          .AddNew
          .Fields!DxNro = 5
          .AddNew
          .Fields!DxNro = 6
    End With
    '
    With oRsPatologia
          .Fields.Append "Id", adInteger
          .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Procedimiento", adVarChar, 255, adFldIsNullable
          .Fields.Append "Indicado", adInteger
          .Fields.Append "Ejecutado", adInteger
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "IdPuntoCarga", adInteger
          .Fields.Append "DxNro", adInteger
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    '
    With oRsFarmacia
          .Fields.Append "Id", adInteger
          .Fields.Append "Tipo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "MedicInsumo", adVarChar, 255, adFldIsNullable
          .Fields.Append "Recetado", adInteger
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Dx", adVarChar, 20, adFldIsNullable
          .Fields.Append "Precio", adDouble
          .Fields.Append "DxNro", adInteger
          .Fields.Append "formaF", adVarChar, 20, adFldIsNullable
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
    '
    
    With oRsVacunasSp
          .Fields.Append "IntervencionP", adVarChar, 3, adFldIsNullable
          .Fields.Append "Valor", adVarChar, 8, adFldIsNullable
          .Fields.Append "esCheck", adBoolean
          .CursorType = adOpenDynamic
          .LockType = adLockOptimistic
          .Open
    End With
ErrCrea:

End Sub

Sub FuaCargaDxDesdeGAlenHos(ByRef oRsDx As Recordset, oConexion As Connection, ml_idAtencion As Long, _
                              ml_IdTipoServicio As Long, ByRef lcDxPrincipal As String, ByRef lcDxPrincipalNro As Long)
        Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
        Dim oRsTmp1 As New Recordset
        Dim lbNuevoRegistro As Boolean
        Dim lnDxNro As Integer, lcDxPrincipal2 As String, lcDxPrincipalNro2 As Long
        Set oRsTmp1 = mo_AdminAdmision.AtencionesDiagnosticosSeleccionarXidAtencion(ml_idAtencion)
        lcDxPrincipal2 = "": lcDxPrincipalNro2 = 0
        If oRsTmp1.RecordCount > 0 Then
           lnDxNro = 1
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
            oRsDx.MoveFirst
            oRsDx.Find "DxIngreso='" & Trim(oRsTmp1.Fields!codigoCIEsinPto) & "'"
            lbNuevoRegistro = True
            If Not oRsDx.EOF Then
               lbNuevoRegistro = False
            End If
            If lbNuevoRegistro = True Then
              oRsDx.MoveFirst
              oRsDx.Find "DxNro=" & lnDxNro
              If Not oRsDx.EOF Then
                 If ml_IdTipoServicio = sghConsultaExterna Then
                    Select Case oRsTmp1.Fields!IdSubclasificacionDx
                    Case 101
                       oRsDx.Fields!DxIngresoPresuntivo = True
                    Case 102
                       oRsDx.Fields!DxIngresoDefinitivo = True
                       lcDxPrincipal = oRsTmp1.Fields!codigoCIEsinPto
                       lcDxPrincipalNro = lnDxNro
                    Case 103
                       oRsDx.Fields!DxIngresoRepetido = True
                    End Select
                    oRsDx.Fields!dxIngreso = oRsTmp1.Fields!codigoCIEsinPto
                    oRsDx.Fields!Descripcion = oRsTmp1.Fields!Descripcion
                    oRsDx.Update
                    '
                    lcDxPrincipal2 = oRsTmp1.Fields!codigoCIEsinPto
                    lcDxPrincipalNro2 = lnDxNro
                Else
                    If oRsTmp1.Fields!IdSubclasificacionDx = 301 Or oRsTmp1.Fields!IdSubclasificacionDx = 303 Or oRsTmp1.Fields!IdSubclasificacionDx = 402 Or oRsTmp1.Fields!IdSubclasificacionDx = 402 Then
                       oRsDx.Fields!DxEgresoDefinitivo = True
                       oRsDx.Fields!DxEgreso = oRsTmp1.Fields!codigoCIEsinPto
                       lcDxPrincipal = oRsTmp1.Fields!codigoCIEsinPto
                       lcDxPrincipalNro = lnDxNro
                    Else
                       oRsDx.Fields!DxIngresoDefinitivo = True
                       oRsDx.Fields!dxIngreso = oRsTmp1.Fields!codigoCIEsinPto
                    End If
                    oRsDx.Fields!Descripcion = oRsTmp1.Fields!Descripcion
                    oRsDx.Update
                 End If
              End If
              lnDxNro = lnDxNro + 1
            End If
            oRsTmp1.MoveNext
           Loop
           If lcDxPrincipal = "" And lcDxPrincipal2 <> "" Then
              lcDxPrincipal = lcDxPrincipal2
              lcDxPrincipalNro = lcDxPrincipalNro2
           End If
        End If
        Set mo_AdminAdmision = Nothing
End Sub

Sub FuaCargaSIDesdeGAlenHos(ByRef oRsFarmacia As Recordset, ByRef oRsPatologia As Recordset, _
                            mo_lnIdTablaLISTBARITEMS As sghOpcionGalenHos, ml_IdCuentaAtencion As Long, _
                            lcInsumo As String, lcMedicamento As String, lcDxPrincipal As String, _
                            lcDxPrincipalNro As Long, lcOtros As String, lcLaboratorio As String, lcImagenes As String)
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim oRsTmp4 As New Recordset
        Dim mo_AdminServiciosComunes As New ReglasComunes
        Dim mo_ReglasFacturacion As New ReglasFacturacion
        Dim lnRecetado As Long, lnIdPuntoCarga As Long, lcPuntoCarga As String
        Select Case mo_lnIdTablaLISTBARITEMS
        Case sghOpcionGalenHos.sghRegistroCitaCE
        Case sghOpcionGalenHos.sghRegistroAtencionCE
                'Solo Receta
                Set oRsTmp2 = mo_AdminServiciosComunes.RecetaDetalleSoloFarmaciaSeleccionarXidCuentaAtencion(ml_IdCuentaAtencion, False)
                If oRsTmp2.RecordCount > 0 Then
                   oRsTmp2.MoveFirst
                   Do While Not oRsTmp2.EOF
                      If oRsTmp2.Fields!Precio > 0 Then
                         oRsFarmacia.AddNew
                         oRsFarmacia.Fields!id = oRsTmp2.Fields!IdProducto
                         oRsFarmacia.Fields!tipo = IIf(oRsTmp2.Fields!TipoProducto = 1, lcInsumo, lcMedicamento)
                         oRsFarmacia.Fields!MedicInsumo = oRsTmp2.Fields!nombre
                         oRsFarmacia.Fields!Recetado = oRsTmp2.Fields!CantidadPedida
                         oRsFarmacia.Fields!Cantidad = 0
                         oRsFarmacia.Fields!dx = lcDxPrincipal
                         oRsFarmacia.Fields!Precio = oRsTmp2.Fields!Precio
                         oRsFarmacia.Fields!codigo = oRsTmp2.Fields!codigo
                         oRsFarmacia.Fields!DxNro = lcDxPrincipalNro
                         oRsFarmacia.Fields!formaF = oRsTmp2.Fields!FormaFarmaceutica
                         oRsFarmacia.Update
                      End If
                      oRsTmp2.MoveNext
                   Loop
                   oRsFarmacia.Sort = "tipo,medicInsumo"
                End If
                oRsTmp2.Close
                'cpt-Consumo en el Mismo consultorio
                Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorIdCuenta(ml_IdCuentaAtencion)
                oRsTmp2.Filter = "idPuntoCarga=1"
                If oRsTmp2.RecordCount > 0 Then
                   oRsTmp2.MoveFirst
                   Do While Not oRsTmp2.EOF
                      If oRsTmp2.Fields!Precio > 0 Then
                            lnIdPuntoCarga = oRsTmp2.Fields!idPuntoCarga
                            lcPuntoCarga = lcOtros
                            oRsPatologia.AddNew
                            oRsPatologia.Fields!id = oRsTmp2.Fields!IdProducto
                            oRsPatologia.Fields!tipo = lcPuntoCarga
                            oRsPatologia.Fields!procedimiento = oRsTmp2.Fields!nombre
                            oRsPatologia.Fields!Indicado = oRsTmp2.Fields!Cantidad
                            oRsPatologia.Fields!ejecutado = oRsTmp2.Fields!Cantidad
                            oRsPatologia.Fields!dx = lcDxPrincipal
                            oRsPatologia.Fields!Precio = oRsTmp2.Fields!Precio
                            oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                            oRsPatologia.Fields!codigo = oRsTmp2.Fields!codigo
                            oRsPatologia.Fields!DxNro = lcDxPrincipalNro
                            oRsPatologia.Update
                        End If
                        oRsTmp2.MoveNext
                   Loop
                   oRsPatologia.Sort = "tipo,procedimiento"
                End If
                oRsTmp2.Close
                
                '******debb-06/08/2015 (inicio) ******
                'consumos de otras cuentas creadas automaticamente(Inmunizaciones, nutricion, etc)
                Set oRsTmp4 = mo_ReglasAdmision.ServiciosAtenSimultaneaMovXcuenta(ml_IdCuentaAtencion)
                oRsTmp4.Filter = "idCuentaAtencion<>" & ml_IdCuentaAtencion
                If oRsTmp4.RecordCount > 0 Then
                   oRsTmp4.MoveFirst
                   Do While Not oRsTmp4.EOF
                        Set oRsTmp2 = mo_ReglasFacturacion.FacturacionServicioDespachoSeleccionarPorIdCuenta(oRsTmp4!IdCuentaAtencion)
                        oRsTmp2.Filter = "idPuntoCarga=1"
                        If oRsTmp2.RecordCount > 0 Then
                           oRsTmp2.MoveFirst
                           Do While Not oRsTmp2.EOF
                              If oRsTmp2.Fields!Precio > 0 Then
                                    lnIdPuntoCarga = oRsTmp2.Fields!idPuntoCarga
                                    lcPuntoCarga = lcOtros
                                    oRsPatologia.AddNew
                                    oRsPatologia.Fields!id = oRsTmp2.Fields!IdProducto
                                    oRsPatologia.Fields!tipo = lcPuntoCarga
                                    oRsPatologia.Fields!procedimiento = oRsTmp2.Fields!nombre
                                    oRsPatologia.Fields!Indicado = oRsTmp2.Fields!Cantidad
                                    oRsPatologia.Fields!ejecutado = oRsTmp2.Fields!Cantidad
                                    oRsPatologia.Fields!dx = lcDxPrincipal
                                    oRsPatologia.Fields!Precio = oRsTmp2.Fields!Precio
                                    oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                                    oRsPatologia.Fields!codigo = oRsTmp2.Fields!codigo
                                    oRsPatologia.Fields!DxNro = lcDxPrincipalNro
                                    oRsPatologia.Update
                                End If
                                oRsTmp2.MoveNext
                           Loop
                           oRsPatologia.Sort = "tipo,procedimiento"
                        End If
                        oRsTmp2.Close
                        oRsTmp4.MoveNext
                   Loop
                End If
                oRsTmp4.Close
                '******debb-06/08/2015 (fin) ******
                
                'cpt-Solo Receta
                Set oRsTmp2 = mo_AdminServiciosComunes.RecetaDetalleSoloServiciosSeleccionarXidCuentaAtencion(ml_IdCuentaAtencion, False)
                If oRsTmp2.RecordCount > 0 Then
                   oRsTmp2.MoveFirst
                   Do While Not oRsTmp2.EOF
                      If oRsTmp2.Fields!Precio > 0 Then
                         lnIdPuntoCarga = oRsTmp2.Fields!idPuntoCarga
                         lcPuntoCarga = lcOtros
                         Select Case lnIdPuntoCarga
                         Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                            lcPuntoCarga = lcLaboratorio
                         Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                            lcPuntoCarga = lcImagenes
                         End Select
                         oRsPatologia.AddNew
                         oRsPatologia.Fields!id = oRsTmp2.Fields!IdProducto
                         oRsPatologia.Fields!tipo = lcPuntoCarga
                         oRsPatologia.Fields!procedimiento = oRsTmp2.Fields!nombre
                         oRsPatologia.Fields!Indicado = oRsTmp2.Fields!CantidadPedida
                         oRsPatologia.Fields!ejecutado = 0
                         oRsPatologia.Fields!dx = lcDxPrincipal
                         oRsPatologia.Fields!Precio = oRsTmp2.Fields!Precio
                         oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                         oRsPatologia.Fields!codigo = oRsTmp2.Fields!codigo
                         oRsPatologia.Fields!DxNro = lcDxPrincipalNro
                         oRsPatologia.Update
                      End If
                      oRsTmp2.MoveNext
                   Loop
                   oRsPatologia.Sort = "tipo,procedimiento"
                End If
                oRsTmp2.Close
        Case Else    'Hospitalizacion, Emergencia, Consulta Externa (opcion SIS=registro FUA)
                'Farmacia solo despachos/devoluciones
                Set oRsTmp1 = SISDespachosFarmaciaIncluyeDevoluciones(ml_IdCuentaAtencion)
                If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                        If oRsTmp1.Fields!Precio > 0 Then
                            lnRecetado = oRsTmp1.Fields!Cantidad
                            Set oRsTmp2 = mo_AdminServiciosComunes.RecetaDetalleSeleccionarXidCuentaIdItemDocumento(ml_IdCuentaAtencion, oRsTmp1.Fields!IdProducto, oRsTmp1.Fields!DocumentoNumero)
                            If oRsTmp2.RecordCount > 0 Then
                               lnRecetado = oRsTmp2.Fields!CantidadPedida
                            End If
                            oRsTmp2.Close
                            oRsFarmacia.AddNew
                            oRsFarmacia.Fields!id = oRsTmp1.Fields!IdProducto
                            oRsFarmacia.Fields!tipo = IIf(oRsTmp1.Fields!TipoProducto = 1, lcInsumo, lcMedicamento)
                            oRsFarmacia.Fields!MedicInsumo = oRsTmp1.Fields!nombre
                            oRsFarmacia.Fields!Recetado = lnRecetado
                            oRsFarmacia.Fields!Cantidad = oRsTmp1.Fields!Cantidad
                            oRsFarmacia.Fields!dx = lcDxPrincipal
                            oRsFarmacia.Fields!Precio = oRsTmp1.Fields!Precio
                            oRsFarmacia.Fields!codigo = oRsTmp1.Fields!codigo
                            oRsFarmacia.Fields!DxNro = lcDxPrincipalNro
                            oRsFarmacia.Fields!formaF = oRsTmp1.Fields!FormaFarmaceutica
                            oRsFarmacia.Update
                        End If
                        oRsTmp1.MoveNext
                   Loop
                   oRsFarmacia.Sort = "tipo,medicInsumo"
               End If
               oRsTmp1.Close
               '****Procedimientos CPT - Despachos
               Set oRsTmp1 = SISDespachosServiciosIncluyeDevoluciones(ml_IdCuentaAtencion)
               If oRsTmp1.RecordCount > 0 Then
                   oRsTmp1.MoveFirst
                   Do While Not oRsTmp1.EOF
                        If oRsTmp1.Fields!Precio > 0 Then
                            lnIdPuntoCarga = oRsTmp1.Fields!idPuntoCarga
                            lcPuntoCarga = lcOtros
                            Select Case lnIdPuntoCarga
                            Case sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica1, sghPuntosCargaBasicos.sghPtoCargaAnatomiaPatologica2, sghPuntosCargaBasicos.sghPtoCargaBancoSangre1, sghPuntosCargaBasicos.sghPtoCargaBancoSangre2, sghPuntosCargaBasicos.sghPtoCargaPatologiaClinica   'Laboratorio
                                 lcPuntoCarga = lcLaboratorio
                            Case sghPuntosCargaBasicos.sghPtoCargaEcogGeneral, sghPuntosCargaBasicos.sghPtoCargaEcogObstetrica, sghPuntosCargaBasicos.sghPtoCargaRayosX, sghPuntosCargaBasicos.sghPtoCargaTomografia  'Imágenes
                                 lcPuntoCarga = lcImagenes
                            End Select
                            oRsPatologia.AddNew
                            oRsPatologia.Fields!id = oRsTmp1.Fields!IdProducto
                            oRsPatologia.Fields!tipo = lcPuntoCarga
                            oRsPatologia.Fields!procedimiento = oRsTmp1.Fields!nombre
                            oRsPatologia.Fields!Indicado = oRsTmp1.Fields!Cantidad
                            oRsPatologia.Fields!ejecutado = oRsTmp1.Fields!Cantidad
                            oRsPatologia.Fields!dx = lcDxPrincipal
                            oRsPatologia.Fields!Precio = oRsTmp1.Fields!Precio
                            oRsPatologia.Fields!idPuntoCarga = lnIdPuntoCarga
                            oRsPatologia.Fields!codigo = oRsTmp1.Fields!codigo
                            oRsPatologia.Fields!DxNro = lcDxPrincipalNro
                            oRsPatologia.Update
                        End If
                        oRsTmp1.MoveNext
                   Loop
                End If
                oRsTmp1.Close
        End Select
        Set oRsTmp1 = Nothing
        Set oRsTmp2 = Nothing
        Set oRsTmp4 = Nothing
        Set mo_AdminServiciosComunes = Nothing
        Set mo_ReglasFacturacion = Nothing

End Sub

Sub FuaCargaTriajeVacunasDesdeGAlenHos(ByRef txtSPpeso As String, ByRef txtSPtalla As String, ByRef txtSPpa As String, _
                                       ByRef txtObservaciones As String, ml_idAtencion As Long, oConexionExterna As Connection)
       Dim oAtencionesCE As New AtencionesCE
       Dim mo_DOAtencionesCE As New DOAtencionesCE
       txtSPpeso = "": txtSPtalla = "":  txtObservaciones = "": txtSPpa = sighentidades.PresionDevuelveVacia
       mo_DOAtencionesCE.idAtencion = ml_idAtencion
       Set oAtencionesCE.Conexion = oConexionExterna
       If oAtencionesCE.SeleccionarPorId(mo_DOAtencionesCE) = False Then
           Exit Sub
       End If
       If Not IsNull(mo_DOAtencionesCE.TriajePeso) Then
          txtSPpeso = mo_DOAtencionesCE.TriajePeso
       End If
       If Not IsNull(mo_DOAtencionesCE.TriajeTalla) Then
          txtSPtalla = mo_DOAtencionesCE.TriajeTalla
       End If
       
       If Not IsNull(mo_DOAtencionesCE.TriajePresion) Then
          txtSPpa = mo_DOAtencionesCE.TriajePresion
       End If
       If Not IsNull(mo_DOAtencionesCE.CitaObservaciones) Then
          txtObservaciones = mo_DOAtencionesCE.CitaObservaciones
       End If
       Set oAtencionesCE = Nothing
       Set mo_DOAtencionesCE = Nothing
End Sub

Sub FUAgrabaServIntermEnTablas(oDoSisFuaAtencion As DoSisFuaAtencion, _
                          oSisFuaAtencionDIA As SisFuaAtencionDIA, oDoSisFuaAtencionDIA As DoSisFuaAtencionDIA, oRsDx As Recordset, _
                          oSisFuaAtencionINS As SisFuaAtencionINS, oDoSisFuaAtencionINS As DoSisFuaAtencionINS, oRsFarmacia As Recordset, _
                          oSisFuaAtencionMED As SisFuaAtencionMED, oDoSisFuaAtencionMED As DoSisFuaAtencionMED, _
                          oSisFuaAtencionSMI As SisFuaAtencionSMI, oDoSisFuaAtencionSMI As DoSisFuaAtencionSMI, oRsVacunasSp As Recordset, _
                          oSisFuaAtencionPRO As SisFuaAtencionPRO, oDoSisFuaAtencionPRO As DoSisFuaAtencionPRO, oRsCpt As Recordset, _
                          mi_opcion As sghOpciones, lcInsumo As String, lcMedicamento As String)
10        On Error GoTo ErrFuaMod
          'Dx
20        If mi_opcion <> sghAgregar Then
30            If oSisFuaAtencionDIA.SisFuaAtencionDIAeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
40               ms_MensajeError = "DIA: " & oSisFuaAtencionDIA.MensajeError: GoTo ErrFuaMod
50            End If
60        End If
70        If oRsDx.RecordCount > 0 Then
             'oDoSisFuaAtencionDIA.ID
80           oDoSisFuaAtencionDIA.IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
90           oDoSisFuaAtencionDIA.CabDniUsuarioRegistra = oDoSisFuaAtencion.CabDniUsuarioRegistra
100          oDoSisFuaAtencionDIA.CabFechaFuaPrimeraVez = oDoSisFuaAtencion.CabFechaFuaPrimeraVez
110          oDoSisFuaAtencionDIA.CabEstado = oDoSisFuaAtencion.CabEstado
120          oDoSisFuaAtencionDIA.CabNroEnvioAlSIS = oDoSisFuaAtencion.CabNroEnvioAlSIS
130          oDoSisFuaAtencionDIA.CabCodigoPuntoDigitacion = oDoSisFuaAtencion.CabCodigoPuntoDigitacion
140          oDoSisFuaAtencionDIA.CabCodigoUDR = oDoSisFuaAtencion.CabCodigoUDR
150          oDoSisFuaAtencionDIA.FuaDisa = oDoSisFuaAtencion.FuaDisa
160          oDoSisFuaAtencionDIA.FuaLote = oDoSisFuaAtencion.FuaLote
170          oDoSisFuaAtencionDIA.FuaNumero = oDoSisFuaAtencion.FuaNumero
180          oDoSisFuaAtencionDIA.CabOrigenDelRegistro = oDoSisFuaAtencion.CabOrigenDelRegistro
190          oDoSisFuaAtencionDIA.CabVersionAplicativo = oDoSisFuaAtencion.CabVersionAplicativo
200          oDoSisFuaAtencionDIA.CabIdentificacionPaquete = oDoSisFuaAtencion.CabIdentificacionPaquete
210          oRsDx.MoveFirst
220          Do While Not oRsDx.EOF
230             If oRsDx.Fields!dxIngreso <> "" Or oRsDx.Fields!DxEgreso <> "" Then
240                   oDoSisFuaAtencionDIA.DxNumero = oRsDx.Fields!DxNro
250                   oDoSisFuaAtencionDIA.DxTipoIE = IIf(oRsDx.Fields!dxIngreso <> "", "I", IIf(oRsDx.Fields!DxEgreso <> "", "E", "T"))
260                   oDoSisFuaAtencionDIA.DxCodigo = Left(Trim(IIf(oRsDx.Fields!DxEgreso <> "", oRsDx.Fields!DxEgreso, oRsDx.Fields!dxIngreso)), 5)
270                   oDoSisFuaAtencionDIA.DxTipoDPR = IIf(oRsDx.Fields!DxIngresoDefinitivo = True, "D", IIf(oRsDx.Fields!DxEgresoDefinitivo = True, "D", _
                                                       IIf(oRsDx.Fields!DxIngresoPresuntivo = True, "P", IIf(oRsDx.Fields!DxIngresoRepetido = True, "R", _
                                                       IIf(oRsDx.Fields!DxEgresoRepetido = True, "R", " ")))))
280                   If Not oSisFuaAtencionDIA.Insertar(oDoSisFuaAtencionDIA) Then
290                          ms_MensajeError = "DIA: " & oSisFuaAtencionDIA.MensajeError: GoTo ErrFuaMod
300                   End If
310                   oRsDx.Fields!id = oDoSisFuaAtencionDIA.id
320                   oRsDx.Update
330             End If
340             oRsDx.MoveNext
350          Loop
360       End If
          'Insumos
370       If mi_opcion <> sghAgregar Then
380           If oSisFuaAtencionINS.SisFuaAtencionINSeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
390              ms_MensajeError = "INS: " & oSisFuaAtencionINS.MensajeError: GoTo ErrFuaMod
400           End If
410       End If
420       oRsFarmacia.Filter = "tipo='" & lcInsumo & "'"
430       If oRsFarmacia.RecordCount > 0 Then
             'oDoSisFuaAtencionins.ID
440          oDoSisFuaAtencionINS.IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
450          oDoSisFuaAtencionINS.CabDniUsuarioRegistra = oDoSisFuaAtencion.CabDniUsuarioRegistra
460          oDoSisFuaAtencionINS.CabFechaFuaPrimeraVez = oDoSisFuaAtencion.CabFechaFuaPrimeraVez
470          oDoSisFuaAtencionINS.CabEstado = oDoSisFuaAtencion.CabEstado
480          oDoSisFuaAtencionINS.CabNroEnvioAlSIS = oDoSisFuaAtencion.CabNroEnvioAlSIS
490          oDoSisFuaAtencionINS.CabCodigoPuntoDigitacion = oDoSisFuaAtencion.CabCodigoPuntoDigitacion
500          oDoSisFuaAtencionINS.CabCodigoUDR = oDoSisFuaAtencion.CabCodigoUDR
510          oDoSisFuaAtencionINS.FuaDisa = oDoSisFuaAtencion.FuaDisa
520          oDoSisFuaAtencionINS.FuaLote = oDoSisFuaAtencion.FuaLote
530          oDoSisFuaAtencionINS.FuaNumero = oDoSisFuaAtencion.FuaNumero
540          oDoSisFuaAtencionINS.CabOrigenDelRegistro = oDoSisFuaAtencion.CabOrigenDelRegistro
550          oDoSisFuaAtencionINS.CabVersionAplicativo = oDoSisFuaAtencion.CabVersionAplicativo
560          oDoSisFuaAtencionINS.CabIdentificacionPaquete = oDoSisFuaAtencion.CabIdentificacionPaquete
570          oDoSisFuaAtencionINS.IdUsuarioAuditoria = oDoSisFuaAtencion.IdUsuarioAuditoria
580          oRsFarmacia.MoveFirst
590          Do While Not oRsFarmacia.EOF
600             oRsDx.MoveFirst
610             oRsDx.Find "DxNro=" & oRsFarmacia.Fields!DxNro
620             oDoSisFuaAtencionINS.idTablaDx = oRsDx.Fields!id
630             oDoSisFuaAtencionINS.codigo = Left(oRsFarmacia.Fields!codigo, 5)
640             oDoSisFuaAtencionINS.DxNumero = oRsFarmacia.Fields!DxNro
650             oDoSisFuaAtencionINS.CantidadPrescrita = oRsFarmacia.Fields!Recetado
660             oDoSisFuaAtencionINS.CantidadEntregada = oRsFarmacia.Fields!Cantidad
                'oDoSisFuaAtencionins.PrecioUnitario = oRsFarmacia.Fields!precio
670             If Not oSisFuaAtencionINS.Insertar(oDoSisFuaAtencionINS) Then
680                    ms_MensajeError = "INS: " & oSisFuaAtencionINS.MensajeError: GoTo ErrFuaMod
690             End If
700             oRsFarmacia.MoveNext
710          Loop
720       End If
          'Medicamentos
730       If oSisFuaAtencionMED.SisFuaAtencionMEDeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
740          ms_MensajeError = "INS: " & oSisFuaAtencionMED.MensajeError
750          GoTo ErrFuaMod
760       End If
770       oRsFarmacia.Filter = "tipo='" & lcMedicamento & "'"
780       If oRsFarmacia.RecordCount > 0 Then
             'oDoSisFuaAtencionMED.ID
790          oDoSisFuaAtencionMED.IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
800          oDoSisFuaAtencionMED.CabDniUsuarioRegistra = oDoSisFuaAtencion.CabDniUsuarioRegistra
810          oDoSisFuaAtencionMED.CabFechaFuaPrimeraVez = oDoSisFuaAtencion.CabFechaFuaPrimeraVez
820          oDoSisFuaAtencionMED.CabEstado = oDoSisFuaAtencion.CabEstado
830          oDoSisFuaAtencionMED.CabNroEnvioAlSIS = oDoSisFuaAtencion.CabNroEnvioAlSIS
840          oDoSisFuaAtencionMED.CabCodigoPuntoDigitacion = oDoSisFuaAtencion.CabCodigoPuntoDigitacion
850          oDoSisFuaAtencionMED.CabCodigoUDR = oDoSisFuaAtencion.CabCodigoUDR
860          oDoSisFuaAtencionMED.FuaDisa = oDoSisFuaAtencion.FuaDisa
870          oDoSisFuaAtencionMED.FuaLote = oDoSisFuaAtencion.FuaLote
880          oDoSisFuaAtencionMED.FuaNumero = oDoSisFuaAtencion.FuaNumero
890          oDoSisFuaAtencionMED.CabOrigenDelRegistro = oDoSisFuaAtencion.CabOrigenDelRegistro
900          oDoSisFuaAtencionMED.CabVersionAplicativo = oDoSisFuaAtencion.CabVersionAplicativo
910          oDoSisFuaAtencionMED.CabIdentificacionPaquete = oDoSisFuaAtencion.CabIdentificacionPaquete
920          oRsFarmacia.MoveFirst
930          Do While Not oRsFarmacia.EOF
940             oRsDx.MoveFirst
950             oRsDx.Find "DxNro=" & oRsFarmacia.Fields!DxNro
960             oDoSisFuaAtencionMED.idTablaDx = oRsDx.Fields!id
970             oDoSisFuaAtencionMED.codigo = Left(oRsFarmacia.Fields!codigo, 5)
980             oDoSisFuaAtencionMED.DxNumero = oRsFarmacia.Fields!DxNro
990             oDoSisFuaAtencionMED.CantidadPrescrita = oRsFarmacia.Fields!Recetado
1000            oDoSisFuaAtencionMED.CantidadEntregada = oRsFarmacia.Fields!Cantidad
                'oDoSisFuaAtencionMED.PrecioUnitario = oRsFarmacia.Fields!precio
1010            If Not oSisFuaAtencionMED.Insertar(oDoSisFuaAtencionMED) Then
                      ' ms_MensajeError = "MED: " & oSisFuaAtencionMED.MensajeError: GoTo ErrFuaMod
1020            End If
1030            oRsFarmacia.MoveNext
1040         Loop
1050      End If
          'SMI - Vacunas y  SP
1060      If oSisFuaAtencionSMI.SisFuaAtencionSMIeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
1070         ms_MensajeError = "SMI: " & oSisFuaAtencionSMI.MensajeError: GoTo ErrFuaMod
1080      End If
1090      If oRsVacunasSp.RecordCount > 0 Then
             'oDoSisFuaAtencionSMI.ID
1100         oDoSisFuaAtencionSMI.IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
1110         oDoSisFuaAtencionSMI.CabDniUsuarioRegistra = oDoSisFuaAtencion.CabDniUsuarioRegistra
1120         oDoSisFuaAtencionSMI.CabFechaFuaPrimeraVez = oDoSisFuaAtencion.CabFechaFuaPrimeraVez
1130         oDoSisFuaAtencionSMI.CabEstado = oDoSisFuaAtencion.CabEstado
1140         oDoSisFuaAtencionSMI.CabNroEnvioAlSIS = oDoSisFuaAtencion.CabNroEnvioAlSIS
1150         oDoSisFuaAtencionSMI.CabCodigoPuntoDigitacion = oDoSisFuaAtencion.CabCodigoPuntoDigitacion
1160         oDoSisFuaAtencionSMI.CabCodigoUDR = oDoSisFuaAtencion.CabCodigoUDR
1170         oDoSisFuaAtencionSMI.FuaDisa = oDoSisFuaAtencion.FuaDisa
1180         oDoSisFuaAtencionSMI.FuaLote = oDoSisFuaAtencion.FuaLote
1190         oDoSisFuaAtencionSMI.FuaNumero = oDoSisFuaAtencion.FuaNumero
1200         oDoSisFuaAtencionSMI.CabOrigenDelRegistro = oDoSisFuaAtencion.CabOrigenDelRegistro
1210         oDoSisFuaAtencionSMI.CabVersionAplicativo = oDoSisFuaAtencion.CabVersionAplicativo
1220         oDoSisFuaAtencionSMI.CabIdentificacionPaquete = oDoSisFuaAtencion.CabIdentificacionPaquete
1230         oRsVacunasSp.MoveFirst
1240         Do While Not oRsVacunasSp.EOF
1250            oDoSisFuaAtencionSMI.IntervencionesPreventivas = Left(oRsVacunasSp.Fields!intervencionP, 3)
1260            oDoSisFuaAtencionSMI.Valor = oRsVacunasSp.Fields!Valor
1270            If Not oSisFuaAtencionSMI.Insertar(oDoSisFuaAtencionSMI) Then
1280                   ms_MensajeError = "SMI: " & oSisFuaAtencionSMI.MensajeError: GoTo ErrFuaMod
1290            End If
1300            oRsVacunasSp.MoveNext
1310         Loop
1320      End If
          'Cpt
1330      If oSisFuaAtencionPRO.SisFuaAtencionPROeliminarPorCuenta(oDoSisFuaAtencion.IdCuentaAtencion, oDoSisFuaAtencion.IdUsuarioAuditoria) = False Then
1340         ms_MensajeError = "CPT: " & oSisFuaAtencionPRO.MensajeError: GoTo ErrFuaMod
1350      End If
1360      If oRsCpt.RecordCount > 0 Then
             'oDoSisFuaAtencionPRO.ID
1370         oDoSisFuaAtencionPRO.IdCuentaAtencion = oDoSisFuaAtencion.IdCuentaAtencion
1380         oDoSisFuaAtencionPRO.CabDniUsuarioRegistra = oDoSisFuaAtencion.CabDniUsuarioRegistra
1390         oDoSisFuaAtencionPRO.CabFechaFuaPrimeraVez = oDoSisFuaAtencion.CabFechaFuaPrimeraVez
1400         oDoSisFuaAtencionPRO.CabEstado = oDoSisFuaAtencion.CabEstado
1410         oDoSisFuaAtencionPRO.CabNroEnvioAlSIS = oDoSisFuaAtencion.CabNroEnvioAlSIS
1420         oDoSisFuaAtencionPRO.CabCodigoPuntoDigitacion = oDoSisFuaAtencion.CabCodigoPuntoDigitacion
1430         oDoSisFuaAtencionPRO.CabCodigoUDR = oDoSisFuaAtencion.CabCodigoUDR
1440         oDoSisFuaAtencionPRO.FuaDisa = oDoSisFuaAtencion.FuaDisa
1450         oDoSisFuaAtencionPRO.FuaLote = oDoSisFuaAtencion.FuaLote
1460         oDoSisFuaAtencionPRO.FuaNumero = oDoSisFuaAtencion.FuaNumero
1470         oDoSisFuaAtencionPRO.CabOrigenDelRegistro = oDoSisFuaAtencion.CabOrigenDelRegistro
1480         oDoSisFuaAtencionPRO.CabVersionAplicativo = oDoSisFuaAtencion.CabVersionAplicativo
1490         oDoSisFuaAtencionPRO.CabIdentificacionPaquete = oDoSisFuaAtencion.CabIdentificacionPaquete
      '       oDoSisFuaAtencionPRO.Resultado                                  'Falta preguntar el dato que se enviará
1500         oRsCpt.MoveFirst
1510         Do While Not oRsCpt.EOF
1520            oRsDx.MoveFirst
1530            oRsDx.Find "DxNro=" & oRsCpt.Fields!DxNro
1540            oDoSisFuaAtencionPRO.idTablaDx = oRsDx.Fields!id
1550            oDoSisFuaAtencionPRO.codigo = Left(oRsCpt.Fields!codigo, 15)
1560            oDoSisFuaAtencionPRO.DxNumero = oRsCpt.Fields!DxNro
1570            oDoSisFuaAtencionPRO.CantidadPrescrita = oRsCpt.Fields!Indicado
1580            oDoSisFuaAtencionPRO.CantidadEjecutada = oRsCpt.Fields!ejecutado
                'oDoSisFuaAtencionPRO.PrecioUnitario = oRsCpt.Fields!precio
1590            If Not oSisFuaAtencionPRO.Insertar(oDoSisFuaAtencionPRO) Then
1600                   ms_MensajeError = "MED: " & oSisFuaAtencionPRO.MensajeError: GoTo ErrFuaMod
1610            End If
1620            oRsCpt.MoveNext
1630         Loop
1640      End If
ErrFuaMod:
1650      ms_MensajeError = ms_MensajeError & sighentidades.DevuelveFuenteDeLineaDelError(Erl(), _
                                                           "Sub FUAgrabaServIntermEnTablas", "reglaSisGalenhos.cls") 'debb-02/05/2016
                   
End Sub


Function ChequeaCodigoEstablecimientoAdscripcion(lcCodigoEstablecimientoAdscripcion As String, _
                                                 ml_IdTipoServicio As sghTipoServicio, _
                                                 lnIdOrigenDelPacienteDesdeFUA As Long, _
                                                 lcCodigoPrestacion As String) As String
         Dim oRsTmp1 As New Recordset
         Dim lcSql As String, lcEstablecimientoOrigen As String, lcCodigoSis As String

         ChequeaCodigoEstablecimientoAdscripcion = ""
         If lcBuscaParametro.SeleccionaFilaParametro(326) <> "S" Then
             Exit Function
         End If
         If lcCodigoEstablecimientoAdscripcion = "" Then
            ChequeaCodigoEstablecimientoAdscripcion = "Ese paciente no tiene CODIGO DE ESTABLECIMIENTO DE ADSCRIPCION"
         Else
            If Val(lcBuscaParametro.SeleccionaFilaParametro(280)) <> Val(lcCodigoEstablecimientoAdscripcion) Then
               If lcBuscaParametro.SeleccionaFilaParametro(282) = "S" Then
                    'Centro de Salud
                    If lcCodigoPrestacion <> "" Then
                       If Sis_m_ServiciosDevuelveSiEsPreventivo(lcCodigoPrestacion) <> 1 Then
                          Sis_m_ee_ssDevuelveCodigoSIS lcCodigoEstablecimientoAdscripcion, lcCodigoSis, lcEstablecimientoOrigen
                          ChequeaCodigoEstablecimientoAdscripcion = "El Paciente SIS está ADSCRITO en el Establecimiento: " & _
                                 lcEstablecimientoOrigen & Chr(13) & "sólo se puede usar CODIGO PRESTACION de SERVICIOS PREVENTIVOS - RC01 "
                       End If
                    End If
               Else
                    'Hospital
                    Select Case ml_IdTipoServicio
                    Case sghConsultaExterna
                         If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
                            Sis_m_ee_ssDevuelveCodigoSIS lcCodigoEstablecimientoAdscripcion, lcCodigoSis, lcEstablecimientoOrigen
                            ChequeaCodigoEstablecimientoAdscripcion = "El Paciente SIS está ADSCRITO en el Establecimiento: " & _
                                           lcEstablecimientoOrigen & Chr(13) & "sólo se puede consultar si es REFERIDO"
                         End If
                    Case sghHospitalizacion
                         If Not (lnIdOrigenDelPacienteDesdeFUA = "4" Or lnIdOrigenDelPacienteDesdeFUA = "6") Then 'Referido CE, ContraReferido
                            Sis_m_ee_ssDevuelveCodigoSIS lcCodigoEstablecimientoAdscripcion, lcCodigoSis, lcEstablecimientoOrigen
                            ChequeaCodigoEstablecimientoAdscripcion = "El Paciente SIS está ADSCRITO en el Establecimiento: " & _
                                           lcEstablecimientoOrigen & Chr(13) & "sólo se puede HOSPITALIZAR si es REFERIDO"
                         End If
                    Case sghEmergenciaConsultorios
                    Case Else                       'Algun procedimiento
                    End Select
               End If
            End If
         End If
         Set oRsTmp1 = Nothing
End Function

'debb
Sub Sis_m_ee_ssDevuelveCodigoSIS(lcCodigoRenaes As String, ByRef lcCodigoSis As String, ByRef lcDescripcion As String)
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  Dim sSQL1 As String
  lcCodigoSis = "": lcDescripcion = ""
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)

  With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "m_eessXcodigoRenaes"
'        Set oParameter = .CreateParameter("@lcCodigoRenaes", adVarChar, adParamInput, 50, Left(lcCodigoRenaes & Space(50), 50)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcCodigoRenaes", adVarChar, adParamInput, 50, Format(lcCodigoRenaes, "0000000000")): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
End With
  If oRecordset.RecordCount > 0 Then
       lcCodigoSis = Trim(oRecordset.Fields!pre_IdEESS)
       lcDescripcion = oRecordset.Fields!pre_Nombre
  End If
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Sub
ManejadorDeError:
  MsgBox Err.Description
End Sub

'debb
Sub Sis_m_ee_ssDevuelveCodigoRENAES(lcCodigoSis As String, ByRef lcCodigoRenaes As String, _
                                                       ByRef lcDescripcionRenaes As String)
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  Dim sSQL1 As String
  lcCodigoRenaes = "": lcDescripcionRenaes = ""
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)

  With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_eessSeleccionarPorId"
            Set oParameter = .CreateParameter("@pre_IdEESS", adVarChar, adParamInput, 10, lcCodigoSis): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
  End With

  If oRecordset.RecordCount > 0 Then
       lcCodigoRenaes = Trim(oRecordset.Fields!pre_CodigoRENAES)
       lcDescripcionRenaes = Trim(oRecordset.Fields!pre_Nombre)
  End If
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Sub
ManejadorDeError:
  MsgBox Err.Description

End Sub
'debb
Function Sis_m_ServiciosDevuelveSiEsPreventivo(lcCodigoPrestacion As String) As Integer
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
  Dim oConexion As New Connection
  Dim sSQL1 As String
  Sis_m_ServiciosDevuelveSiEsPreventivo = 0
  oConexion.CommandTimeout = 300
  oConexion.CursorLocation = adUseClient
  oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)

  With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_serviciosXcodSer"
            Set oParameter = .CreateParameter("@codSer", adVarChar, adParamInput, 3, lcCodigoPrestacion): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
  End With
  If oRecordset.RecordCount > 0 Then
       Sis_m_ServiciosDevuelveSiEsPreventivo = oRecordset.Fields!ser_idTipoServicio
  End If
  oConexion.Close
  Set oRecordset = Nothing
  Set oConexion = Nothing
  Set oCommand = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function


Sub Sis_FuaAtencionMarcaAutomaticamente(ByRef lbEsAmbulatoria As Boolean, ByRef lbEsEmergencia As Boolean, _
                                        ByRef lbEsReferencia As Boolean, ml_IdTipoServicio As sghTipoServicio, _
                                        lbEsPacienteReferidoOrigen As Boolean)
    lbEsAmbulatoria = False: lbEsEmergencia = False: lbEsReferencia = False
    If lbEsPacienteReferidoOrigen = True Then
        lbEsReferencia = True
    Else
        Select Case ml_IdTipoServicio
        Case sghHospitalizacion, sghEmergenciaConsultorios
             lbEsEmergencia = True
        Case sghConsultaExterna
             lbEsAmbulatoria = True
        End Select
    End If
End Sub
Sub CrearReporteObservaciones_excel1(mrs_Tmp As Recordset)
'Dim oExcel As Excel.Application
'Dim oWorkBookPlantilla As Workbook
'Dim oWorkBook As Workbook
'Dim oWorkSheet As Worksheet
'Dim mo_ReporteUtil As New SIGHEntidades.ReporteUtil
'Dim iFila As Long
'Dim lnTotal As Double
'        On Error GoTo ManejadorErrorExcel
'        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
'        'Crea nueva hoja
'        Set oWorkBook = oExcel.Workbooks.Add
'        'Abre, copia y cierra la plantilla
'        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
'        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
'        oWorkBookPlantilla.Close
'        'Activa la primera hoja
'        Set oWorkSheet = oWorkBook.Sheets(1)
'        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
'        iFila = 2
'        oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
'        iFila = 4
'        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
'        oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
'        oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
'        oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
'        oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
'        oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
'        oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
'        oWorkSheet.Cells(iFila, 8).Value = "Servicio"
'        oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
'        oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
'        oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        iFila = 6: lnTotal = 0
'        mrs_Tmp.MoveFirst
'        Do While Not mrs_Tmp.EOF
'           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
'           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
'           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
'           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
'           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
'           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
'           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
'           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
'           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("FechaEgreso").Value
'           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("observaciones").Value
'           iFila = iFila + 1
'           mrs_Tmp.MoveNext
'        Loop
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
'        oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
'        oExcel.Visible = True
'        oWorkSheet.PrintPreview
'        'oWorkSheet.PrintOut
'        Set oWorkSheet = Nothing
'        Set oExcel = Nothing
'        Exit Sub
'ManejadorErrorExcel:
'    Select Case Err.Number
'    Case 1004
'        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
'    Case Else
'        MsgBox Err.Description
'        Resume
'    End Select
'    Exit Sub
Dim mo_ReporteUtil As New sighentidades.ReporteUtil
Dim iFila As Long
Dim lnTotal As Double
Dim lcSql As String
Dim lbEsOpenOffice As Boolean
Dim lnHwnd As Long
        On Error GoTo ManejadorErrorExcel
        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
        
        If lbEsOpenOffice = True Then
            Dim ServiceManager As Object
            Dim Desktop As Object
            Dim Document As Object
            Dim Feuille As Object
            Dim Plage As Object
            Dim args()
            Dim Chemin As String
            Dim Fichier As String
            Dim lcArchivoExcel As String
            Dim PrintArea(0)
            Dim Style As Object
            Dim Border As Object
            'encabezado
            Dim PageStyles As Object
            Dim Sheet As Object
            Dim StyleFamilies As Object
            Dim DefPage As Object
            Dim Htext As Object
            Dim Hcontent As Object
            Dim ret As Long
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
        End If
        
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
    
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            'Crea nueva hoja
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
            oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
            oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        iFila = 2
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Lista de Cuentas Observadas (No fueron grabados sus FUAS)")
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
        End If
        iFila = 4
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nro Cuenta")
            Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Apellido Paterno")
            Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Apellido Materno")
            Call Feuille.getcellbyposition(4, iFila - 1).setFormula("Primer Nombre")
            Call Feuille.getcellbyposition(5, iFila - 1).setFormula("N° Historia Clinica")
            Call Feuille.getcellbyposition(6, iFila - 1).setFormula("Tipo Servicio")
            Call Feuille.getcellbyposition(7, iFila - 1).setFormula("Servicio")
            Call Feuille.getcellbyposition(8, iFila - 1).setFormula("Fecha Ingreso")
            Call Feuille.getcellbyposition(9, iFila - 1).setFormula((IIf(IsNull("Fecha Alta"), "", "Fecha Alta")))
            Call Feuille.getcellbyposition(10, iFila - 1).setFormula((IIf(IsNull("Observaciones"), "", "Observaciones")))
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
            oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
            oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
            oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
            oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
            oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
            oWorkSheet.Cells(iFila, 8).Value = "Servicio"
            oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
            oWorkSheet.Cells(iFila, 10).Value = (IIf(IsNull("Fecha Alta"), "", "Fecha Alta"))
            oWorkSheet.Cells(iFila, 11).Value = (IIf(IsNull("Observaciones"), "", "Observaciones"))
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
        End If
        iFila = 6: lnTotal = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(mrs_Tmp.Fields("idCuentaAtencion").Value)
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("ApellidoPaterno").Value)
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("apellidoMaterno").Value)
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("PrimerNombre").Value)
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("nroHistoriaClinica").Value)
                Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp.Fields("tipoServicio").Value)
                Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("Servicio").Value)
                Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp.Fields("FechaIngreso").Value)
                Call Feuille.getcellbyposition(9, iFila - 1).setFormula((IIf(IsNull(mrs_Tmp.Fields("FechaEgreso").Value), "", mrs_Tmp.Fields("FechaEgreso").Value)))
                Call Feuille.getcellbyposition(10, iFila - 1).setFormula((IIf(IsNull(mrs_Tmp.Fields("observaciones").Value), "", mrs_Tmp.Fields("observaciones").Value)))
            Else
                oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
                oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
                oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
                oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
                oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
                oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
                oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
                oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
                oWorkSheet.Cells(iFila, 10).Value = (IIf(IsNull(mrs_Tmp.Fields("FechaEgreso").Value), "", mrs_Tmp.Fields("FechaEgreso").Value))
                oWorkSheet.Cells(iFila, 11).Value = (IIf(IsNull(mrs_Tmp.Fields("observaciones").Value), "", mrs_Tmp.Fields("observaciones").Value))
            End If
            iFila = iFila + 1
            mrs_Tmp.MoveNext
        Loop
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":K" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount)))
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
            oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
        End If
        If lbEsOpenOffice = True Then
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
'            oWorkBook.Close SaveChanges:=False
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            Set oWorkSheet = Nothing
            Set oExcel = Nothing
        End If
        Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select
    Exit Sub

End Sub

'Funcion ya fue reemplazada por Funcion mo_SisConsumoWeb.WebServiceSISBuscarAfiliado Fcv 17072015
Function ConsultaWebSisBuscarAfiliado(lcDNI As String, lcDisa As String, lcLote As String, lcNumero As String, _
                                             lcCorrelativo As String, lcTabla As String, _
                                             wxParametro323 As String) As Recordset
    On Error GoTo IR_ERROR
    
    Dim oSisConsumoWeb As New SIGHNegocios.SisConsumoWeb
    Set ConsultaWebSisBuscarAfiliado = oSisConsumoWeb.WebServiceSISBuscarAfiliado(lcDNI, lcDisa, lcLote, lcNumero, _
                                                                                  lcCorrelativo, lcTabla, wxParametro323)
    
'    Dim clnt
'    Dim strsvc As String
'    Dim Datos() As String
'    Dim Datos2() As String
'    Dim oRsTmpAfiliadoSIS As New Recordset
'    Dim oCrypKey As New CrypKey.Util
'    Dim Security As String
'    With oRsTmpAfiliadoSIS
'          .Fields.Append "ApPaterno", adVarChar, 40, adFldIsNullable
'          .Fields.Append "ApMaterno", adVarChar, 40, adFldIsNullable
'          .Fields.Append "PNombre", adVarChar, 70, adFldIsNullable
'          .Fields.Append "SNombre", adVarChar, 70, adFldIsNullable
'          .Fields.Append "Fnacimiento", adDate, 8, adFldIsNullable
'          .Fields.Append "cAfiliacion", adVarChar, 50, adFldIsNullable
'          .Fields.Append "estadoSis", adVarChar, 1, adFldIsNullable
'          .Fields.Append "fBajaOK", adDate, 8, adFldIsNullable
'          .Fields.Append "DNI", adVarChar, 10, adFldIsNullable
'          .Fields.Append "sexo", adVarChar, 1, adFldIsNullable
'          .Fields.Append "distritoDomicilio", adVarChar, 6, adFldIsNullable
'          .Fields.Append "cDisa", adVarChar, 3, adFldIsNullable
'          .Fields.Append "cFormato", adVarChar, 2, adFldIsNullable
'          .Fields.Append "cNumero", adVarChar, 10, adFldIsNullable
'          .Fields.Append "AfiliacionNroIntegrante", adVarChar, 2, adFldIsNullable
'          .Fields.Append "codigo", adVarChar, 2, adFldIsNullable
'          .Fields.Append "idSiaSis", adInteger
'          .Fields.Append "MotivoBaja", adVarChar, 70, adFldIsNullable
'          .Fields.Append "CodigoEstablAdscripcion", adVarChar, 10, adFldIsNullable
'          .Fields.Append "AfiliacionFecha", adDate, 8, adFldIsNullable
'          .LockType = adLockOptimistic
'          .Open
'    End With
'    Set ConsultaWebSisBuscarAfiliado = oRsTmpAfiliadoSIS
'    '
'    Set clnt = CreateObject("MSSOAP.SoapClient30")
'    clnt.ClientProperty("ServerHTTPRequest") = True
'    clnt.MSSoapInit "http://app.sis.gob.pe/edi/RecepcionTrama.asmx?WSDL"    'clnt.MSSoapInit oCrypKey.DecryptString(wxParametro323)
'    If lcDNI = "" Then
'       strsvc = clnt.BuscarAsegurados(lcDisa, lcLote, lcNumero, lcCorrelativo, lcTabla)
'    Else
'       'strsvc = clnt.BuscarAseguradosDocIdent("00017077", "1", lcDNI)
'    End If
'    '
'    If strsvc <> "" Then 'Valida trama vacia - Actualizado 23092014
'        Datos = Split(strsvc, "|")
'        If Datos(9) <> "" Then
'            oRsTmpAfiliadoSIS.AddNew
'            oRsTmpAfiliadoSIS.Fields!apPaterno = Datos(9)
'            oRsTmpAfiliadoSIS.Fields!apMaterno = Datos(10)
'            oRsTmpAfiliadoSIS.Fields!Pnombre = Datos(11)
'            oRsTmpAfiliadoSIS.Fields!sNombre = Datos(12)
'            oRsTmpAfiliadoSIS.Fields!Sexo = Datos(13)
'            If Datos(14) <> "" Then
'               oRsTmpAfiliadoSIS.Fields!Fnacimiento = CDate(Datos(14))
'            End If
'            oRsTmpAfiliadoSIS.Fields!DistritoDomicilio = Datos(15)
'            oRsTmpAfiliadoSIS.Fields!estadoSis = Datos(16)
'            If Datos(17) <> "" Then
'                If IsDate(Datos(17)) Then
'                    oRsTmpAfiliadoSIS.Fields!fbajaOK = Format(Datos(17), sighentidades.FormatoFechaCorta)
'                Else
'                    oRsTmpAfiliadoSIS.Fields!fbajaOK = "" 'CDate(Datos(17))
'                End If
'            End If
'            If Datos(6) = "1" Then
'               oRsTmpAfiliadoSIS.Fields!DNI = Datos(18)
'            End If
'            oRsTmpAfiliadoSIS.Fields!cDisa = Datos(2)
'            oRsTmpAfiliadoSIS.Fields!cFormato = Datos(3)
'            oRsTmpAfiliadoSIS.Fields!cNumero = Datos(4)
'            oRsTmpAfiliadoSIS.Fields!idSiasis = Val(Datos(0))
'            oRsTmpAfiliadoSIS.Fields!AfiliacionNroIntegrante = Datos(5)
'            oRsTmpAfiliadoSIS.Fields!Codigo = Datos(1)
'
'            If Datos(7) = "" Then
'                'En caso que no devuelva codigo de establecimiento
'                If lcDNI <> "" Then
'                    lcDisa = Datos(2)
'                    lcLote = Datos(3)
'                    lcNumero = Datos(4)
'                    lcCorrelativo = ""
'                    lcTabla = Datos(1)
'                    strsvc = clnt.BuscarAsegurados(lcDisa, lcLote, lcNumero, lcCorrelativo, lcTabla)
'                    Datos2 = Split(strsvc, "|")
'                    If Datos2(9) <> "" Then
'                        oRsTmpAfiliadoSIS.Fields!CodigoEstablAdscripcion = Datos2(7)
'                    Else
'                        oRsTmpAfiliadoSIS.Fields!CodigoEstablAdscripcion = ""
'                    End If
'                Else
'                    oRsTmpAfiliadoSIS.Fields!CodigoEstablAdscripcion = ""
'                End If
'            Else
'                oRsTmpAfiliadoSIS.Fields!CodigoEstablAdscripcion = Datos(7)
'            End If
'
'            oRsTmpAfiliadoSIS.Fields!AfiliacionFecha = Datos(8)
'            oRsTmpAfiliadoSIS.Fields!cAfiliacion = Datos(2) & " " & Datos(3) & " " & Datos(4)
'            oRsTmpAfiliadoSIS.Update
'            Set ConsultaWebSisBuscarAfiliado = oRsTmpAfiliadoSIS
'        End If
'    End If
IR_ERROR:
    If Err.Number <> 0 Then
        Resume Next
        MsgBox Err.Description
    End If
End Function


Sub SisFiliacionesActualizarAfiliadoDesdeWEB(lcDNI As String, lcDisa As String, lcLote As String, lcNumero As String, _
                                             lcCorrelativo As String, lcTabla As String, wxParametro323 As String)
                                             
    Dim oRsBuscaPacientesSis As New Recordset
    If lcDNI <> "" And lcLote <> "R" Then
'        Set oRsBuscaPacientesSis = ConsultaWebSisBuscarAfiliado(lcDNI, "", "", "", "", "", wxParametro323)
        'fcv 17072015
        Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(lcDNI, "", "", "", "", "", wxParametro323)
        If oRsBuscaPacientesSis.RecordCount = 0 Then
            If lcTabla <> "" And lcLote <> "" And lcNumero <> "" Then
'                Set oRsBuscaPacientesSis = ConsultaWebSisBuscarAfiliado("", lcDisa, lcLote, lcNumero, _
'                                                                                      lcCorrelativo, lcTabla, wxParametro323)
                'FCV 17072015
                Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado("", lcDisa, lcLote, lcNumero, _
                                                                                            lcCorrelativo, lcTabla, wxParametro323)
            End If
        End If
    Else
'        Set oRsBuscaPacientesSis = ConsultaWebSisBuscarAfiliado("", lcDisa, lcLote, lcNumero, _
'                                                                          lcCorrelativo, lcTabla, wxParametro323)
        'FCV 17072015
        Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado("", lcDisa, lcLote, lcNumero, _
                                                                                    lcCorrelativo, lcTabla, wxParametro323)
        If oRsBuscaPacientesSis.RecordCount = 0 And lcDNI <> "" Then
            Set oRsBuscaPacientesSis = mo_SisConsumoWeb.WebServiceSISBuscarAfiliado(lcDNI, "", "", "", "", "", wxParametro323)
        End If
    End If
    If oRsBuscaPacientesSis.RecordCount > 0 Then
        Dim oConexionExterna As New Connection
        Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
        Dim ms_MensajeError As String
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        oConexionExterna.CursorLocation = adUseClient
        ms_MensajeError = mo_ReglasSISgalenhos.SisFiliacionesBuscaYactualizaDatosXafiliado(oConexionExterna, _
                                                    oRsBuscaPacientesSis.Fields!idSiasis, _
                                                    oRsBuscaPacientesSis.Fields!codigo, _
                                                    oRsBuscaPacientesSis.Fields!cDisa, _
                                                    oRsBuscaPacientesSis.Fields!cFormato, _
                                                    oRsBuscaPacientesSis.Fields!cNumero, _
                                                    oRsBuscaPacientesSis.Fields!AfiliacionNroIntegrante, _
                                                    IIf(IsNull(oRsBuscaPacientesSis.Fields!DNI), "", "1"), _
                                                    oRsBuscaPacientesSis.Fields!CodigoEstablAdscripcion, _
                                                    oRsBuscaPacientesSis.Fields!AfiliacionFecha, _
                                                    oRsBuscaPacientesSis.Fields!apPaterno, _
                                                    oRsBuscaPacientesSis.Fields!apMaterno, _
                                                    oRsBuscaPacientesSis.Fields!Pnombre, _
                                                    IIf(IsNull(oRsBuscaPacientesSis.Fields!sNombre), "", oRsBuscaPacientesSis.Fields!sNombre), _
                                                    oRsBuscaPacientesSis.Fields!Sexo, _
                                                    oRsBuscaPacientesSis.Fields!Fnacimiento, _
                                                    oRsBuscaPacientesSis.Fields!DistritoDomicilio, _
                                                    oRsBuscaPacientesSis.Fields!estadoSis, _
                                                    IIf(IsNull(oRsBuscaPacientesSis.Fields!fbajaOK), 0, oRsBuscaPacientesSis.Fields!fbajaOK), _
                                                    IIf(IsNull(oRsBuscaPacientesSis.Fields!DNI), "", oRsBuscaPacientesSis.Fields!DNI), _
                                                    IIf(IsNull(oRsBuscaPacientesSis.Fields!MotivoBaja), "", oRsBuscaPacientesSis.Fields!MotivoBaja))
      oConexionExterna.Close
      Set oConexionExterna = Nothing
      Set mo_ReglasSISgalenhos = Nothing
      If ms_MensajeError <> "" Then
         MsgBox ms_MensajeError
      End If
  End If
  Set oRsBuscaPacientesSis = Nothing
End Sub


Function SisFiliacionesXdni(lcDNI As String, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set SisFiliacionesXdni = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "SisFiliacionesXdni"
        Set oParameter = .CreateParameter("@Dni", adVarChar, adParamInput, 10, lcDNI):   .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set SisFiliacionesXdni = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'
Function FuaDefaultsCptFarmaciaSeleccionarTodos(oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set FuaDefaultsCptFarmaciaSeleccionarTodos = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "FuaDefaultsCptFarmaciaSeleccionarTodos"
        Set oRecordset = .Execute
        
   End With
   Set FuaDefaultsCptFarmaciaSeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

'
Function sisFuaAtencionXrangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionXrangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionXrangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionXrangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function sisFuaAtencionDIAxRangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionDIAxRangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionDIAxRangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionDIAxRangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function
'
Function sisFuaAtencionINSxRangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionINSxRangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionINSxRangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionINSxRangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'
Function sisFuaAtencionMEDxRangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionMEDxRangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionMEDxRangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionMEDxRangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'
Function sisFuaAtencionPROxRangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionPROxRangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionPROxRangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionPROxRangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function
'
Function sisFuaAtencionSMIxRangoDeCuentas(lnIdCuentaInicial As Long, lnIdCuentaFinal As Long, oConexionExterna As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionSMIxRangoDeCuentas = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionSMIxRangoDeCuentas"
        Set oParameter = .CreateParameter("@IdCuentaInicial", adInteger, adParamInput, 0, lnIdCuentaInicial): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaFinal", adInteger, adParamInput, 0, lnIdCuentaFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionSMIxRangoDeCuentas = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function sisFuaAtencionXnroAfiliacion(lcDisa As String, lcLote As String, lcNumero As String, oConexionExterna As Connection) As Recordset

On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    Set sisFuaAtencionXnroAfiliacion = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionXnroAfiliacion"
        Set oParameter = .CreateParameter("@Disa", adVarChar, adParamInput, 3, lcDisa):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lote", adVarChar, adParamInput, 2, lcLote):   .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@numero", adVarChar, adParamInput, 16, lcNumero):  .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set sisFuaAtencionXnroAfiliacion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Sub SisFuaAtencionINS_MED_limpiaCantidadEntregada(lnIdCuentaAtencion As Long, oConexionExterna As Connection)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionINS_MED_limpiaCantidadEntregada"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        
   End With
   Set oCommand = Nothing
   Exit Sub
ManejadorDeError:
   MsgBox Err.Description
End Sub
Sub SisFuaAtencionPRO_limpiaCantidadEjecutada(lnIdCuentaAtencion As Long, oConexionExterna As Connection)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionPRO_limpiaCantidadEjecutada"
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, lnIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        
   End With
   Set oCommand = Nothing
   Exit Sub
ManejadorDeError:
   MsgBox Err.Description
End Sub
Function sisFuaAtencionXCuenta(lnIdCuenta As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set sisFuaAtencionXCuenta = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SisFuaAtencionSeleccionarPorId"
        Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, lnIdCuenta): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set sisFuaAtencionXCuenta = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Sub ImportaEESSsis(lcArchivoAcreditadosZIP As String)
     On Error GoTo ErrImporSIS
     Dim lcArchivoTmp As String, lcArchivo As String, lcSql As String
     Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
     Dim mo_ReglasComunes As New ReglasComunes
     Dim Fichero As Integer, Registro As String, Campos() As String, Columna As Single
     Dim lnIdSiaSis As Long, lcCodigo As String, lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String
     Dim lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, lcDocumentoTipo As String
     Dim lcCodigoEstablAdscripcion As String, ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String
     Dim lcPnombre As String, lcOnombres As String, lcGenero As String, ldFnacimiento As Date
     Dim lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, lcDocumentoNumero As String
     Dim lcMotivoBaja As String, lcfecha As String, lcHoraInicioProceso As String, lcHoraFinalProceso As String
     Dim oRsTmp1 As New Recordset, oConexion As New Connection, oConexionSIS As New Connection
     Dim lineas() As String, lnRegistros As Long, lnContador As Long, lcMensajeError As String
     
     Dim lcCodIns As String, lcNombre As String, lcPresen As String, lcConcen As String, lcFormaFarm As String
     Dim lnCosto As Double, lcMesAlta As String, lcMesBaja As String, ldFecBaja As Date, lcDocBaja As String, lcIdEstado As String
     Dim lcIdEESS As String, lcAfilia As String, lcUCI As String, lcIdCategoriaEESS As String
     Dim lcIdDisa  As String, lcIdOdsis As String, lcIdUbigeo As String, lcCodEjeAdm As String, lcVrae As String
     Dim lcUmbral  As String, lcAisped As String, lcEsmn As String, lcCodigoRenaes As String
     Dim lnUltimoIdEstablecimiento As Long
     Dim lcIdFormaFarm As String, lcDescripcion As String, lnTopeMin As Long
     Dim lnTopeMaxNoHosp As Long, lnTopeMaxHops As Long
     Dim lcIdResAtencion As String, lcIdTipoDocumento As String, lcApePaterno As String, lcApeMaterno As String
     Dim lcPriNombre  As String, lcOtrNombre As String, lcIdTipoPersonalSalud As String, lcColegiatura As String
     Dim lcIdEspecialidad  As String, lcNroEspecialidad As String
     Dim lcPetitorio As String, lcPetitorio2010 As String, lcPetitorio2005 As String
     Dim mo_Teclado As New sighentidades.Teclado

     Dim oDOCatalogoBienesInsumos As New DOCatalogoBienesInsumos, oCatalogoBienesInsumos As New CatalogoBienesInsumos
     Dim oEstablecimiento As New Establecimientos, oDOEstablecimiento As New DOEstablecimiento
     Dim oEstablecimientosNoMinsa As New EstablecimientosNoMinsa, oDOEstablecimientoNoMinsa As New DOEstablecimientoNoMinsa
     Dim oDom_insumos As New Dom_insumos, oM_insumos As New m_insumos
     Dim oDom_medicamentos As New Dom_medicamentos, oM_medicamentos As New m_medicamentos
     Dim oDom_eess As New Dom_eess, oM_eess As New m_eess
     Dim oDoa_presentaciones As New Doa_presentaciones, oA_presentaciones As New a_presentaciones
     Dim oDoa_resatencion As New Doa_resatencion, oA_resatencion As New a_resatencion
     
     Dim oCrypKey As New CrypKey.Util
     Dim fila As Long    'fila As Single
     ms_MensajeError = ""
     lcArchivoTmp = "c:\debbSis"
     DOS_CreandoRutaEnDisco lcArchivoTmp
     sighentidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(318)), lcArchivoAcreditadosZIP, lcArchivoTmp, False
     
     oConexionSIS.CommandTimeout = 300
     oConexionSIS.CursorLocation = adUseClient
     oConexionSIS.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
     oConexionSIS.BeginTrans
     
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighentidades.CadenaConexion
     oConexion.BeginTrans
     
     Set oCatalogoBienesInsumos.Conexion = oConexion
     Set oEstablecimiento.Conexion = oConexion
     Set oEstablecimientosNoMinsa.Conexion = oConexion
     
     Set oM_insumos.Conexion = oConexionSIS
     Set oM_medicamentos.Conexion = oConexionSIS
     Set oM_eess.Conexion = oConexionSIS
     Set oA_presentaciones.Conexion = oConexionSIS
     Set oA_resatencion.Conexion = oConexionSIS
     
     '******************INSUMOS**************
     lcArchivo = "insumos_m.txt"
     lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
     lnRegistros = UBound(lineas)
     Fichero = FreeFile
     Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
     lnContador = 0
     While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lcCodIns = ""
        lcNombre = ""
        lcPresen = ""
        lcConcen = ""
        lcFormaFarm = ""
        lnCosto = 0
        lcMesAlta = ""
        lcMesBaja = ""
        ldFecBaja = 0
        lcDocBaja = ""
        lcIdEstado = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                lcCodIns = Campos(Columna)
            Case 1
                lcNombre = Campos(Columna)
            Case 2
                lcPresen = Campos(Columna)
            Case 3
                lcConcen = Campos(Columna)
            Case 4
                lcFormaFarm = Campos(Columna)
            Case 5
                lnCosto = Val(Campos(Columna))
            Case 6
                lcMesAlta = Campos(Columna)
            Case 7
                lcMesBaja = Campos(Columna)
            Case 8
                lcfecha = Campos(Columna)
                lcfecha = Right(lcfecha, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Left(lcfecha, 4)
                If sighentidades.EsFecha(lcfecha, "DD/MM/AAAA") Then
                   ldFecBaja = CDate(lcfecha)
                End If
            Case 9
                lcDocBaja = Campos(Columna)
            Case 10
                lcIdEstado = Campos(Columna)
            End Select
        Next
        Set oRsTmp1 = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodIns, oConexion)
        If oRsTmp1.RecordCount = 0 And Val(lcCodIns) <= 99999 And mo_Teclado.TextoEsSoloNumeros(lcCodIns) Then
            oDOCatalogoBienesInsumos.codigo = lcCodIns
            oDOCatalogoBienesInsumos.Concentracion = Left(lcConcen, 100)
            oDOCatalogoBienesInsumos.Denominacion = Left(lcNombre, 100)
            'oDOCatalogoBienesInsumos.Fabricante = ""
            oDOCatalogoBienesInsumos.FormaFarmaceutica = Left(lcFormaFarm, 10)
            'oDOCatalogoBienesInsumos.IdCentroCosto
            'oDOCatalogoBienesInsumos.IdClasificacionBienInsumo
            'oDOCatalogoBienesInsumos.IdGrupoFarmacologico
            'oDOCatalogoBienesInsumos.IdPaisOrigen
            'oDOCatalogoBienesInsumos.IdPartida
            'oDOCatalogoBienesInsumos.IdProducto
            'oDOCatalogoBienesInsumos.IdSubGrupoFarmacologico
            oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = 3
            oDOCatalogoBienesInsumos.IdUsuarioAuditoria = ml_idUsuario
            'oDOCatalogoBienesInsumos.MaterialEnvase
            oDOCatalogoBienesInsumos.nombre = Left(lcNombre & " " & lcPresen & " " & lcConcen & " " & lcFormaFarm, 300)
            'oDOCatalogoBienesInsumos.NombreComercial
            'oDOCatalogoBienesInsumos.Petitorio
            oDOCatalogoBienesInsumos.Presentacion = Left(lcPresen, 100)
            'oDOCatalogoBienesInsumos.PresentacionEnvase
            oDOCatalogoBienesInsumos.TipoProducto = 1 '0=medi   1=insumo
            If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) = False Then
               ms_MensajeError = ms_MensajeError & Chr(13) & oCatalogoBienesInsumos.MensajeError
            End If
        End If
        oRsTmp1.Close
        'oDom_insumos.ins_CodIns = lcCodIns
        'If oM_insumos.SeleccionarPorId(oDom_insumos) = False Then
           oDom_insumos.IdUsuarioAuditoria = ml_idUsuario
           oDom_insumos.ins_CodIns = lcCodIns
           oDom_insumos.ins_Concen = lcConcen
           oDom_insumos.ins_Costo = lnCosto
           oDom_insumos.ins_DocBaja = lcDocBaja
           oDom_insumos.ins_FecBaja = ldFecBaja
           oDom_insumos.ins_FormaFarmaceutica = lcFormaFarm
           oDom_insumos.ins_IdEstado = lcIdEstado
           oDom_insumos.ins_Nombre = lcNombre
           'oDom_insumos.ins_Observacion
           'oDom_insumos.ins_Petitorio
           oDom_insumos.ins_Presen = lcPresen
           If oM_insumos.Insertar(oDom_insumos) = False Then
         '     ms_MensajeError = ms_MensajeError & Chr(13) & oM_insumos.MensajeError
           End If
        'End If
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "Insumos")
        fila = fila + 1
    Wend
    Close #Fichero
    '******************MEDICAMENTOS**************
    lcArchivo = "medicamentos_m.txt"
    lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
    lnRegistros = UBound(lineas)
    Fichero = FreeFile
    Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
    lnContador = 0
    fila = 0
    While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lcCodIns = ""
        lcNombre = ""
        lcPresen = ""
        lcConcen = ""
        lcFormaFarm = ""
        lnCosto = 0
        lcPetitorio = ""
        lcPetitorio2010 = ""
        lcPetitorio2005 = ""
        lcMesAlta = ""
        lcMesBaja = ""
        ldFecBaja = 0
        lcDocBaja = ""
        lcIdEstado = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                lcCodIns = Campos(Columna)
            Case 1
                lcNombre = Campos(Columna)
            Case 2
                lcPresen = Campos(Columna)
            Case 3
                lcConcen = Campos(Columna)
            Case 4
                lcFormaFarm = Campos(Columna)
            Case 5
                lnCosto = Val(Campos(Columna))
            Case 6
                lcPetitorio = Campos(Columna)
            Case 7
                lcPetitorio2010 = Campos(Columna)
            Case 8
                lcPetitorio2005 = Campos(Columna)
            Case 9
                lcMesAlta = Campos(Columna)
            Case 10
                lcMesBaja = Campos(Columna)
            Case 11
                lcfecha = Campos(Columna)
                lcfecha = Right(lcfecha, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Left(lcfecha, 4)
                If sighentidades.EsFecha(lcfecha, "DD/MM/AAAA") Then
                   ldFecBaja = CDate(lcfecha)
                End If
            Case 12
                lcDocBaja = Campos(Columna)
            Case 13
                lcIdEstado = Campos(Columna)
            End Select
        Next
        Set oRsTmp1 = oCatalogoBienesInsumos.SeleccionarPorCodigo(lcCodIns, oConexion)
        If oRsTmp1.RecordCount = 0 And Val(lcCodIns) <= 99999 And mo_Teclado.TextoEsSoloNumeros(lcCodIns) Then
            oDOCatalogoBienesInsumos.codigo = lcCodIns
            oDOCatalogoBienesInsumos.Concentracion = Left(lcConcen, 100)
            oDOCatalogoBienesInsumos.Denominacion = Left(lcNombre, 100)
            'oDOCatalogoBienesInsumos.Fabricante = ""
            oDOCatalogoBienesInsumos.FormaFarmaceutica = Left(lcFormaFarm, 10)
            'oDOCatalogoBienesInsumos.IdCentroCosto
            'oDOCatalogoBienesInsumos.IdClasificacionBienInsumo
            'oDOCatalogoBienesInsumos.IdGrupoFarmacologico
            'oDOCatalogoBienesInsumos.IdPaisOrigen
            'oDOCatalogoBienesInsumos.IdPartida
            'oDOCatalogoBienesInsumos.IdProducto
            'oDOCatalogoBienesInsumos.IdSubGrupoFarmacologico
            oDOCatalogoBienesInsumos.idTipoSalidaBienInsumo = 3
            oDOCatalogoBienesInsumos.IdUsuarioAuditoria = ml_idUsuario
            'oDOCatalogoBienesInsumos.MaterialEnvase
            oDOCatalogoBienesInsumos.nombre = Left(lcNombre & " " & lcPresen & " " & lcConcen & " " & lcFormaFarm, 300)
            'oDOCatalogoBienesInsumos.NombreComercial
            'oDOCatalogoBienesInsumos.Petitorio
            oDOCatalogoBienesInsumos.Presentacion = Left(lcPresen, 100)
            'oDOCatalogoBienesInsumos.PresentacionEnvase
            oDOCatalogoBienesInsumos.TipoProducto = 0 '0=medi   1=insumo
            If oCatalogoBienesInsumos.Insertar(oDOCatalogoBienesInsumos) = False Then
               ms_MensajeError = ms_MensajeError & Chr(13) & oCatalogoBienesInsumos.MensajeError
            End If
        End If
        oRsTmp1.Close
        'oDom_medicamentos.ins_CodIns = lcCodIns
        'If oM_medicamentos.SeleccionarPorId(oDom_medicamentos) = False Then
           oDom_medicamentos.IdUsuarioAuditoria = ml_idUsuario
           oDom_medicamentos.med_CodMed = lcCodIns
           oDom_medicamentos.med_Concen = lcConcen
           oDom_medicamentos.med_Costo = lnCosto
           oDom_medicamentos.med_FecBaja = ldFecBaja
           oDom_medicamentos.med_FFDigemid = ldFecBaja
           oDom_medicamentos.med_FormaFarmaceutica = lcFormaFarm
           oDom_medicamentos.med_IdEstado = lcIdEstado
           oDom_medicamentos.med_Nombre = lcNombre
           'oDom_medicamentos.ins_Observacion
           oDom_medicamentos.med_Petitorio = lcPetitorio
           oDom_medicamentos.med_Petitorio2005 = lcPetitorio2005
           oDom_medicamentos.med_Petitorio2010 = lcPetitorio2010
           oDom_medicamentos.med_Presen = lcPresen
           If oM_medicamentos.Insertar(oDom_medicamentos) = False Then
         '     ms_MensajeError = ms_MensajeError & Chr(13) & oM_medicamentos.MensajeError
           End If
       ' End If
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "Medicamentos")
        fila = fila + 1
    Wend
    Close #Fichero
    '******************EESS**************
    lcArchivo = "eess_m.txt"
    lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
    lnRegistros = UBound(lineas)
    Fichero = FreeFile
    Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
    lnContador = 0
    fila = 0
    If lnRegistros > 0 Then
        Set oRsTmp1 = mo_ReglasComunes.EstablecimientosSeleccionarPorFiltro(" 1=1 order by idEstablecimiento desc", oConexion)
        lnUltimoIdEstablecimiento = 1
        If oRsTmp1.RecordCount > 0 Then
           lnUltimoIdEstablecimiento = oRsTmp1.Fields!idEstablecimiento + 1
        End If
    End If
    While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lcIdEESS = ""
        lcNombre = ""
        lcAfilia = ""
        lcUCI = ""
        lcIdCategoriaEESS = ""
        lcIdDisa = ""
        lcIdOdsis = ""
        lcIdUbigeo = ""
        lcCodEjeAdm = ""
        lcVrae = ""
        lcUmbral = ""
        lcAisped = ""
        lcEsmn = ""
        lcCodigoRenaes = ""
        lcIdEstado = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                lcIdEESS = Campos(Columna)
            Case 1
                lcNombre = Campos(Columna)
            Case 2
                lcAfilia = Campos(Columna)
            Case 3
                lcUCI = Campos(Columna)
            Case 4
                lcIdCategoriaEESS = Campos(Columna)
            Case 5
                lcIdDisa = Campos(Columna)
            Case 6
                lcIdOdsis = Campos(Columna)
            Case 7
                lcIdUbigeo = Campos(Columna)
            Case 8
                lcCodEjeAdm = Campos(Columna)
            Case 9
                lcVrae = Campos(Columna)
            Case 10
                lcUmbral = Campos(Columna)
            Case 11
                lcAisped = Campos(Columna)
            Case 12
                lcEsmn = Campos(Columna)
            Case 13
                lcCodigoRenaes = Trim(Campos(Columna))
            Case 14
                lcIdEstado = Campos(Columna)
            End Select
        Next
        oDOEstablecimiento.codigo = Right(lcCodigoRenaes, 5)
        oDOEstablecimiento.IdTipo = IIf(InStr(sighentidades.SISdevuelveCategoriaPS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghPS, _
                                    IIf(InStr(sighentidades.SISdevuelveCategoriaCS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghCS, _
                                    IIf(InStr(sighentidades.SISdevuelveCategoriaHospital, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghHospital, sghCategoriaEstablecimiento.sghOtros)))
                                    '3=PS,2=CS,1=Hosp
        If oDOEstablecimiento.IdTipo > 0 Then
           If oEstablecimiento.SeleccionarPorCodigo(oDOEstablecimiento) = False Then
                oDOEstablecimiento.codigo = Right(lcCodigoRenaes, 5)
                oDOEstablecimiento.IdDistrito = Val(lcIdUbigeo)
                oDOEstablecimiento.idEstablecimiento = lnUltimoIdEstablecimiento
                oDOEstablecimiento.IdTipo = IIf(InStr(sighentidades.SISdevuelveCategoriaPS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghPS, _
                                    IIf(InStr(sighentidades.SISdevuelveCategoriaCS, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghCS, _
                                    IIf(InStr(sighentidades.SISdevuelveCategoriaHospital, lcIdCategoriaEESS) > 0, sghCategoriaEstablecimiento.sghHospital, sghCategoriaEstablecimiento.sghOtros)))
                                    '3=PS,2=CS,1=Hosp
                oDOEstablecimiento.IdUsuarioAuditoria = ml_idUsuario
                oDOEstablecimiento.nombre = lcNombre
                If oEstablecimiento.Insertar(oDOEstablecimiento) = False Then
                    ms_MensajeError = ms_MensajeError & Chr(13) & oEstablecimiento.MensajeError
                Else
                    lnUltimoIdEstablecimiento = lnUltimoIdEstablecimiento + 1
                End If
           End If
        Else
           oDOEstablecimientoNoMinsa.codigo = Right(lcCodigoRenaes, 5)
           Set oRsTmp1 = mo_ReglasComunes.EstablecimientosNoMinsaSeleccionarPorCodigo(oDOEstablecimientoNoMinsa.codigo, oConexion)
           oDOEstablecimientoNoMinsa.IdDistrito = Val(lcIdUbigeo)
           oDOEstablecimientoNoMinsa.IdUsuarioAuditoria = ml_idUsuario
           oDOEstablecimientoNoMinsa.nombre = lcNombre
           If oRsTmp1.RecordCount = 0 Then
                'oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa
                oDOEstablecimientoNoMinsa.IdTipoSubsector = 4
                If oEstablecimientosNoMinsa.Insertar(oDOEstablecimientoNoMinsa) = False Then
                End If
           Else
                oDOEstablecimientoNoMinsa.IdEstablecimientoNoMinsa = oRsTmp1.Fields!IdEstablecimientoNoMinsa
                If oEstablecimientosNoMinsa.Modificar(oDOEstablecimientoNoMinsa) = False Then
                End If
           End If
        End If
        'oDom_eess.pre_IdEESS = lcIdEESS
        'If oM_eess.SeleccionarPorId(oDom_eess) = False Then
           oDom_eess.IdUsuarioAuditoria = ml_idUsuario
           oDom_eess.pre_Afilia = lcAfilia
           oDom_eess.pre_Aisped = lcAisped
           oDom_eess.pre_CodEjeAdm = lcCodEjeAdm
           oDom_eess.pre_CodigoRENAES = lcCodigoRenaes
           oDom_eess.pre_esmn = lcEsmn
           oDom_eess.pre_IdCategoriaEESS = lcIdCategoriaEESS
           oDom_eess.pre_IdDisa = lcIdDisa
           oDom_eess.pre_IdEESS = lcIdEESS
           oDom_eess.pre_IdEstado = lcIdEstado
           oDom_eess.pre_IdOdsis = lcIdOdsis
           oDom_eess.pre_IdUbigeo = lcIdUbigeo
           oDom_eess.pre_Nombre = lcNombre
           oDom_eess.pre_UCI = lcUCI
           oDom_eess.pre_Umbral = lcUmbral
           oDom_eess.pre_Vrae = lcVrae
           If oM_eess.Insertar(oDom_eess) = False Then
         '     ms_MensajeError = ms_MensajeError & Chr(13) & oM_eess.MensajeError
           End If
        'End If
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "EESS")
        fila = fila + 1
    Wend
    Close #Fichero
    '******************FormasFarmaceuticas**************
    lcArchivo = "formasFarmaceuticas_m.txt"
    lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
    lnRegistros = UBound(lineas)
    Fichero = FreeFile
    Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
    lnContador = 0
    fila = 0
    While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lcIdFormaFarm = ""
        lcDescripcion = ""
        lcFormaFarm = ""
        lnTopeMin = 0
        lnTopeMaxNoHosp = 0
        lnTopeMaxHops = 0
        lcIdEstado = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                lcIdFormaFarm = Campos(Columna)
            Case 1
                lcDescripcion = Campos(Columna)
            Case 2
                lcFormaFarm = Campos(Columna)
            Case 3
                lnTopeMin = Val(Campos(Columna))
            Case 4
                lnTopeMaxNoHosp = Val(Campos(Columna))
            Case 5
                lnTopeMaxHops = Val(Campos(Columna))
            Case 6
                lcIdEstado = Campos(Columna)
            End Select
        Next
        'oDoa_presentaciones.tpre_IdPresentacion = lcIdFormaFarm
        'If oA_presentaciones.SeleccionarPorId(oDoa_presentaciones) = False Then
           oDoa_presentaciones.IdUsuarioAuditoria = ml_idUsuario
           oDoa_presentaciones.tpre_Abreviatura = lcFormaFarm
           oDoa_presentaciones.tpre_Descripcion = lcDescripcion
           oDoa_presentaciones.tpre_IdEstado = lcIdEstado
           oDoa_presentaciones.tpre_IdPresentacion = lcIdFormaFarm
           oDoa_presentaciones.tpre_TopeHosp = lnTopeMaxHops
           oDoa_presentaciones.tpre_TopeMinimo = lnTopeMin
           oDoa_presentaciones.tpre_TopeNoHosp = lnTopeMaxNoHosp
           If oA_presentaciones.Insertar(oDoa_presentaciones) = False Then
         '     ms_MensajeError = ms_MensajeError & Chr(13) & oA_presentaciones.MensajeError
           End If
        'End If
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "FormaFarm")
        fila = fila + 1
    Wend
    Close #Fichero
    '******************Medicos**************
    lcArchivo = "resAtencion_m.txt"
    lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
    lnRegistros = UBound(lineas)
    Fichero = FreeFile
    Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
    lnContador = 0
    fila = 0
    While Not EOF(Fichero)
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lcIdResAtencion = ""
        lcIdTipoDocumento = ""
        lcApePaterno = ""
        lcApeMaterno = ""
        lcPriNombre = ""
        lcOtrNombre = ""
        lcIdTipoPersonalSalud = ""
        lcColegiatura = ""
        lcIdEspecialidad = ""
        lcNroEspecialidad = ""
        lcIdEstado = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                lcIdResAtencion = Campos(Columna)
            Case 1
                lcIdTipoDocumento = Campos(Columna)
            Case 2
                lcApePaterno = Campos(Columna)
            Case 3
                lcApeMaterno = Campos(Columna)
            Case 4
                lcPriNombre = Campos(Columna)
            Case 5
                lcOtrNombre = Campos(Columna)
            Case 6
                lcIdTipoPersonalSalud = Campos(Columna)
            Case 7
                lcColegiatura = Campos(Columna)
            Case 8
                lcIdEspecialidad = Campos(Columna)
            Case 9
                lcNroEspecialidad = Campos(Columna)
            Case 10
                lcIdEstado = Campos(Columna)
            End Select
        Next
'        oDoa_resatencion.pers_IdTipoDocumento = lcIdTipoDocumento
'        oDoa_resatencion.pers_IdResAtencion = lcIdResAtencion
'        If oA_resatencion.SeleccionarPorId(oDoa_resatencion) = False Then
           oDoa_resatencion.IdUsuarioAuditoria = ml_idUsuario
           oDoa_resatencion.pers_ApeMaterno = lcApeMaterno
           oDoa_resatencion.pers_ApePaterno = lcApePaterno
           oDoa_resatencion.pers_Colegiatura = lcColegiatura
           oDoa_resatencion.pers_IdEspecialidad = lcIdEspecialidad
           oDoa_resatencion.pers_IdEstado = lcIdEstado
           oDoa_resatencion.pers_IdResAtencion = lcIdResAtencion
           oDoa_resatencion.pers_IdTipoDocumento = lcIdTipoDocumento
           oDoa_resatencion.pers_IdTipoPersonalSalud = lcIdTipoPersonalSalud
           oDoa_resatencion.pers_NroEspecialidad = lcNroEspecialidad
           oDoa_resatencion.pers_OtrNombre = lcOtrNombre
           oDoa_resatencion.pers_PriNombre = lcPriNombre
           If oA_resatencion.Insertar(oDoa_resatencion) = False Then
       '       ms_MensajeError = ms_MensajeError & Chr(13) & oA_resatencion.MensajeError
           End If
      '  End If
        RaiseEvent ProgressActualizaValor(fila, lnRegistros, "Medicos")
        fila = fila + 1
    Wend
    Close #Fichero
    '
    oConexionSIS.CommitTrans
    oConexion.CommitTrans
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set mo_ReglasSISgalenhos = Nothing
    oConexionSIS.Close
    Set oConexionSIS = Nothing
    Set oDOCatalogoBienesInsumos = Nothing
    Set oCatalogoBienesInsumos = Nothing
    Set oEstablecimiento = Nothing
    Set oDOEstablecimiento = Nothing
    Set oDom_insumos = Nothing
    Set oM_insumos = Nothing
    Set oDom_medicamentos = Nothing
    Set oM_medicamentos = Nothing
    Set oDom_eess = Nothing
    Set oM_eess = Nothing
    Set oDoa_presentaciones = Nothing
    Set oA_presentaciones = Nothing
    Set oDoa_resatencion = Nothing
    Set oA_resatencion = Nothing
    Set oEstablecimientosNoMinsa = Nothing
    Set oDOEstablecimientoNoMinsa = Nothing
    Set mo_Teclado = Nothing
    
    Exit Sub
ErrImporSIS:
    If ms_MensajeError = "" Then
       ms_MensajeError = Err.Description
    End If
    MsgBox ms_MensajeError
    oConexionSIS.RollbackTrans
    oConexion.RollbackTrans
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set mo_ReglasSISgalenhos = Nothing
    oConexionSIS.Close
    Set oConexionSIS = Nothing
    Set oDOCatalogoBienesInsumos = Nothing
    Set oCatalogoBienesInsumos = Nothing
    Set oEstablecimiento = Nothing
    Set oDOEstablecimiento = Nothing
    Set oDom_insumos = Nothing
    Set oM_insumos = Nothing
    Set oDom_medicamentos = Nothing
    Set oM_medicamentos = Nothing
    Set oDom_eess = Nothing
    Set oM_eess = Nothing
    Set oDoa_presentaciones = Nothing
    Set oA_presentaciones = Nothing
    Set oDoa_resatencion = Nothing
    Set oA_resatencion = Nothing
    Set oEstablecimientosNoMinsa = Nothing
    Set oDOEstablecimientoNoMinsa = Nothing
    Set mo_Teclado = Nothing
    'Resume
End Sub


Function a_resatencionSeleccionarPorDNI(lcDNI As String, Optional oConexionSIS As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexionSIS Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionSIS Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionSIS
        End If
        .CommandTimeout = 150
        .CommandText = "a_resatencionSeleccionarPorId"
        Set oParameter = .CreateParameter("@pers_IdResAtencion", adVarChar, adParamInput, 9, Left(lcDNI, 9)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@pers_IdTipoDocumento", adVarChar, adParamInput, 1, "1"): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set a_resatencionSeleccionarPorDNI = oRecordset
   Set oCommand = Nothing
   Set oConexion = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function



Function PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS(lcDNI As String, lcApaterno As String, lcAmaterno As String, _
         lcPnombre As String, lcSnombre As String, lcIdTipoSexo As String, ldFechaNacimiento As Date, _
         wxParametroJAMO As String, ldFechaActualServidor As Date, ByRef lnIdSiaSis As Long, lcCodigo As String, _
         lbMuestraMensaje As Boolean) As Boolean
    PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS = False
    Dim oRecordset As New Recordset
    Dim lcSql As String
    lnIdSiaSis = 0
    lcCodigo = ""
    If lcDNI <> "" Then
       lcSql = "   where documentoTipo=1 and DocumentoNumero='" & lcDNI & "'"
    Else
       If lcSnombre = "" Then
            lcSql = "   where  paterno='" & lcApaterno & "' and materno='" & lcAmaterno & "' and pnombre='" & lcPnombre & _
                    "'  and genero='" & IIf(Left(lcIdTipoSexo, 1) = "1", "1", "0") & _
                    "' and fnacimiento=convert(datetime,'" & ldFechaNacimiento & "',103)"
       Else
            lcSql = "   where  paterno='" & lcApaterno & "' and materno='" & lcAmaterno & "' and pnombre='" & lcPnombre & _
                    "' and onombres='" & lcSnombre & "' and genero='" & IIf(Left(lcIdTipoSexo, 1) = "1", "1", "0") & _
                    "' and fnacimiento=convert(datetime,'" & ldFechaNacimiento & "',103)"
       End If
    End If
    Set oRecordset = SisFiltraPacientesAfiliados(lcSql, wxParametroJAMO)
    If oRecordset.RecordCount = 0 Then
       If lbMuestraMensaje = True Then
            MsgBox "No se encontró Paciente en tabla de FILIACIONES DEL SIS" & Chr(13) & Chr(13) & _
                   "Por favor busque al Paciente en tabla del SIS AFILIACIONES, pulsando check en   'Buscar en SIS'", vbInformation, "SIS"
       End If
       Set oRecordset = Nothing
       Exit Function
    End If
    If Sis_ValidaSiEsAfiliadoActualDelSIS(oRecordset, ldFechaActualServidor) = False Then
       If lbMuestraMensaje = True Then
          MsgBox "Por favor busque al Paciente en tabla del SIS AFILIACIONES, pulsando check en   'Buscar en SIS'", vbInformation, "SIS"
       End If
       Set oRecordset = Nothing
       Exit Function
    End If
    lnIdSiaSis = oRecordset!idSiasis
    lcCodigo = oRecordset!codigo
    PacienteBuscadoEnTablaGalenHosTieneAfiliacionSIS = True
    Set oRecordset = Nothing
    
End Function

Function sisFuaAtencionUltimoCorrelativoAutomatico(oConexionExterna As Connection) As String
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionUltimoCorrelativoAutomatico"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
      sisFuaAtencionUltimoCorrelativoAutomatico = oRecordset!FuaNumero
   Else
      sisFuaAtencionUltimoCorrelativoAutomatico = ""
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Sub GeneraTXTAtencion(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer, orst As Recordset, lnNroEnvio As String, lnMesEnvio As String, lnAnioEnvio As String)
        On Error GoTo ErrGenTXT
        Dim swValida As Boolean
        Dim sw As Boolean 'variable para verificar palotes en lclinea
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, wlCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long, lcDNI As String, wlNro As Integer
        Const lcSeparadorCampo As String = "|"
        Dim lcDisa As String, lcFormato As String, lcNumero As String, lcSql As String, lnCuenta As Long
        Dim oConexionExterna As New Connection
        Dim oRsAfiliadosSIS As New Recordset
        Dim rstCatalogo As New Recordset
        Dim oDoSisFuaAtencion As New DoSisFuaAtencion, oSisFuaAtencion As New SisFuaAtencion

        Dim oDoCatalogoServicio As New DOCatalogoServicio, oCatalogoServicio As New CatalogoServicios
        Dim tCampoCondicionNombre As String, tCampoCondicionValor As String, tCampoFormatoTexto As String

        Dim lsarr() As String, wlbFlag As Boolean
        Dim lnCont1, lnCont2 As Integer 'Frank 11092014

        Dim fso As Object 'crea objeto para manejar archivos
        Dim act As Object 'crea objeto para manejar archivos
        
        Dim strCampo As String 'para capturar el valor de campo de consistencia
        Dim lbEsOpenOffice As Boolean
        Dim lnHwnd As Long
        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
        
        If lbEsOpenOffice = True Then
            Dim ServiceManager As Object
            Dim Desktop As Object
            Dim Document As Object
            Dim Feuille As Object
            Dim Plage As Object
            Dim args()
            Dim Chemin As String
            Dim Fichier As String
            Dim lcArchivoExcel As String
            Dim PrintArea(0)
            Dim Style As Object
            Dim Border As Object
            'encabezado
            Dim PageStyles As Object
            Dim Sheet As Object
            Dim StyleFamilies As Object
            Dim DefPage As Object
            Dim Htext As Object
            Dim Hcontent As Object
            Dim ret As Long
        Else
            Dim oExcel As Excel.Application
            Dim oWorkBookPlantilla As Workbook
            Dim oWorkBook As Workbook
            Dim oWorkSheet As Worksheet
        End If
            Dim lnFilaExcel As Integer
            Dim lnColumnaExcel As Integer
       
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        oConexionExterna.CursorLocation = adUseClient
        
        Set oSisFuaAtencion.Conexion = oConexionExterna
    
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\Errores Genera SIS.ods"
'                    FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                    Chemin = "file:///" & App.Path & "\Plantillas\"
'                    Chemin = Replace(Chemin, "\", "/")
'                    Fichier = Chemin & "/OpenOffice.ods"
            Fichier = Format(Time, "hhmmss") & ".ods"
            FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
            lcArchivoExcel = Fichier
            Chemin = "file:///" & App.Path & "\Plantillas\"
            Chemin = Replace(Chemin, "\", "/")
            Fichier = Chemin & "/" & lcArchivoExcel
            '
    
            Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
            Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
            Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
            Set Feuille = Document.getSheets().getByIndex(0)
            'Encabezado de Pagina
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            '<<<< crea excel para errores
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
    
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\Errores Genera SIS.xls")
            oWorkBookPlantilla.Worksheets(1).Copy Before:=oWorkBook.Sheets(1)
    
            oWorkBookPlantilla.Close
    
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        Set fso = CreateObject("scripting.filesystemobject") 'creo objeto de manejo de archivos
        Set act = fso.CreateTextFile(lcArchivoTXT, True) 'creo el archivo con nombre asignado por parametro
        lnUltimoCampo = oRsTmp1.Fields.Count
        lnContador = 1
        oRsTmp1.MoveFirst
        lnFilaExcel = 4
        Do While Not oRsTmp1.EOF
           DoEvents: lnContador = lnContador + 1
           lnCuenta = oRsTmp1!IdCuentaAtencion

           lcLinea = ""
           lcDNI = ""
           swValida = True
           sw = True 'inicailizo sw en true
           
           orst.MoveFirst   'mueve el registro al primero para iniciar la comparacion
'           For lnCont = 0 To lnUltimoCampo - 1
            Do While Not orst.EOF 'Frank 10092014
            
                For lnCont1 = 0 To lnUltimoCampo - 1
                    If oRsTmp1.Fields(lnCont1).Name = orst.Fields!Campo Then
                        lnCont = lnCont1
                        Exit For
                    End If
                Next
'********************************************************************modificacion para reducir campos
                If (orst.Fields!Estado.Value = True) Then 'genera el filtro para solo los activos de recordset de estadostrama
                    If UCase(oRsTmp1.Fields(lnCont).Name) = "CABNROENVIOALSIS" Then
                        oDoSisFuaAtencion.IdCuentaAtencion = lnCuenta
                        If Not oSisFuaAtencion.SeleccionarPorId(oDoSisFuaAtencion) Then
                           MsgBox oSisFuaAtencion.MensajeError: GoTo ErrGenTXT
                        End If
                        oDoSisFuaAtencion.CabNroEnvioAlSIS = lnNroEnvio
                        If Not oSisFuaAtencion.Modificar(oDoSisFuaAtencion) Then
                           MsgBox oSisFuaAtencion.MensajeError: GoTo ErrGenTXT
                        End If
                    End If
                    
                    lcCampoValor = ""
                    If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                         Select Case oRsTmp1.Fields(lnCont).Type
                         Case adDate
                              lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                         Case adChar, adVarChar
                              lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                              If lcCampoValor = lcSeparadorCampo Then
                                 lcCampoValor = "_"
                              End If
                         Case adDouble, adInteger
                              lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                         Case Else
                              lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                         End Select
                    End If
                    If IsNull(oRsTmp1.Fields(lnCont).Value) Or oRsTmp1.Fields(lnCont).Value = "" Then 'Frank caso extra
                        If oRsTmp1.Fields(lnCont).Name = "GrupoEtareo" Then
                            lcCampoValor = "0"
                        End If
                        If oRsTmp1.Fields(lnCont).Name = "FuaConceptoPr" Then
                            lcCampoValor = "1"
                        End If
                        If oRsTmp1.Fields(lnCont).Name = "FuaDestino" Then
                            lcCampoValor = "1"
                        End If
                        If oRsTmp1.Fields(lnCont).Name = "FuaCodigoPrestacion" Then
                            lcCampoValor = "056"
                        End If
                        If oRsTmp1.Fields(lnCont).Name = "Situacion" Then
                            lcCampoValor = "2"
                        End If
                    End If
                    
                    If oRsTmp1.Fields(lnCont).Name = "CabNroEnvioAlSIS" Then
                        lcCampoValor = lnNroEnvio
                    End If
                    If oRsTmp1.Fields(lnCont).Name = "Mes" Then
                        lcCampoValor = lnMesEnvio
                    End If
                    If oRsTmp1.Fields(lnCont).Name = "Anio" Then
                        lcCampoValor = lnAnioEnvio
                    End If
                    
'***********************************comprobacion de dato ok
                    If orst.Fields!obligatorio = "si" Then
                        If (IsNull(lcCampoValor) Or lcCampoValor = "") Then
                            swValida = False
                        Else
                            If Not (orst.Fields!formato = "" Or IsNull(orst.Fields!formato)) Then
                                'lcCampoValor = Format(lcCampoValor, orst.Fields!formato)
                                If oRsTmp1.Fields(lnCont).Type = adChar Or oRsTmp1.Fields(lnCont).Type = adVarChar Then
                                    lcCampoValor = Mid(Format(lcCampoValor, orst.Fields!formato), 1, Len(orst.Fields!formato))
                                Else
                                    lcCampoValor = Format(lcCampoValor, orst.Fields!formato)
                                End If
                            End If
                            swValida = True
                            If Not (orst.Fields!observaciones = "" Or IsNull(orst.Fields!observaciones)) Then
                                ReDim lsarr(0)
                                lsarr = Split(orst.Fields!observaciones, ",")
                                    swValida = compruebaDato(lsarr, lcCampoValor)
                            End If
                        End If
                     End If
                    If orst.Fields!obligatorio = "condicional" Then
                        For wlCont = 0 To lnUltimoCampo - 1
                            If oRsTmp1.Fields(wlCont).Name = orst.Fields!CampoCondicion Then
                                wlNro = wlCont
                                Exit For
                            End If
                        Next wlCont
                        If Not IsNull(orst.Fields!Valor) Then
                            lsarr = Split(orst.Fields!Valor, ",")
                        End If
                        wlbFlag = False
                        For wlCont = 0 To UBound(lsarr)
                            If (oRsTmp1.Fields(wlNro).Value = lsarr(wlCont)) Then wlbFlag = True
                        Next wlCont
                        If wlbFlag Then 'si es obligatorio
                            If (IsNull(lcCampoValor) Or lcCampoValor = "") Then
                                swValida = False
                            Else
                                'pone formato si lo hubiera 'Frank
                                If Not (orst.Fields!formato = "" Or IsNull(orst.Fields!formato)) Then
                                    If oRsTmp1.Fields(lnCont).Type = adChar Or oRsTmp1.Fields(lnCont).Type = adVarChar Then
                                        If oRsTmp1.Fields(lnCont).Name = "FUAMedicoDNI" Or oRsTmp1.Fields(lnCont).Name = "DocumentoNumero" Then
                                            lcCampoValor = Mid(Format(lcCampoValor, "00000000"), 1, 8)
                                        Else
                                            lcCampoValor = Mid(Format(lcCampoValor, orst.Fields!formato), 1, Len(orst.Fields!formato))
                                        End If
                                    Else
                                        lcCampoValor = Format(lcCampoValor, orst.Fields!formato)
                                    End If
                                End If
                                'comprueba si existe en tablas maestras si las tiene valida
                                If Not (orst.Fields!observaciones = "" Or IsNull(orst.Fields!observaciones)) Then
                                    ReDim lsarr(0)
                                    lsarr = Split(orst.Fields!observaciones, ",")
                                    If Not compruebaDato(lsarr, lcCampoValor) Then
                                        lcCampoValor = ""
                                    End If
                                End If
                            End If
                        Else
                            If oRsTmp1.Fields(lnCont).Name = "AfiliacionNroIntegrante" Then
                                lcCampoValor = ""
                            End If
                        End If
                    End If
'***********************************modficacion para modelar los datos de salida segun trama********
                    If lcCampoValor = "__/__/____" Then lcCampoValor = ""
                    If (oRsTmp1.Fields(lnCont).Name = "DxTipoDPR") Then
                        If lcCampoValor = "D" Then lcCampoValor = "1"
                        If lcCampoValor = "P" Then lcCampoValor = "2"
                        If lcCampoValor = "C" Then lcCampoValor = "3"
                        If lcCampoValor = "R" Then lcCampoValor = "4"
                    End If
                    If (oRsTmp1.Fields(lnCont).Name = "FuaAtencionFecha") Or (oRsTmp1.Fields(lnCont).Name = "FuaHospitalizadoFingreso") Then
                        If (IsNull(lcCampoValor)) Or lcCampoValor = "" Then
                            lcCampoValor = ""
                        Else
                            If IsDate(lcCampoValor) Then
                                lcCampoValor = Format(lcCampoValor, sighentidades.DevuelveFechaSoloFormato_DMY)
                            Else
                                lcCampoValor = forDDMMYYYY(lcCampoValor)
                            End If
                        End If
                    End If
                
                    If (oRsTmp1.Fields(lnCont).Name = "FuaHospitalizadoFalta") Or (oRsTmp1.Fields(lnCont).Name = "fnacimiento") Then
                        If (IsNull(lcCampoValor)) Or lcCampoValor = "" Then
                            lcCampoValor = ""
                        Else
                            If IsDate(lcCampoValor) Then
                                lcCampoValor = Format(lcCampoValor, sighentidades.DevuelveFechaSoloFormato_DMY)
                            Else
                                lcCampoValor = forDDMMYYYY(lcCampoValor)
                            End If
                        End If
                    End If
                    'salida de codigo de catalogo de servicios, cambiando por el codigosis que se modifico en tabla
                    If Right(lcArchivoTXT, 7) = "PRO.txt" And oRsTmp1.Fields(lnCont).Name = "Codigo" Then
                        Set rstCatalogo = oCatalogoServicio.SeleccionarPorCodigo(lcCampoValor)
                        lcCampoValor = rstCatalogo.Fields!CodigoSIS
                    End If
                    '*******************************fin modificacion para modelar los datos de salida segun trama********
                    If sw Then
                       lcLinea = lcLinea & lcCampoValor
                       sw = False
                    Else
                       lcLinea = lcLinea & lcSeparadorCampo & lcCampoValor
                    End If
                End If 'cierra el if de filtro
                orst.MoveNext 'mueve el recorset de estado de trama
'           Next
           Loop 'Frank 10092014
           
           lcLinea = lcLinea + Chr(10) 'inserta el caracter cl
        
           If swValida Then
                act.Write (lcLinea) 'graba la linea
           Else
                 For lnCont2 = 0 To lnUltimoCampo - 1
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(lnCont2 + 1, lnFilaExcel - 1).setFormula("'" & oRsTmp1(lnCont2).Value)
                    Else
                        oWorkSheet.Cells(lnFilaExcel, lnCont2 + 2).Value = "'" & oRsTmp1(lnCont2).Value
                    End If
                 Next
                 lnFilaExcel = lnFilaExcel + 1
            End If
            oRsTmp1.MoveNext
        Loop
        act.Close
        'oWorkSheet.PrintOut
        If lbEsOpenOffice = True Then
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
'            oWorkBook.Close SaveChanges:=False
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            Set oWorkSheet = Nothing
            Set oExcel = Nothing
            Set oConexionExterna = Nothing
            Set oRsAfiliadosSIS = Nothing
            Set oDoSisFuaAtencion = Nothing
            Set oSisFuaAtencion = Nothing
        End If
    Exit Sub
ErrGenTXT:
'Resume
        MsgBox Err.Description
        Resume Next
       ' Resume

End Sub

Function compruebaDato(oDatos() As String, wdato As String) As Boolean
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    Dim oRespuesta As Boolean
    
    oConexion.CommandTimeout = 300
    If oDatos(0) = "SIGH_SIS" Then
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    End If
    If oDatos(0) = "SIGH_Externa" Then
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    End If
    oConexion.CursorLocation = adUseClient
    'crear procedimiento de consulta a todos los campos
    With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = oConexion
             .CommandTimeout = 150
             .CommandText = "CompruebaDatos"
             Set oParameter = .CreateParameter("@Tabla", adVarChar, adParamInput, 50, oDatos(1)):   .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@Campo", adVarChar, adParamInput, 50, oDatos(2)):   .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@Valor", adVarChar, adParamInput, 50, wdato):   .Parameters.Append oParameter
             Set oRecordset = .Execute
             
        End With
        If oRecordset.RecordCount > 0 Then oRespuesta = True Else oRespuesta = False
        compruebaDato = oRespuesta
        Set oRecordset = Nothing
End Function


'modifcado por samuel 31/03/2014
Function forDDMMYYYY(lcfecha As String) As String
    Dim res As String
    res = Mid(lcfecha, 7, 2) & "/" & Mid(lcfecha, 5, 2) & "/" & Mid(lcfecha, 1, 4)
    forDDMMYYYY = res
End Function

'modifcado por samuel 31/03/2014
Sub GeneraTXTResumen(oEsRes As DoSisFuaResumen, lcArchivoTXT As String, lcCodigoUDR As String, lcIdPPDD As String)
    On Error GoTo ErrGenTXTResumen
        Dim txtlinea As String
        Dim fso
        Dim act
        Set fso = CreateObject("scripting.filesystemobject")
        Set act = fso.CreateTextFile(lcArchivoTXT, True)
            act.Write (oEsRes.Anio + Chr(10))
            act.Write (Format(oEsRes.Mes, "00") + Chr(10))
            act.Write (Format(oEsRes.NroEnvio, "00") + Chr(10))
            act.Write (oEsRes.NomPaquete + Chr(10))
            act.Write (oEsRes.VersionGTI + Chr(10))
            
            act.Write (lcCodigoUDR + Chr(10))
            act.Write (lcIdPPDD + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write ("0" + Chr(10))
            
            act.Write (Trim(Str(oEsRes.CantFilATE)) + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilSMI)) + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilDIA)) + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilMED)) + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilINS)) + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilPRO)) + Chr(10))
            act.Write ("0" + Chr(10))
            act.Write (Trim(Str(oEsRes.CantFilUSU)) + Chr(10))
            act.Close
        Exit Sub
ErrGenTXTResumen:
'Resume
        MsgBox Err.Description
End Sub
'modifcado por samuel 31/03/2014
Sub GeneraTXTUsuario(oEsRes As DoSisFuaUsuario, lcArchivoTXT As String, lcCodigoUDR As String, lcIdPPDD As String)
    Dim lcSeparadorSaltoLinea As String
    lcSeparadorSaltoLinea = Chr(10)
    On Error GoTo ErrGenTXTUsuario
        Dim sep As String
        Dim txtlinea As String
        Dim fso
        Dim act
        sep = "|"
        Set fso = CreateObject("scripting.filesystemobject")
        Set act = fso.CreateTextFile(lcArchivoTXT, True)

            txtlinea = Trim(oEsRes.DNI) & sep & oEsRes.TipoDoc & sep & oEsRes.ApellidoMat & sep & oEsRes.ApellidoPat & sep
            txtlinea = txtlinea & oEsRes.PrimerNombre & sep & oEsRes.SegundoNombre & sep & Format(oEsRes.NroEnvio, "00") & sep
            txtlinea = txtlinea & lcIdPPDD & sep & lcCodigoUDR & sep & oEsRes.Periodo & sep
            txtlinea = txtlinea & Format(oEsRes.Mes, "00") & sep & "0" & sep & Format(oEsRes.CodigoEstablecimiento, "0000000000") & sep & "" & sep & "0" + Chr(10)
        act.Write (txtlinea)
        act.Close
        Exit Sub
ErrGenTXTUsuario:
Resume
        MsgBox Err.Description
End Sub

Function sisFuaAtencionUltimoCorrelativo(Optional oConexionExterna As Connection) As String
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New Connection
Dim ms_MensajeError  As String
    If oConexionExterna Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    End If
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexionExterna Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexionExterna
        End If
        .CommandTimeout = 150
        .CommandText = "sisFuaAtencionUltimoCorrelativo"
        Set oRecordset = .Execute
   End With
   If oRecordset.RecordCount > 0 Then
      sisFuaAtencionUltimoCorrelativo = oRecordset!FuaNumero
   Else
      sisFuaAtencionUltimoCorrelativo = ""
   End If
   Set oCommand = Nothing
   If oConexionExterna Is Nothing Then
      oConexion.Close
   End If
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

'FUA FRANK
Function SisFuaAtencionConsultarUltimoNumero(mc_NumeroInicio As String, mc_NumeroFinal As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFuaAtencionConsultarUltimoNumero = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = " SisFuaAtencionConsultarUltimoNumero"
        Set oParameter = .CreateParameter("@NumeroInicio", adVarChar, adParamInput, 8, mc_NumeroInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NumeroFinal", adVarChar, adParamInput, 8, mc_NumeroFinal): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaAtencionConsultarUltimoNumero = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description

End Function

'Frank 10112014
Function TotalRegistrosArchivoTexto(nombreFichero As String) As Long
    Dim numlib As Integer, isOpen As Boolean
    Dim VarTexto As String
    Dim LineasArchivo As Long
    On Error GoTo Manejador_Error
    ' Obtengo el siguiente número libre de archivo
    numlib = FreeFile()
    Open nombreFichero For Input As #numlib
    ' Se ha abierto el fichero sin problemas
    isOpen = True
    ' Leo todo el contenido en una única operación
    LineasArchivo = 0
    Do Until EOF(numlib)
        Line Input #numlib, VarTexto
        LineasArchivo = LineasArchivo + 1
    Loop
    TotalRegistrosArchivoTexto = LineasArchivo
Manejador_Error:
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

Sub EliminarArchivosDirectorio(ByVal lcDirectorio As String)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(lcDirectorio) Then
        Kill lcDirectorio & "\*.*"
    End If
    Set fs = Nothing
End Sub

Function RetornarCodigoUDR() As String
    RetornarCodigoUDR = Right("000" & lcBuscaParametro.SeleccionaFilaParametro(305), 3)
End Function

Function RetornarCodigoPPDD() As String
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
Dim lnIdPtoDigitacionSIS As Long
    lnIdPtoDigitacionSIS = CLng(lcBuscaParametro.SeleccionaFilaParametro(304))
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "m_pinstalacionSeleccionarPorId"
        Set oParameter = .CreateParameter("@pin_IdPPDD", adInteger, adParamInput, 0, lnIdPtoDigitacionSIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   RetornarCodigoPPDD = "000"
   If oRecordset.RecordCount > 0 Then
        oRecordset.MoveFirst
        RetornarCodigoPPDD = Right("000" & oRecordset.Fields!pin_pdig, 3)
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function TipoFormatoSIS() As Recordset
     Dim oRecordset As New ADODB.Recordset
     Dim oCommand As New ADODB.Command
     Dim oConexion As New ADODB.Connection
     Dim lcBuscaParametro As New SIGHDatos.Parametros
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "TipoFormatoSIS"
         Set oRecordset = .Execute
         Set oRecordset.ActiveConnection = Nothing
    End With
    Set TipoFormatoSIS = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Set oCommand = Nothing
    Set lcBuscaParametro = Nothing
End Function


Function SisFuaColegiosSeleccionarPorCodigoNombre(mc_colegioCodigo As String, mc_colegio As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFuaColegiosSeleccionarPorCodigoNombre = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "m_IIEESeleccionarTodo"
        Set oParameter = .CreateParameter("@colegioCodigo", adVarChar, adParamInput, 20, mc_colegioCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@colegio", adVarChar, adParamInput, 100, mc_colegio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaColegiosSeleccionarPorCodigoNombre = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function

Function SisFuaColegiosYaEstaEnUso(mc_colegioCodigo As String) As Recordset
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SisFuaColegiosYaEstaEnUso = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = " SisFuaColegiosYaEstaEnUso"
        Set oParameter = .CreateParameter("@colegioCodigo", adVarChar, adParamInput, 20, mc_colegioCodigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SisFuaColegiosYaEstaEnUso = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   MsgBox Err.Description
End Function


Function SisFuaColegioSeccionSeleccionarTodos() As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_IIEE_NivelSeleccionarTodo"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set SisFuaColegioSeccionSeleccionarTodos = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function SisFuaColegioGradoSeleccionarPorNivel(mc_IdNivel As String) As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_IIEE_GradoSeleccionarPorIdNivel"
            Set oParameter = .CreateParameter("@IdNivel", adVarChar, adParamInput, 1, mc_IdNivel): .Parameters.Append oParameter
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set SisFuaColegioGradoSeleccionarPorNivel = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

Function SisFuaColegioTurnoSeleccionarTodos() As ADODB.Recordset

    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRecordset As New Recordset
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = oConexion
            .CommandTimeout = 150
            .CommandText = "m_IIEE_TurnoSeleccionarTodo"
            Set oRecordset = .Execute
            Set oRecordset.ActiveConnection = Nothing
    End With
    Set SisFuaColegioTurnoSeleccionarTodos = oRecordset
    oConexion.Close
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
  Exit Function
ManejadorDeError:
  MsgBox Err.Description

End Function

