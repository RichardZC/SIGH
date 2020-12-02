VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Pacientes
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DoPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim lnNroErrores As Integer
   lnNroErrores = 0
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PacientesAgregar"
           Set oParameter = .CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdPaisNacimiento = 0, Null, oTabla.IdPaisNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdPaisProcedencia = 0, Null, oTabla.IdPaisProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TercerNombre", adVarChar, adParamInput, 40, IIf(oTabla.TercerNombre = "", Null, oTabla.TercerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdProcedencia = 0, Null, oTabla.IdProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGradoInstruccion", adInteger, adParamInput, 0, IIf(oTabla.IdGradoInstruccion = 0, Null, oTabla.IdGradoInstruccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCivil", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCivil = 0, Null, oTabla.IdEstadoCivil)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoOcupacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoOcupacion = 0, Null, oTabla.IdTipoOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoDomicilio = 0, Null, oTabla.IdCentroPobladoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(oTabla.NombrePadre = "", Null, oTabla.NombrePadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(oTabla.NombreMadre = "", Null, oTabla.NombreMadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdPaisDomicilio = 0, Null, oTabla.IdPaisDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoNacimiento = 0, Null, oTabla.IdCentroPobladoNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroPobladoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoProcedencia = 0, Null, oTabla.IdCentroPobladoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoProcedencia = 0, Null, oTabla.IdDistritoProcedencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoDomicilio = 0, Null, oTabla.IdDistritoDomicilio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDistritoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoNacimiento = 0, Null, oTabla.IdDistritoNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, IIf(oTabla.FichaFamiliar = "", Null, oTabla.FichaFamiliar)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEtnia", adVarChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@GrupoSanguineo", adVarChar, adParamInput, 10, IIf(oTabla.GrupoSanguineo = "", Null, oTabla.GrupoSanguineo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FactorRh", adVarChar, adParamInput, 10, IIf(oTabla.FactorRh = "", Null, oTabla.FactorRh)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@UsoWebReniec", adBoolean, adParamInput, 0, IIf(oTabla.UsoWebReniec = True, 1, 0)): .Parameters.Append oParameter
           
            '<(Inicio) Añadido Por: WABG el: 22/10/2020-10:27:47 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
           Set oParameter = .CreateParameter("@validacionReniec", adBoolean, adParamInput, 0, IIf(oTabla.validacionReniec = True, 1, 0)): .Parameters.Append oParameter
            '</(Fin) Añadido Por: WABG el: 22/10/2020-10:27:47 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
           
           Set oParameter = .CreateParameter("@IdIdioma", adInteger, adParamInput, 0, IIf(oTabla.idIdioma = 0, Null, oTabla.idIdioma)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Email", adVarChar, adParamInput, 50, IIf(oTabla.Email = "", Null, oTabla.Email)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madreDocumento", adVarChar, adParamInput, 12, IIf(oTabla.madreDocumento = "", Null, oTabla.madreDocumento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madreApellidoPaterno", adVarChar, adParamInput, 1000, IIf(oTabla.madreApellidoPaterno = "", Null, oTabla.madreApellidoPaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madreApellidoMaterno", adVarChar, adParamInput, 1000, IIf(oTabla.madreApellidoMaterno = "", Null, oTabla.madreApellidoMaterno)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madrePrimerNombre", adVarChar, adParamInput, 1000, IIf(oTabla.madrePrimerNombre = "", Null, oTabla.madrePrimerNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madreSegundoNombre", adVarChar, adParamInput, 1000, IIf(oTabla.madreSegundoNombre = "", Null, oTabla.madreSegundoNombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrdenHijo", adInteger, adParamInput, 0, oTabla.NroOrdenHijo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@madreTipoDocumento", adInteger, adParamInput, 0, oTabla.madreTipoDocumento): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@sector", adVarChar, adParamInput, 2, oTabla.sector): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@sectorista", adInteger, adParamInput, 0, oTabla.sectorista): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
       oTabla.IdPaciente = .Parameters("@IdPaciente")
   End With
 
   Insertar = True
   ms_MensajeError = ""
   GrabaHistorico oTabla
Exit Function
ManejadorDeError:
   If lnNroErrores = 0 And Err.Number = -2147217873 Then
      'El idPaciente ya existe
      ActualizaCorrelativosDeNroHistoria
      lnNroErrores = lnNroErrores + 1
      Resume
   End If
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub GrabaHistorico(oTabla As DoPaciente)
    On Error GoTo ErrHistorico
    Dim lcSql As String
    Dim oRsParametros As New Recordset
    Dim oRsMDB As New Recordset
    Dim oConexionMDB As New Connection
    lcSql = "select ValorTexto from Parametros where idparametro=581"
    oRsParametros.Open lcSql, mo_Conexion, adOpenKeyset, adLockOptimistic
    oConexionMDB.CommandTimeout = 900
    oConexionMDB.CursorLocation = adUseClient
    oConexionMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source =" & Trim(oRsParametros!ValorTexto) & "\parametros.mdb"
    lcSql = "INSERT INTO Pacientes" & _
"           (idPaciente,ApellidoPaterno" & _
"           ,ApellidoMaterno" & _
"           ,PrimerNombre" & _
"           ,SegundoNombre" & _
"           ,FechaNacimiento" & _
"           ,NroDocumento" & _
"           ,DireccionDomicilio" & _
"           ,Autogenerado" & _
"           ,IdTipoSexo" & _
"           ,NroHistoriaClinica" & _
"           ,IdTipoNumeracion" & _
"           ,IdDocIdentidad" & _
"           ,Observacion,fecha0" & _
"          )" & _
"     Values" & _
"           (" & oTabla.IdPaciente & _
"           ,'" & oTabla.ApellidoPaterno & _
"           ','" & oTabla.ApellidoMaterno & _
"           ','" & oTabla.PrimerNombre & _
"           ','" & oTabla.SegundoNombre & _
"           ','" & oTabla.FechaNacimiento & _
"           ','" & oTabla.NroDocumento
   lcSql = lcSql & "','" & oTabla.DireccionDomicilio & _
                "   ','" & oTabla.Autogenerado & _
                "   '," & oTabla.idTipoSexo & _
                "   ," & oTabla.NroHistoriaClinica & _
                "   ," & oTabla.IdTipoNumeracion & _
                "   ," & oTabla.IdDocIdentidad & _
                "   ,'" & oTabla.Observacion & "','" & Format(Date, "dd/mm/yyyy") & _
"'           )"

   oRsMDB.Open lcSql, oConexionMDB, adOpenKeyset, adLockOptimistic
    oConexionMDB.Close
ErrHistorico:
    Set oRsParametros = Nothing
    Set oRsMDB = Nothing
    Set oConexionMDB = Nothing
End Sub

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Pacientes
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoPaciente, lbVerificaPacientesRepetidosAntesDeGrabarDatos As Boolean) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   Modificar = False
   ms_MensajeError = ""
   If lbVerificaPacientesRepetidosAntesDeGrabarDatos = True Then
      ms_MensajeError = PacientesVerificaRepetidosAntesDeGrabarDatos(oTabla)
   End If
   If ms_MensajeError = "" Then
        With oCommand
            .CommandType = adCmdStoredProc
            Set .ActiveConnection = mo_Conexion
            .CommandText = "PacientesModificar"
                Set oParameter = .CreateParameter("@IdPaisNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdPaisNacimiento = 0, Null, oTabla.IdPaisNacimiento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoMaterno = "", Null, oTabla.ApellidoMaterno)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@DireccionDomicilio", adVarChar, adParamInput, 100, IIf(oTabla.DireccionDomicilio = "", Null, oTabla.DireccionDomicilio)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 150, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdPaisProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdPaisProcedencia = 0, Null, oTabla.IdPaisProcedencia)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, IIf(oTabla.ApellidoPaterno = "", Null, oTabla.ApellidoPaterno)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@TercerNombre", adVarChar, adParamInput, 40, IIf(oTabla.TercerNombre = "", Null, oTabla.TercerNombre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Telefono", adVarChar, adParamInput, 10, IIf(oTabla.Telefono = "", Null, oTabla.Telefono)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSexo = 0, Null, oTabla.idTipoSexo)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdProcedencia = 0, Null, oTabla.IdProcedencia)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdGradoInstruccion", adInteger, adParamInput, 0, IIf(oTabla.IdGradoInstruccion = 0, Null, oTabla.IdGradoInstruccion)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdEstadoCivil", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCivil = 0, Null, oTabla.IdEstadoCivil)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdTipoOcupacion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoOcupacion = 0, Null, oTabla.IdTipoOcupacion)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdCentroPobladoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoDomicilio = 0, Null, oTabla.IdCentroPobladoDomicilio)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NombrePadre", adVarChar, adParamInput, 20, IIf(oTabla.NombrePadre = "", Null, oTabla.NombrePadre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NombreMadre", adVarChar, adParamInput, 20, IIf(oTabla.NombreMadre = "", Null, oTabla.NombreMadre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdPaisDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdPaisDomicilio = 0, Null, oTabla.IdPaisDomicilio)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdCentroPobladoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoNacimiento = 0, Null, oTabla.IdCentroPobladoNacimiento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdCentroPobladoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdCentroPobladoProcedencia = 0, Null, oTabla.IdCentroPobladoProcedencia)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdDistritoProcedencia", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoProcedencia = 0, Null, oTabla.IdDistritoProcedencia)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdDistritoDomicilio", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoDomicilio = 0, Null, oTabla.IdDistritoDomicilio)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdDistritoNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdDistritoNacimiento = 0, Null, oTabla.IdDistritoNacimiento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, IIf(oTabla.FichaFamiliar = "", Null, oTabla.FichaFamiliar)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdEtnia", adVarChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@GrupoSanguineo", adVarChar, adParamInput, 10, IIf(oTabla.GrupoSanguineo = "", Null, oTabla.GrupoSanguineo)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@FactorRh", adVarChar, adParamInput, 10, IIf(oTabla.FactorRh = "", Null, oTabla.FactorRh)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@UsoWebReniec", adBoolean, adParamInput, 0, IIf(oTabla.UsoWebReniec = True, 1, 0)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@IdIdioma", adInteger, adParamInput, 0, IIf(oTabla.idIdioma = 0, Null, oTabla.idIdioma)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@Email", adVarChar, adParamInput, 50, IIf(oTabla.Email = "", Null, oTabla.Email)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madreDocumento", adVarChar, adParamInput, 12, IIf(oTabla.madreDocumento = "", Null, oTabla.madreDocumento)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madreApellidoPaterno", adVarChar, adParamInput, 1000, IIf(oTabla.madreApellidoPaterno = "", Null, oTabla.madreApellidoPaterno)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madreApellidoMaterno", adVarChar, adParamInput, 1000, IIf(oTabla.madreApellidoMaterno = "", Null, oTabla.madreApellidoMaterno)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madrePrimerNombre", adVarChar, adParamInput, 1000, IIf(oTabla.madrePrimerNombre = "", Null, oTabla.madrePrimerNombre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madreSegundoNombre", adVarChar, adParamInput, 1000, IIf(oTabla.madreSegundoNombre = "", Null, oTabla.madreSegundoNombre)): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@NroOrdenHijo", adInteger, adParamInput, 0, oTabla.NroOrdenHijo): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@madreTipoDocumento", adInteger, adParamInput, 0, oTabla.madreTipoDocumento): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@sector", adVarChar, adParamInput, 2, oTabla.sector): .Parameters.Append oParameter
                Set oParameter = .CreateParameter("@sectorista", adInteger, adParamInput, 0, oTabla.sectorista): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
            .Parameters.Append oParameter
            .Execute
        End With
        Modificar = True
        ms_MensajeError = ""
   Else
        MsgBox ms_MensajeError
   End If
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PacientesEliminar"
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PacientesSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       CargaDatosDelPaciente oTabla, oRecordset
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

Sub CargaDatosDelPaciente(ByRef oTabla As DoPaciente, oRecordset As Recordset)
       oTabla.IdPaisNacimiento = IIf(IsNull(oRecordset!IdPaisNacimiento), 0, oRecordset!IdPaisNacimiento)
       oTabla.ApellidoMaterno = IIf(IsNull(oRecordset!ApellidoMaterno), "", oRecordset!ApellidoMaterno)
       oTabla.DireccionDomicilio = IIf(IsNull(oRecordset!DireccionDomicilio), "", oRecordset!DireccionDomicilio)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.IdTipoNumeracion = IIf(IsNull(oRecordset!IdTipoNumeracion), 0, oRecordset!IdTipoNumeracion)
       oTabla.IdPaisProcedencia = IIf(IsNull(oRecordset!IdPaisProcedencia), 0, oRecordset!IdPaisProcedencia)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.ApellidoPaterno = IIf(IsNull(oRecordset!ApellidoPaterno), "", oRecordset!ApellidoPaterno)
       oTabla.PrimerNombre = IIf(IsNull(oRecordset!PrimerNombre), "", oRecordset!PrimerNombre)
       oTabla.SegundoNombre = IIf(IsNull(oRecordset!SegundoNombre), "", oRecordset!SegundoNombre)
       oTabla.TercerNombre = IIf(IsNull(oRecordset!TercerNombre), "", oRecordset!TercerNombre)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
       oTabla.Telefono = IIf(IsNull(oRecordset!Telefono), "", oRecordset!Telefono)
       oTabla.Autogenerado = IIf(IsNull(oRecordset!Autogenerado), "", oRecordset!Autogenerado)
       oTabla.idTipoSexo = IIf(IsNull(oRecordset!idTipoSexo), 0, oRecordset!idTipoSexo)
       oTabla.IdProcedencia = IIf(IsNull(oRecordset!IdProcedencia), 0, oRecordset!IdProcedencia)
       oTabla.IdGradoInstruccion = IIf(IsNull(oRecordset!IdGradoInstruccion), 0, oRecordset!IdGradoInstruccion)
       oTabla.IdEstadoCivil = IIf(IsNull(oRecordset!IdEstadoCivil), 0, oRecordset!IdEstadoCivil)
       oTabla.IdDocIdentidad = IIf(IsNull(oRecordset!IdDocIdentidad), 0, oRecordset!IdDocIdentidad)
       oTabla.IdTipoOcupacion = IIf(IsNull(oRecordset!IdTipoOcupacion), 0, oRecordset!IdTipoOcupacion)
       oTabla.IdCentroPobladoDomicilio = IIf(IsNull(oRecordset!IdCentroPobladoDomicilio), 0, oRecordset!IdCentroPobladoDomicilio)
       oTabla.NombrePadre = IIf(IsNull(oRecordset!NombrePadre), "", oRecordset!NombrePadre)
       oTabla.NombreMadre = IIf(IsNull(oRecordset!NombreMadre), "", oRecordset!NombreMadre)
       oTabla.IdPaisDomicilio = IIf(IsNull(oRecordset!IdPaisDomicilio), 0, oRecordset!IdPaisDomicilio)
       oTabla.NroHistoriaClinica = IIf(IsNull(oRecordset!NroHistoriaClinica), 0, oRecordset!NroHistoriaClinica)
       oTabla.IdCentroPobladoNacimiento = IIf(IsNull(oRecordset!IdCentroPobladoNacimiento), 0, oRecordset!IdCentroPobladoNacimiento)
       oTabla.IdCentroPobladoProcedencia = IIf(IsNull(oRecordset!IdCentroPobladoProcedencia), 0, oRecordset!IdCentroPobladoProcedencia)
       oTabla.IdDistritoProcedencia = IIf(IsNull(oRecordset!IdDistritoProcedencia), 0, oRecordset!IdDistritoProcedencia)
       oTabla.IdDistritoDomicilio = IIf(IsNull(oRecordset!IdDistritoDomicilio), 0, oRecordset!IdDistritoDomicilio)
       oTabla.IdDistritoNacimiento = IIf(IsNull(oRecordset!IdDistritoNacimiento), 0, oRecordset!IdDistritoNacimiento)
       oTabla.FichaFamiliar = IIf(IsNull(oRecordset!FichaFamiliar), "", oRecordset!FichaFamiliar)
       oTabla.IdEtnia = IIf(IsNull(oRecordset!IdEtnia), "", oRecordset!IdEtnia)
       oTabla.GrupoSanguineo = IIf(IsNull(oRecordset!GrupoSanguineo), "", oRecordset!GrupoSanguineo)
       oTabla.FactorRh = IIf(IsNull(oRecordset!FactorRh), "", oRecordset!FactorRh)
       oTabla.UsoWebReniec = IIf(IsNull(oRecordset!UsoWebReniec), 0, oRecordset!UsoWebReniec)
       
'<(Inicio) Añadido Por: WABG el: 22/10/2020-10:30:37 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
       oTabla.validacionReniec = IIf(IsNull(oRecordset!validacionReniec), 0, oRecordset!validacionReniec)
'</(Fin) Añadido Por: WABG el: 22/10/2020-10:30:37 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
       
       oTabla.idIdioma = IIf(IsNull(oRecordset!idIdioma), 0, oRecordset!idIdioma)
       oTabla.Email = IIf(IsNull(oRecordset!Email), "", oRecordset!Email)
       oTabla.madreDocumento = IIf(IsNull(oRecordset!madreDocumento), "", oRecordset!madreDocumento)
       oTabla.madreApellidoPaterno = IIf(IsNull(oRecordset!madreApellidoPaterno), "", oRecordset!madreApellidoPaterno)
       oTabla.madreApellidoMaterno = IIf(IsNull(oRecordset!madreApellidoMaterno), "", oRecordset!madreApellidoMaterno)
       oTabla.madrePrimerNombre = IIf(IsNull(oRecordset!madrePrimerNombre), "", oRecordset!madrePrimerNombre)
       oTabla.madreSegundoNombre = IIf(IsNull(oRecordset!madreSegundoNombre), "", oRecordset!madreSegundoNombre)
       oTabla.NroOrdenHijo = IIf(IsNull(oRecordset!NroOrdenHijo), 0, oRecordset!NroOrdenHijo)
       oTabla.madreTipoDocumento = IIf(IsNull(oRecordset!madreTipoDocumento), 0, oRecordset!madreTipoDocumento)
       oTabla.sector = IIf(IsNull(oRecordset!sector), "", oRecordset!sector)
       oTabla.sectorista = IIf(IsNull(oRecordset!sector), 0, oRecordset!sectorista)
       '***En sighDatos.Atenciones, tambien se carga los mismos datos, si hay algun cambio aquí también se
       '   hará en sighDatos.Atenciones
       '***Aqui tambien en: SetDefaults
End Sub


Function SeleccionarPorHistoriaClinicaDefinitiva(ByVal oTabla As DoPaciente) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorHistoriaClinicaDefinitiva = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PacientesSeleccionarPorHistoriaClinicaDefinitiva"
       Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorHistoriaClinicaDefinitiva = True
       CargaDatosDelPaciente oTabla, oRecordset
       
   Else
       SeleccionarPorHistoriaClinicaDefinitiva = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
   Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function



Function Filtrar(ByVal oTabla As DoPaciente, lbUsaApellidoMaternoVacio As Boolean, _
                 lbUsaApellidoPaternoVacio As Boolean, lcApellidoVacio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set Filtrar = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarTodos"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, oTabla.SegundoNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, oTabla.IdDocIdentidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, oTabla.FichaFamiliar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        sWhere = ""
        If lbUsaApellidoPaternoVacio = True Then
           sWhere = sWhere & "ApellidoPaterno='" & lcApellidoVacio & "' and "
        End If
        If lbUsaApellidoMaternoVacio = True Then
           sWhere = sWhere & "ApellidoMaterno='" & lcApellidoVacio & "' and "
        End If
        If sWhere <> "" Then
           oRecordset.Filter = Left(sWhere, Len(sWhere) - 5)
        End If
        
     End With
     Set Filtrar = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
     Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-16/02/2011
Function FiltrarConHistoriasTemporales(ByVal oTabla As DoPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set FiltrarConHistoriasTemporales = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarConHistoriasTemporales"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@dni", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, oTabla.FichaFamiliar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set FiltrarConHistoriasTemporales = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
     Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'debb-16/02/2011
Function FiltrarConHistoriasDefinitivas(ByVal oTabla As DoPaciente, lcSinApellido As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set FiltrarConHistoriasDefinitivas = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarConHistoriasDefinitivas"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@dni", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FichaFamiliar", adVarChar, adParamInput, 20, oTabla.FichaFamiliar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
        sWhere = ""
        If oTabla.ApellidoPaterno = lcSinApellido Then
           sWhere = sWhere & "ApellidoPaterno='" & lcSinApellido & "' and "
        End If
        If oTabla.ApellidoMaterno = lcSinApellido Then
           sWhere = sWhere & "ApellidoMaterno='" & lcSinApellido & "' and "
        End If
        If sWhere <> "" Then
           oRecordset.Filter = Left(sWhere, Len(sWhere) - 5)
        End If
        
     End With
     Set FiltrarConHistoriasDefinitivas = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
     Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ObtenerConElMismoNombre(ByVal oTabla As DoPaciente) As ADODB.Recordset
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
        .CommandText = "PacientesXApellidosYnombres"
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, oTabla.SegundoNombre): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoNombre = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
Function ObtenerConLaMismaHistoriaDefinitiva(ByVal oTabla As DoPaciente) As ADODB.Recordset
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
        .CommandText = "PacientesObtenerConLaMismaHistoriaDefinitiva"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConLaMismaHistoriaDefinitiva = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConElMismoAutogenerado(ByVal oTabla As DoPaciente) As ADODB.Recordset
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
        .CommandText = "PacientesObtenerConElMismoAutogenerado"
        Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, oTabla.Autogenerado): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoAutogenerado = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'***************daniel barrantes**************
'***************tiene CITA el mismo dia en Varios Consultorios CE
'***************
Function TieneCita(FechaIngreso As Date, lIdServicio As Long, lIdPaciente As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
    sSql = ""
    sWhere = ""
    If lIdPaciente > 0 Then
       sSql = sSql + " where dbo.Citas.IdPaciente = '" & lIdPaciente & "' and "
    Else
       sSql = sSql + " where "
    End If
    sSql = sSql + " dbo.Citas.Fecha = '" & FechaIngreso & "'"
    If lIdServicio > 0 Then
       sSql = sSql + " and dbo.Citas.IdServicio = " & lIdServicio
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PacientesTieneCita"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1500, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set TieneCita = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function ActualizarNroHistoriaYTipoNumeracion(ByVal oTabla As DoPaciente) As Boolean
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
        .CommandText = "PacientesActualizarNroHistoriaYTipoNumeracion"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, oTabla.IdTipoNumeracion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ActualizarNroHistoriaYTipoNumeracion = True
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarHistoriasParaAdmision(ByVal lNroHistoriaClinica As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter

     Set FiltrarHistoriasParaAdmision = Nothing
     
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarHistoriasParaAdmision"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, lNroHistoriaClinica): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    
    Set FiltrarHistoriasParaAdmision = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ObtenerIdPacientePorHistoriaClinica(lNroHistoriaClinica As Long, lIdTipoNumeracion As Long) As Long
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
        .CommandText = "PacientesObtenerIdPacientePorHistoriaClinica"
        Set oParameter = .CreateParameter("@lNroHistoriaClinica", adInteger, adParamInput, 0, lNroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lIdTipoNumeracion", adInteger, adParamInput, 0, lIdTipoNumeracion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   ObtenerIdPacientePorHistoriaClinica = 0
   If Not (oRecordset.EOF And oRecordset.BOF) Then
        ObtenerIdPacientePorHistoriaClinica = oRecordset!IdPaciente
   End If
   oRecordset.Close
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SePuedeEliminar(lIdPaciente As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SePuedeEliminar = False
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "PacientesSePuedeEliminar"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Respuesta", adInteger, adParamOutput, 0): .Parameters.Append oParameter
        .Execute
        SePuedeEliminar = .Parameters("@Respuesta").Value = 1
    End With
    ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
       Trace Err.Number & " " + Err.Description, "SePuedeEliminar"
Exit Function
End Function

Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: Pacientes", vbInformation, "Pacientes"
End Sub

Function PacientesVerificaRepetidosAntesDeGrabarDatos(oTabla As DoPaciente) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "PacientesVerificaRepetidosAntesDeGrabarDatos"
        Set oParameter = .CreateParameter("@SinProblemas", adVarChar, adParamOutput, 100): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Autogenerado", adVarChar, adParamInput, 30, IIf(oTabla.Autogenerado = "", Null, oTabla.Autogenerado)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, IIf(oTabla.NroHistoriaClinica = 0, Null, oTabla.NroHistoriaClinica)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoNumeracion = 0, Null, oTabla.IdTipoNumeracion)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oTabla.IdPaciente): .Parameters.Append oParameter
        .Execute
        PacientesVerificaRepetidosAntesDeGrabarDatos = .Parameters("@SinProblemas").Value
    End With
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
       Trace Err.Number & " " + Err.Description, ""
Exit Function
End Function




Function FiltrarTodosSoloHistoriasDefinitivas(ByVal oTabla As DoPaciente, lcSinApellido As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set FiltrarTodosSoloHistoriasDefinitivas = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarTodosSoloHistoriasDefinitivas"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, oTabla.SegundoNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, oTabla.IdDocIdentidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     sWhere = ""
     If lcSinApellido = oTabla.ApellidoPaterno Then
        sWhere = sWhere & " apellidoPaterno='" & lcSinApellido & "' and "
     End If
     If lcSinApellido = oTabla.ApellidoMaterno Then
        sWhere = sWhere & " apellidoMaterno='" & lcSinApellido & "' and "
     End If
     If sWhere <> "" Then
        oRecordset.Filter = Left(sWhere, Len(sWhere) - 4)
     End If
     
     Set FiltrarTodosSoloHistoriasDefinitivas = oRecordset
     ms_MensajeError = ""
     Set oRecordset = Nothing
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function PacientesFiltraPorNroDocumentoYtipo(lcNroDocumento As String, lnIdDocIdentidad As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set PacientesFiltraPorNroDocumentoYtipo = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
        .CommandText = "PacientesFiltraPorNroDocumentoYtipo"
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, lcNroDocumento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, lnIdDocIdentidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set PacientesFiltraPorNroDocumentoYtipo = oRecordset
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function PacientesSeleccionarPorDNI(lcDNI As String, ByRef oTabla As DoPaciente) As Boolean
    PacientesSeleccionarPorDNI = False
    Dim oRsTmp1 As New Recordset
    
    Set oRsTmp1 = PacientesFiltraPorNroDocumentoYtipo(lcDNI, 1)
    If oRsTmp1.RecordCount > 0 Then
       CargaDatosDelPaciente oTabla, oRsTmp1
       PacientesSeleccionarPorDNI = True
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Function

Sub ActualizaCorrelativosDeNroHistoria()
    Dim oRsTmp As New Recordset
    Dim oRsTmp1 As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim lnIdPaciente As Long

    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PacientesSeleccionarXTipoNumeracionHistoriaClinica"
        Set oRsTmp = .Execute
        Set oRsTmp.ActiveConnection = Nothing
    End With
  
    If oRsTmp.RecordCount > 0 Then
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF

          Set oRsTmp1 = HistoriasClinicasXIdPaciente(oRsTmp.Fields!IdPaciente, mo_Conexion)
          
          If oRsTmp1.RecordCount > 0 Then
                Set oCommand = Nothing
                Set oParameter = Nothing
                With oCommand
                      .CommandType = adCmdStoredProc
                      Set .ActiveConnection = mo_Conexion
                      .CommandTimeout = 150
                      .CommandText = "PacientesActualizarNroHistoriaYTipoNumeracion"
                      Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oRsTmp1.Fields!NroHistoriaClinica): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdTipoNumeracion", adInteger, adParamInput, 0, oRsTmp.Fields!IdTipoNumeracion): .Parameters.Append oParameter
                      Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, oRsTmp.Fields!IdPaciente): .Parameters.Append oParameter
                      .Execute
                End With
          
          
             
          End If
          oRsTmp1.Close
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oRsTmp1 = Nothing

End Sub

Function HistoriasClinicasXIdPaciente(lnIdPaciente As Long, oConexion As Connection) As Recordset
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oRsTmp1 As New Recordset
    With oCommand
          .CommandType = adCmdStoredProc
          Set .ActiveConnection = oConexion
          .CommandTimeout = 150
          .CommandText = "HistoriasClinicasXIdPaciente"
          Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
          Set oRsTmp1 = .Execute
    End With
    Set HistoriasClinicasXIdPaciente = oRsTmp1
    Set oCommand = Nothing
    Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Sub SetDefaults(ByVal oTabla As DoPaciente)
   oTabla.ApellidoMaterno = ""
   oTabla.ApellidoPaterno = ""
   oTabla.Autogenerado = ""
   oTabla.DireccionDomicilio = ""
   oTabla.Email = ""
   oTabla.FactorRh = ""
   oTabla.FechaNacimiento = 0
   oTabla.FichaFamiliar = ""
   oTabla.GrupoSanguineo = ""
   oTabla.IdCentroPobladoDomicilio = 0
   oTabla.IdCentroPobladoNacimiento = 0
   oTabla.IdCentroPobladoProcedencia = 0
   oTabla.IdDistritoDomicilio = 0
   oTabla.IdDistritoNacimiento = 0
   oTabla.IdDocIdentidad = 0
   oTabla.IdEstadoCivil = 0
   oTabla.IdEtnia = 0
   oTabla.IdGradoInstruccion = 0
   oTabla.idIdioma = 0
   oTabla.IdPaciente = 0
   oTabla.IdPaisDomicilio = 0
   oTabla.IdPaisNacimiento = 0
   oTabla.IdPaisProcedencia = 0
   oTabla.IdProcedencia = 0
   oTabla.IdTipoNumeracion = 0
   oTabla.IdTipoOcupacion = 0
   oTabla.idTipoSexo = 0
   oTabla.IdUsuarioAuditoria = 0
   oTabla.madreApellidoMaterno = ""
   oTabla.madreApellidoPaterno = ""
   oTabla.madreDocumento = ""
   oTabla.madrePrimerNombre = ""
   oTabla.madreSegundoNombre = ""
   oTabla.madreTipoDocumento = 0
   oTabla.NombreMadre = ""
   oTabla.NombrePadre = ""
   oTabla.NroDocumento = ""
   oTabla.NroHistoriaClinica = 0
   oTabla.NroOrdenHijo = 0
   oTabla.Observacion = ""
   oTabla.PrimerNombre = ""
   oTabla.sector = ""
   oTabla.sectorista = 0
   oTabla.SegundoNombre = ""
   oTabla.Telefono = ""
   oTabla.TercerNombre = ""
   oTabla.UsoWebReniec = False
   
'<(Inicio) Añadido Por: WABG el: 22/10/2020-10:31:42 p.m.en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
   oTabla.validacionReniec = False
'</(Fin) Añadido Por: WABG el: 22/10/2020-10:31:42 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
   
End Sub


Function FiltrarTodosSoloHistoriasDefinit_rap(ByVal oTabla As DoPaciente, lcSinApellido As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     Set FiltrarTodosSoloHistoriasDefinit_rap = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
       .CommandText = "PacientesFiltrarTodosSoloHistoriasDefinit_rap"
        Set oParameter = .CreateParameter("@NroHistoriaClinica", adInteger, adParamInput, 0, oTabla.NroHistoriaClinica): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoPaterno", adVarChar, adParamInput, 40, oTabla.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ApellidoMaterno", adVarChar, adParamInput, 40, oTabla.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 40, oTabla.PrimerNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 40, oTabla.SegundoNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idDocIdentidad", adInteger, adParamInput, 0, oTabla.IdDocIdentidad): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 12, oTabla.NroDocumento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     sWhere = ""
     If lcSinApellido = oTabla.ApellidoPaterno Then
        sWhere = sWhere & " apellidoPaterno='" & lcSinApellido & "' and "
     End If
     If lcSinApellido = oTabla.ApellidoMaterno Then
        sWhere = sWhere & " apellidoMaterno='" & lcSinApellido & "' and "
     End If
     If sWhere <> "" Then
        oRecordset.Filter = Left(sWhere, Len(sWhere) - 4)
     End If
     
     Set FiltrarTodosSoloHistoriasDefinit_rap = oRecordset
     ms_MensajeError = ""
     Set oRecordset = Nothing
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

