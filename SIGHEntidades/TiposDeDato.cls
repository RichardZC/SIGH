VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiposDeDato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Tipos de Datos
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
'GLCC 02/11/20 CAMBIO36 INICIO
'ANTES: Const wxNueve As String = "9"
'GLCC 02/11/20 CAMBIO36 FIN
Dim Unidades$(9), Decenas$(9), Oncenas$(9)
Dim Veintes$(9), Centenas$(9)

Public Type sghBusquedaPrestamoHistorias
    IdTipoServicio As Long
    idServicio  As Long
    FechaPrestamoRequeridaDesde As Date
    FechaPrestamoRequeridaHasta As Date
    ComparadorFecha As String
    IdEstadoPrestamo As Long
    IdEnvio As Long
End Type

Public Type Edad
    Edad As Integer
    TipoEdad As Integer
    NombreEdad As String
    'mgaray
    EdadAnio As Integer
    EdadMes As Integer
    EdadDia As Integer
End Type

Public Type EpisodioClinico
    idEpisodio As Long
    lbNuevoEpisodio As Boolean
    lbCierreEpisodio As Boolean
End Type


Function Descifrar$(numero%)
Static SAL$(4)
Dim I%, CT As Double, DC As Double, DU As Double, UD  As Double
Dim VARIABLE$

For I% = 1 To 4: SAL$(I%) = " ": Next I%
VARIABLE$ = String$(3 - Len(Trim$(str$(numero%))), "0") + Trim$(str$(numero%))
CT = Val(Mid$(VARIABLE$, 1, 1)): '*** CENTENA
DC = Val(Mid$(VARIABLE$, 2, 1)): '*** DECENA
DU = Val(Mid$(VARIABLE$, 2, 2)): '*** DECENA + UNIDAD
UD = Val(Mid$(VARIABLE$, 3, 1)): '*** UNIDAD
If numero% = 100 Then
        SAL$(1) = "CIEN "
Else
        If CT <> 0 Then SAL$(1) = Centenas$(CT)
        If DC <> 0 Then
                If DU <> 10 And DU <> 20 Then
                        If DC = 1 Then SAL$(2) = Oncenas$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                        If DC = 2 Then SAL$(2) = Veintes$(UD): Descifrar$ = Trim$(SAL$(1) + " " + SAL$(2)): Exit Function
                End If
                SAL$(2) = " " + Decenas$(DC)
                If UD <> 0 Then SAL$(3) = "Y "
        End If
        If UD <> 0 Then SAL$(4) = Unidades$(UD)
End If
Descifrar = Trim$(SAL$(1) + SAL$(2) + SAL$(3) + SAL$(4))
End Function


'***************daniel barrantes**************
'***************Devuelve en LETRAS un MONTO
'***************usado al Imprimir Boleta
Function Numlet$(NUM#)
Dim DEC$, MILM$, MILL$, MILE$, UNID$
ReDim SALI$(11)
Dim var$, I%, AUX$
'NUM# = Round(NUM#, 2)
var$ = Trim$(str$(NUM#))
If InStr(var$, ".") = 0 Then
        var$ = var$ + ".00"
End If
If InStr(var$, ".") = Len(var$) - 1 Then
        var$ = var$ + "0"
End If
var$ = String$(15 - Len(LTrim$(var$)), "0") + LTrim$(var$)
DEC$ = Mid$(var$, 14, 2)
MILM$ = Mid$(var$, 1, 3)
MILL$ = Mid$(var$, 4, 3)
MILE$ = Mid$(var$, 7, 3)
UNID$ = Mid$(var$, 10, 3)
For I% = 1 To 11: SALI$(I%) = " ": Next I%
I% = 0
Unidades$(1) = "UNA    "
Unidades$(2) = "DOS    "
Unidades$(3) = "TRES   "
Unidades$(4) = "CUATRO "
Unidades$(5) = "CINCO  "
Unidades$(6) = "SEIS   "
Unidades$(7) = "SIETE  "
Unidades$(8) = "OCHO   "
Unidades$(9) = "NUEVE  "

Decenas$(1) = "DIEZ      "
Decenas$(2) = "VEINTE    "
Decenas$(3) = "TREINTA "
Decenas$(4) = "CUARENTA "
Decenas$(5) = "CINCUENTA "
Decenas$(6) = "SESENTA "
Decenas$(7) = "SETENTA "
Decenas$(8) = "OCHENTA "
Decenas$(9) = "NOVENTA "

Oncenas$(1) = "ONCE       "
Oncenas$(2) = "DOCE       "
Oncenas$(3) = "TRECE      "
Oncenas$(4) = "CATORCE    "
Oncenas$(5) = "QUINCE     "
Oncenas$(6) = "DIECISEIS  "
Oncenas$(7) = "DIECISIETE "
Oncenas$(8) = "DIECIOCHO  "
Oncenas$(9) = "DIECINUEVE "

Veintes$(1) = "VEINTIUNA    "
Veintes$(2) = "VEINTIDOS    "
Veintes$(3) = "VEINTITRES   "
Veintes$(4) = "VEINTICUATRO "
Veintes$(5) = "VEINTICINCO  "
Veintes$(6) = "VEINTISEIS   "
Veintes$(7) = "VEINTISIETE  "
Veintes$(8) = "VEINTIOCHO   "
Veintes$(9) = "VEINTINUEVE  "

Centenas$(1) = "       CIENTO "
Centenas$(2) = "   DOSCIENTOS "
Centenas$(3) = "  TRESCIENTOS "
Centenas$(4) = "CUATROCIENTOS "
Centenas$(5) = "   QUINIENTOS "
Centenas$(6) = "  SEISCIENTOS "
Centenas$(7) = "  SETECIENTOS "
Centenas$(8) = "  OCHOCIENTOS "
Centenas$(9) = "  NOVECIENTOS "

If NUM# > 999999999999.99 Then Numlet$ = " ": Exit Function
If Val(MILM$) >= 1 Then
        SALI$(2) = " MIL ": '** MILES DE MILLONES
        SALI$(4) = " MILLONES "
        If Val(MILM$) <> 1 Then
                Unidades$(1) = "UN     "
                Veintes$(1) = "VEINTIUN     "
                SALI$(1) = Descifrar$(Val(MILM$))
        End If
End If
If Val(MILL$) >= 1 Then
        If Val(MILL$) < 2 Then
                SALI$(3) = "UN ": '*** UN MILLON
                If Trim$(SALI$(4)) <> "MILLONES" Then
                        SALI$(4) = " MILLON "
                End If
        Else
                SALI$(4) = " MILLONES ": '*** VARIOS MILLONES
                Unidades$(1) = "UN     "
                Veintes$(1) = "VEINTIUN     "
                SALI$(3) = Descifrar$(Val(MILL$))
        End If
End If
'For I% = 2 To 9
'        Centenas$(I%) = Mid$(Centenas(I%), 1, 11) + "AS"
'Next I%
If Val(MILE$) > 0 Then
        SALI$(6) = " MIL ": '*** MILES
        If Val(MILE$) <> 1 Then
                SALI$(5) = Descifrar$(Val(MILE$))
        End If
End If
Unidades$(1) = "UNO    "
Veintes$(1) = "VEINTIUN"
If Val(UNID$) >= 1 Then
        SALI$(7) = Descifrar$(Val(UNID$)):  '*** CIENTOS
        If Val(DEC$) >= 10 Then
            SALI$(8) = " CON ": '*** DECIMALES
            SALI$(10) = Descifrar$(Val(DEC$))
        End If
End If
If Val(MILM$) = 0 And Val(MILL$) = 0 And Val(MILE$) = 0 And Val(UNID$) = 0 Then SALI$(7) = " CERO "
AUX$ = ""
For I% = 1 To 11
        AUX$ = AUX$ + SALI$(I%)
Next I%
Numlet$ = Trim$(AUX$)
End Function


'***************daniel barrantes**************
'***************Devuelve un numero entero (sin decimales)
'***************
Function DevuelveNumeroSinDecimales(lnNumeroDouble As Double) As Long
   Dim lcNumero As String: Dim lnPosDec As Integer: Dim lcDecimal As String
   lcDecimal = "."       'En este sistema se usa COMA
   lcNumero = Trim(str(lnNumeroDouble))
   lnPosDec = InStr(1, lcNumero, lcDecimal)
   If lnPosDec > 0 Then
      DevuelveNumeroSinDecimales = Val(Left(lcNumero, lnPosDec - 1))
   Else
      DevuelveNumeroSinDecimales = Val(lcNumero)
   End If
End Function

'***************daniel barrantes**************
'***************Devuelve solo los Decimales de un Importe
'***************
Function DevuelveSoloDecimales(lnNumeroDouble As Double) As String
   Dim lcNumero As String: Dim lnPosDec As Integer: Dim lcDecimal As String
   lcDecimal = "."       'En este sistema se usa COMA
   lcNumero = Trim(str(lnNumeroDouble))
   lnPosDec = InStr(1, lcNumero, lcDecimal)
   If lnPosDec > 0 Then
   DevuelveSoloDecimales = Left(Mid(lcNumero, lnPosDec + 1, 2) & "0", 2)
   Else
   DevuelveSoloDecimales = "00"
   End If
End Function

''***************daniel barrantes**************
''***************Devuelve Dias de estancia de un Paciente en Hospitalizacion o Emergencia
''***************
Function DiasDeEstanciaEnHospitalizacionEmergencia(ldFechaIngreso As Date, lcHoraIngreso As String, ldFechaAlta As Date, lcHoraAlta As String) As Integer
        Dim lnDiasEstancia As Integer
        Dim lcFechaInicial1 As String, lcFechaFinal1 As String
        lcFechaInicial1 = ldFechaIngreso & " " & lcHoraIngreso
        lcFechaFinal1 = Format(ldFechaAlta, "dd/mm/yyyy") & " " & lcHoraAlta
        lnDiasEstancia = DateDiff("d", lcFechaInicial1, lcFechaFinal1)
        If lnDiasEstancia = 0 Then lnDiasEstancia = 1      'Estuvo un solo dia
        DiasDeEstanciaEnHospitalizacionEmergencia = lnDiasEstancia
End Function


'***************daniel barrantes**************
'***************Devuelve Horas de estancia de un Paciente en Hospitalizacion o Emergencia
'***************
Function HorasDelPacienteEnHospitalizacionEmergencia(ldFechaIngreso As Date, lcHoraIngreso As String, ldFechaAlta As Date, lcHoraAlta As String) As Integer
    HorasDelPacienteEnHospitalizacionEmergencia = DateDiff("h", ldFechaIngreso & " " & lcHoraIngreso, ldFechaAlta & " " & lcHoraAlta)
End Function

Function DevuelveRangoExcelAimprimir(lcRangoActual As String, lnNroFilasDelReporte As Long) As String
    DevuelveRangoExcelAimprimir = Left(lcRangoActual, InStrRev(lcRangoActual, "$")) & Trim(str(lnNroFilasDelReporte + 2))
End Function


Function ElijeSiEsEstrategicoDevuelveId(lnIdTipoSalidaBienInsumo As Long) As Long
      If lnIdTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghVentaEstrategico Then
        If MsgBox("El Tipo de Salida es: 'IntervSanitarias/Ventas'" & Chr(13) & Chr(13) & "lo Registrará como VENTAS ? ", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
           ElijeSiEsEstrategicoDevuelveId = sghTipoSalidaItemFarmacia.sghSoloVenta
        Else
           ElijeSiEsEstrategicoDevuelveId = sghTipoSalidaItemFarmacia.sghSoloEstrategico
        End If
      Else
        ElijeSiEsEstrategicoDevuelveId = lnIdTipoSalidaBienInsumo
      End If
End Function

Function ElijeSiEsEstrategicoDevuelveNombre(lnIdTipoSalidaBienInsumo As Long) As String
        Select Case lnIdTipoSalidaBienInsumo
        Case sghTipoSalidaItemFarmacia.sghSoloVenta
             ElijeSiEsEstrategicoDevuelveNombre = "Ventas"
        Case sghTipoSalidaItemFarmacia.sghSoloEstrategico
             ElijeSiEsEstrategicoDevuelveNombre = "Estratégico"
        Case sghTipoSalidaItemFarmacia.sghDonaciones
             ElijeSiEsEstrategicoDevuelveNombre = "Donación"
        End Select
End Function

Function TipoServicioDevuelveNombreCorto(lnIdTipoServicio As Long) As String
    Select Case lnIdTipoServicio
    Case sghConsultaExterna
         TipoServicioDevuelveNombreCorto = "(Cons.Ext)"
    Case sghHospitalizacion
         TipoServicioDevuelveNombreCorto = "(Hosp)"
    Case sghEmergenciaConsultorios
         TipoServicioDevuelveNombreCorto = "(Emerg)"
    Case Else
         TipoServicioDevuelveNombreCorto = "(ARFSIS)"
    End Select
End Function

'debb-09/08/2016
Function CopyRecordset(ByVal rs As ADODB.Recordset, lcNroHistoriaClinicaName As String) As ADODB.Recordset
    On Error Resume Next
    Dim oCopyRst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim lcMensaje As String

    Set oCopyRst = New ADODB.Recordset

    ' Recorremos la colección de campos para ir
    ' copiándolos con las mismas propiedades
    For Each fld In rs.Fields
        If UCase(lcNroHistoriaClinicaName) = UCase(fld.Name) Then
            oCopyRst.Fields.Append _
                fld.Name, _
                adVarChar, _
                10, _
                104
        Else
            oCopyRst.Fields.Append _
                fld.Name, _
                fld.Type, _
                fld.DefinedSize, _
                fld.Attributes
            If Err.Number = 3001 Then
                oCopyRst.Fields.Append _
                fld.Name, _
                fld.Type, _
                1, _
                8
            End If
            ' Tenemos en cuenta las propiedades de precisión y
            ' escala numérica para los tipos de datos numéricos
            ' que tengan establecidas dichas propiedades.
            If ((fld.Type = adNumeric) Or (fld.Type = adDecimal)) Then
                With oCopyRst
                    .Fields(.Fields.Count - 1).Precision = fld.Precision
                    .Fields(.Fields.Count - 1).NumericScale = fld.NumericScale
                End With
            End If
        End If
    Next
      
    ' Devolvemos el nuevo objeto Recordset
    oCopyRst.Open
    Set CopyRecordset = oCopyRst

End Function




Function HCigualDNI_DevuelveRsConHistoriaOCHOdigitos(rsRs As Recordset, NroHistoriaClinicaName As String) As Recordset
    Dim oRsTmpHC As New Recordset
    Dim lnNroCamposRs As Integer, lnCampo As Integer, lcNroHistoria As String
    Dim oEdad As Edad, lbContiuar9 As Boolean
    Dim oEntidades As New SIGHEntidades.FechaHora
    
    Set oRsTmpHC = CopyRecordset(rsRs, "")
    lnNroCamposRs = oRsTmpHC.Fields.Count
    If rsRs.RecordCount > 0 Then
        rsRs.MoveFirst
        Do While Not rsRs.EOF
            oRsTmpHC.AddNew
            For lnCampo = 0 To lnNroCamposRs - 1
                If UCase(rsRs.Fields(lnCampo).Name) = UCase(NroHistoriaClinicaName) Then
                   lcNroHistoria = Trim(str(rsRs.Fields(lnCampo).Value))
                  'GLCC 02/11/20 CAMBIO36 INICIO
                   'ANTERIOR: If Len(lcNroHistoria) = 9 And Left(lcNroHistoria, 1) = wxNueve Then
                   If Len(lcNroHistoria) = 8 Then
                  'GLCC 02/11/20 CAMBIO36 FIN
                      oRsTmpHC.Fields(lnCampo).Value = Val(Right(lcNroHistoria, 8))
                   Else
                      oRsTmpHC.Fields(lnCampo).Value = rsRs.Fields(lnCampo).Value
                   End If
                ElseIf UCase(rsRs.Fields(lnCampo).Name) = "EDADPACIENTE" Then
                   oEdad = oEntidades.CalcularEdad(rsRs!FechaNacimiento, Date)
                   oRsTmpHC.Fields(lnCampo).Value = oEdad.Edad & " " & oEdad.NombreEdad
                Else
                   oRsTmpHC.Fields(lnCampo).Value = rsRs.Fields(lnCampo).Value
                End If
            Next
            oRsTmpHC.Update
            rsRs.MoveNext
        Loop
    End If
    Set oEntidades = Nothing
    Set HCigualDNI_DevuelveRsConHistoriaOCHOdigitos = oRsTmpHC.Clone
End Function


Function DevuelveNroOrdenParaHemoglobina(lnIdAtencion As Long) As Long
    If lnIdAtencion < 1000000 Then
       DevuelveNroOrdenParaHemoglobina = Val("999" + Trim(str(lnIdAtencion)))
    Else
       DevuelveNroOrdenParaHemoglobina = Val("99" + Trim(str(lnIdAtencion)))
    End If
End Function

Function DevuelveRutaConSlashInvertida(lcRuta As String) As String
  If Right(Trim(lcRuta), 1) = "\" Then
     DevuelveRutaConSlashInvertida = Trim(lcRuta)
  Else
     DevuelveRutaConSlashInvertida = Trim(lcRuta) & "\"
  End If
End Function
