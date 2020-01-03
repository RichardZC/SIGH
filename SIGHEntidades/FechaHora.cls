VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FechaHora"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de tipo FechaHora
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Property Get HORA_VACIA_HM() As String
    HORA_VACIA_HM = "__:__"
End Property
Property Get FECHA_VACIA_DMY() As String
    FECHA_VACIA_DMY = "__/__/____"
End Property
Property Get FECHA_VACIA_DMY_HM() As String
    FECHA_VACIA_DMY_HM = "__/__/____ __:__"
End Property

Property Get FECHA_VACIA_DMY_HMS() As String
    FECHA_VACIA_DMY_HMS = "__/__/____ __:__:__"
End Property

Function DMA2AMD(ByVal sFecha As String) As String
        'DD/MM/AAAA TO AAAA/MM/DD
        DMA2AMD = Right(sFecha, 4) + "/" + Mid(sFecha, 4, 2) + "/" + Left(sFecha, 2)
End Function
Function AMD2DMA(ByVal sFecha As String) As String
        'AAAA/MM/DD to DD/MM/AAAA
        AMD2DMA = Right(sFecha, 2) + "/" + Mid(sFecha, 6, 2) + "/" + Left(sFecha, 4)
End Function
Function ValidaHora(ByVal sHora As String) As Boolean
    ValidaHora = False
    If Right(sHora, 2) = "__" Or Left(sHora, 2) = "__" Then
       Exit Function
    End If
    
    If Val(Right(sHora, 2)) > 59 Then
        Exit Function
    End If
    
    If Val(Left(sHora, 2)) > 23 Then
        Exit Function
    End If
    
    ValidaHora = True
End Function

Function EsHora(sHora As String) As Boolean   'Formato: "hh:mm"
    Dim lnHora As Integer
    Dim lnMinuto As Integer
    Dim lnCorr As Integer
    EsHora = True
    If IsNull(sHora) Then
       EsHora = False
       Exit Function
    End If
    For lnCorr = 1 To 5
        If InStr("1234567890:", Mid(sHora, lnCorr, 1)) = 0 Then
            EsHora = False
            Exit Function
        End If
    Next
    lnHora = Val(Left(sHora, 2))
    lnMinuto = Val(Right(sHora, 2))
    If Not (lnHora >= 0 And lnHora <= 23) Then
       EsHora = False
       Exit Function
    End If
    If Not (lnMinuto >= 0 And lnMinuto <= 59) Then
       EsHora = False
       Exit Function
    End If
End Function


Function EsFecha(sFecha As String, sFormato As String, Optional lbConsideraHorasMinutos As Boolean) As Boolean
Dim iAno As Integer
Dim iMes As Integer
Dim iDia As Integer
Dim iIniDia As Integer
Dim iIniMes As Integer
Dim iIniAno As Integer
Dim bBisiesto  As Boolean
On Error GoTo EsFecha_err
    
    'formato es el formato frente al que se va a comparar la fecha
    EsFecha = False
    If lbConsideraHorasMinutos = False Then
       If Len(sFecha) <> 10 Then Exit Function
    End If
    If Not IsDate(sFecha) Then Exit Function
    
    iIniDia = InStr(sFormato, "D")
    iIniMes = InStr(sFormato, "M")
    iIniAno = InStr(sFormato, "AAAA")

    iDia = CInt(Mid(sFecha, iIniDia, 2))
    iMes = CInt(Mid(sFecha, iIniMes, 2))
    iAno = CInt(Mid(sFecha, iIniAno, 4))

    If iAno < 1800 Or iAno > 2100 Then
        Exit Function
    End If
    
    If iMes < 1 Or iMes > 12 Then
        Exit Function
    End If
    
    If iAno Mod 4 = 0 Then
        If iAno Mod 100 = 0 Then
            If iAno Mod 400 = 0 Then
                bBisiesto = True
            Else
                bBisiesto = False
            End If
        Else
            bBisiesto = True
        End If
    Else
        bBisiesto = False
    End If
    
    Select Case iMes
    Case 1, 3, 5, 7, 8, 10, 12
        If iDia < 1 Or iDia > 31 Then
            Exit Function
        End If
    Case 4, 6, 9, 11
        If iDia < 1 Or iDia > 30 Then
            Exit Function
        End If
    Case 2
        If bBisiesto Then
            If iDia < 1 Or iDia > 29 Then
                Exit Function
            End If
        Else
            If iDia < 1 Or iDia > 28 Then
                Exit Function
            End If
        End If
    End Select
    
    EsFecha = True
    
Exit Function
EsFecha_err:
    Exit Function
End Function
Function DiasDelMes(iAno As Integer, iMes As Integer) As Integer
Dim bBisiesto As Boolean

    If iAno Mod 4 = 0 Then
        If iAno Mod 100 = 0 Then
            If iAno Mod 400 = 0 Then
                bBisiesto = True
            Else
                bBisiesto = False
            End If
        Else
            bBisiesto = True
        End If
    Else
        bBisiesto = False
    End If
    
    Select Case iMes
    Case 1, 3, 5, 7, 8, 10, 12
        DiasDelMes = 31
    Case 4, 6, 9, 11
        DiasDelMes = 30
    Case 2
        If bBisiesto Then
            DiasDelMes = 29
        Else
            DiasDelMes = 28
        End If
    End Select

End Function

Public Function EdadDevuelveTipo(lnIdTipoEdad As Long) As String
    Select Case lnIdTipoEdad
    Case 1
         EdadDevuelveTipo = "A"
    Case 2
         EdadDevuelveTipo = "M"
    Case 3
         EdadDevuelveTipo = "D"
    Case Else
         EdadDevuelveTipo = "H"
    End Select
End Function
Public Function EdadActual(dFechaNac As Date, dDate As Date) As Integer
'OBJETIVO: Calcula la Edad
'REQUIERE: dFechaNac =Fecha de Nacimiento
'                   dDate= Fecha Actual
         Dim iUno As Integer
         iUno = 0
         If Month(dDate) < Month(dFechaNac) Then
            iUno = 1
         Else
            If Month(dDate) = Month(dFechaNac) Then
               If Not (day(dDate) >= day(dFechaNac)) Then
                  iUno = 1
               End If
            End If
         End If
         EdadActual = Year(dDate) - Year(dFechaNac) - iUno
End Function
Public Function EdadActualEnDias(dFechaNac As Date, dDate As Date) As Long
    EdadActualEnDias = DateDiff("d", dFechaNac, dDate)
End Function

Public Function EdadActualEnFormatoYYYYMMDD(dFechaNac As Date, dDate As Date) As String
    Dim lcMDA As String
    lcMDA = Format(DateDiff("d", dFechaNac, dDate) + 1, "yyyymmdd")
    EdadActualEnFormatoYYYYMMDD = Right("0000" & Trim(str(Val(Left(lcMDA, 4)) - 1900)), 4) & Mid(lcMDA, 5, 4)
End Function


Function PrimerFechaDDMMYYDelMesActual() As String
    PrimerFechaDDMMYYDelMesActual = "01" & "/" + Right("00" & Month(Date), 2) & "/" & Year(Date)
End Function
Function UltimaFechaDDMMYYDelMesActual() As String
Dim PrimeDiaSiguienteMes As Date
    PrimeDiaSiguienteMes = CDate("01" & "/" & Month(DateAdd("m", 1, Date)) & "/" & Year(DateAdd("m", 1, Date)))
    UltimaFechaDDMMYYDelMesActual = CStr(PrimeDiaSiguienteMes - 1)
End Function
Function PrimerFechaDDMMYYDelMesAnterior() As String
    PrimerFechaDDMMYYDelMesAnterior = CDate("01" & "/" & Month(DateAdd("m", -1, Date)) & "/" & Year(DateAdd("m", -1, Date)))
End Function
Function UltimaFechaDDMMYYDelMesAnterior() As String
Dim PrimeDiaSiguienteMes As Date
    PrimeDiaSiguienteMes = CDate("01" & "/" & Month(Date) & "/" & Year(Date))
    UltimaFechaDDMMYYDelMesAnterior = CStr(PrimeDiaSiguienteMes - 1)
End Function

'***************daniel barrantes**************
'***************Devuelve el Nombre del Mes
'***************
Function DevuelveNombreMes(lnNroMes As Integer) As String
    Select Case lnNroMes
    Case 1
        DevuelveNombreMes = "Enero"
    Case 2
        DevuelveNombreMes = "Febrero"
    Case 3
        DevuelveNombreMes = "Marzo"
    Case 4
        DevuelveNombreMes = "Abril"
    Case 5
        DevuelveNombreMes = "Mayo"
    Case 6
        DevuelveNombreMes = "Junio"
    Case 7
        DevuelveNombreMes = "Julio"
    Case 8
        DevuelveNombreMes = "Agosto"
    Case 9
        DevuelveNombreMes = "Setiembre"
    Case 10
        DevuelveNombreMes = "Octubre"
    Case 11
        DevuelveNombreMes = "Noviembre"
    Case 12
        DevuelveNombreMes = "Diciembre"
    End Select
End Function


Function DevuelveEdadEnMeses(ldFechaNacimiento As Date, ldFechaActual As Date) As Long
    DevuelveEdadEnMeses = DateDiff("m", ldFechaNacimiento, ldFechaActual)
End Function

Function DevuelveEdadEnSemanas(ldFechaNacimiento As Date, ldFechaActual As Date) As Long
    DevuelveEdadEnSemanas = DateDiff("w", ldFechaNacimiento, ldFechaActual)
End Function

Function DevuelveEdadGestacional(ldFechaFUM As Date, ldFechaActual As Date) As Integer
    Dim lnDiaFUM As Integer: Dim lnMesFUM As Integer: Dim lnAnioFUM As Integer
    Dim lnDia As Integer: Dim lnMes As Integer: Dim lnAnio As Integer
    Dim lnDiasRestantesMesFUM As Long: Dim lnDiasRangoMeses As Long
    Const lnDiasMesG As Integer = 28
    Const lnSemanasMesG As Integer = 7
    lnDia = day(ldFechaActual)
    lnMes = Month(ldFechaActual)
    lnAnio = Year(ldFechaActual)
    lnDiaFUM = day(ldFechaFUM)
    lnMesFUM = Month(ldFechaFUM)
    lnAnioFUM = Year(ldFechaFUM)
    If lnAnioFUM = lnAnio And lnMesFUM = lnMes Then
        DevuelveEdadGestacional = Round((lnDia - lnDiaFUM) / lnSemanasMesG, 0)
    Else
        If lnAnioFUM = lnAnio Then
            lnDiasRestantesMesFUM = lnDiasMesG - lnDiaFUM
            lnDiasRangoMeses = ((lnMes - lnMesFUM) * lnDiasMesG)
            lnDiasRangoMeses = lnDiasRangoMeses - lnDiasMesG
            DevuelveEdadGestacional = Round((lnDiasRestantesMesFUM + lnDiasRangoMeses + lnDia) / lnSemanasMesG, 0)
        Else
            lnDiasRestantesMesFUM = lnDiasMesG - lnDiaFUM
            lnDiasRangoMeses = ((12 + lnMes - lnMesFUM) * lnDiasMesG)
            lnDiasRangoMeses = lnDiasRangoMeses - lnDiasMesG
            DevuelveEdadGestacional = Round((lnDiasRestantesMesFUM + lnDiasRangoMeses + lnDia) / lnSemanasMesG, 0)
        End If
    End If
    If DevuelveEdadGestacional >= 50 Then
       DevuelveEdadGestacional = 0
    End If
End Function

Function DevuelveUltimoDiaDelMes(lnMes As Integer, lnAnio As Integer) As Integer
    Dim lnResiduo1 As Double
   Dim lnResiduo2 As Double
   Select Case lnMes
   Case 1, 3, 5, 7, 8, 10, 12
      DevuelveUltimoDiaDelMes = 31
   Case 2
      lnResiduo1 = Round(lnAnio / 4, 0)
      lnResiduo2 = lnAnio / 4
      If lnResiduo1 = lnResiduo2 Then
            DevuelveUltimoDiaDelMes = 29
      Else
            DevuelveUltimoDiaDelMes = 28
      End If
   Case Else
       DevuelveUltimoDiaDelMes = 30
   End Select
End Function

'debb-10/12/2018
Function VerificaSiRangoEsDeUnMesCompleto(mda_FechaInicio As Date, mda_FechaFin As Date, lc_CodigoItem As String) As Boolean
    VerificaSiRangoEsDeUnMesCompleto = False
    If Val(lc_CodigoItem) > 0 Then
       Exit Function
    End If
'    If Format(mda_FechaInicio, "hh:mm:ss") <> "00:00:00" Then
'       Exit Function
'    End If
    If day(mda_FechaInicio) <> 1 Then
       Exit Function
    End If
'    If Format(mda_FechaFin, "hh:mm:ss") <> "23:59:59" Then
'       Exit Function
'    End If
    If day(mda_FechaFin) <> DevuelveUltimoDiaDelMes(Month(mda_FechaFin), Year(mda_FechaFin)) Then
       Exit Function
    End If
    If Month(mda_FechaInicio) <> Month(mda_FechaFin) Then
       Exit Function
    End If
    VerificaSiRangoEsDeUnMesCompleto = True
End Function

Function VerificaClaveMesDia(lcClave2 As String) As Boolean
    If Len(lcClave2) = 0 Then
        VerificaClaveMesDia = False
    Else
        If Val(Left(lcClave2, 2)) = Month(Date) And Val(Right(lcClave2, 2)) = day(Date) Then
           VerificaClaveMesDia = True
        Else
           VerificaClaveMesDia = False
        End If
    End If
End Function

Function VerificaSiRangoEsDeDosOmasMesesCompletos(mda_FechaInicio As Date, mda_FechaFin As Date) As Boolean
    VerificaSiRangoEsDeDosOmasMesesCompletos = False
    If Format(mda_FechaInicio, "hh:mm:ss") <> "00:00:00" Then
       Exit Function
    End If
    If day(mda_FechaInicio) <> 1 Then
       Exit Function
    End If
    If Format(mda_FechaFin, "hh:mm:ss") <> "23:59:59" Then
       Exit Function
    End If
    If day(mda_FechaFin) <> DevuelveUltimoDiaDelMes(Month(mda_FechaFin), Year(mda_FechaFin)) Then
       Exit Function
    End If
    If Month(mda_FechaInicio) = Month(mda_FechaFin) And Year(mda_FechaInicio) = Year(mda_FechaFin) Then
       Exit Function
    End If
    VerificaSiRangoEsDeDosOmasMesesCompletos = True
End Function


Public Function DevuelveFechaNacimiento(lcFechaActual As String, lcHoraActual As String, lnEdad As Integer, lnTipoEdad As sghTipoEdades) As Date
   Dim ldFechaActual As Date
   ldFechaActual = CDate(lcFechaActual & " " & lcHoraActual)
   Select Case lnTipoEdad
   Case sghAño
        DevuelveFechaNacimiento = DateAdd("yyyy", -lnEdad, ldFechaActual)
   Case sghMeses
        DevuelveFechaNacimiento = DateAdd("m", -lnEdad, ldFechaActual)
   Case sghDias
        DevuelveFechaNacimiento = DateAdd("d", -lnEdad, ldFechaActual)
   Case sghHoras
        DevuelveFechaNacimiento = DateAdd("h", -lnEdad, ldFechaActual)
   End Select
End Function



Function DevuelveCodigoDxSinPUNTO(lcDxConPunto As String) As String
    Dim lcDx As String, lnLen As Integer
    lcDx = ""
    For lnLen = 1 To Len(lcDxConPunto)
        If Mid(lcDxConPunto, lnLen, 1) <> "." Then
           lcDx = lcDx & Mid(lcDxConPunto, lnLen, 1)
        End If
    Next
    DevuelveCodigoDxSinPUNTO = lcDx
End Function




Function ConvierteEnAnioUnMesOdia(lnNumeroMD As Integer, IndexComboDMA As Integer) As Double
    Select Case IndexComboDMA
    Case 0  'Se elijio Dia
         ConvierteEnAnioUnMesOdia = Round(lnNumeroMD / 365, 5)
    Case 1  'Se elijio Mes
         ConvierteEnAnioUnMesOdia = Round(lnNumeroMD / 12, 5)
    Case 2  'Se elijio Año
         ConvierteEnAnioUnMesOdia = lnNumeroMD
    Case 4  'Horas
         ConvierteEnAnioUnMesOdia = Round(lnNumeroMD / (365 * 24), 5)
    End Select
End Function





Function CalcularEdad(daFechaNacimiento As Date, daFechaReferencia As Date) As Edad
Dim iEdad As Integer
Dim oEdad As Edad
    
    iEdad = DateDiff("yyyy", daFechaNacimiento, daFechaReferencia)
    
    If day(daFechaNacimiento) = 29 And Month(daFechaNacimiento) = 2 Then
        If CDate(28 & "/" & Month(daFechaNacimiento) & "/" & Year(daFechaReferencia)) > daFechaReferencia Then
            iEdad = iEdad - 1
        End If
    Else
        If CDate(day(daFechaNacimiento) & "/" & Month(daFechaNacimiento) & "/" & Year(daFechaReferencia)) > daFechaReferencia Then
            iEdad = iEdad - 1
        End If
    End If
    
    If iEdad > 0 Then
        oEdad.Edad = iEdad
        oEdad.TipoEdad = 1
        oEdad.NombreEdad = "Años"
        CalcularEdad = oEdad
        Exit Function
    End If
    
    iEdad = DateDiff("d", daFechaNacimiento, daFechaReferencia)
    If iEdad > 30 Then
        iEdad = DateDiff("m", daFechaNacimiento, daFechaReferencia)
        'mgaray20141013
        If day(daFechaNacimiento) > day(daFechaReferencia) Then
            iEdad = iEdad - 1
        End If
        If iEdad = 12 Then
           oEdad.Edad = 11
           oEdad.TipoEdad = 2
           oEdad.NombreEdad = "Meses"
           CalcularEdad = oEdad
           Exit Function
        ElseIf iEdad > 0 Then
            oEdad.Edad = iEdad
            oEdad.TipoEdad = 2
            oEdad.NombreEdad = "Meses"
            CalcularEdad = oEdad
            
            Exit Function
        End If
    Else
        If iEdad = 0 Then
            iEdad = DateDiff("h", daFechaNacimiento, daFechaReferencia)
            oEdad.Edad = IIf(iEdad = 0, 1, iEdad)    'debb-07/02/2017
            oEdad.TipoEdad = 4
            oEdad.NombreEdad = "Horas"
            CalcularEdad = oEdad
            Exit Function
        Else
            oEdad.Edad = iEdad
            oEdad.TipoEdad = 3
            oEdad.NombreEdad = "Días"
            CalcularEdad = oEdad
            Exit Function
        End If
    End If
End Function

Function CalculaSiEsRecienNacido(daFechaNacimiento As Date, daFechaReferencia As Date) As Integer
    Dim iEdad As Long
    iEdad = DateDiff("d", daFechaNacimiento, daFechaReferencia)
    If iEdad <= 28 Then
       CalculaSiEsRecienNacido = 1     'Es un Recien nacido
    Else
       CalculaSiEsRecienNacido = 0     'No es recien nacido
    End If
End Function

 Function EdadEnDias(oEdad As Edad) As Integer
    On Error Resume Next
    Select Case oEdad.TipoEdad
    Case 1
        EdadEnDias = oEdad.Edad * 365
    Case 2
        EdadEnDias = oEdad.Edad * 30
    Case 3
        EdadEnDias = oEdad.Edad
    Case 4
        EdadEnDias = 1
    End Select
End Function

Function DiaDeLASemana(daFecha As Date)
Dim day As Integer

    day = DatePart("w", daFecha)
    Select Case day
        Case 1
            DiaDeLASemana = "Domingo"
        Case 2
            DiaDeLASemana = "Lunes"
        Case 3
            DiaDeLASemana = "Martes"
        Case 4
            DiaDeLASemana = "Miercoles"
        Case 5
            DiaDeLASemana = "Jueves"
        Case 6
            DiaDeLASemana = "Viernes"
        Case 7
            DiaDeLASemana = "Sábado"
    End Select
End Function


Function DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaMovimiento As Date) As Date
    Dim lcUltDiaMes As String, ldFechaMov As Date
    lcUltDiaMes = Trim(str(DevuelveUltimoDiaDelMes(Month(ldFechaMovimiento), Year(ldFechaMovimiento))))
    DevuelveFechaHoraFinalDelMesDelMovimiento = CDate(lcUltDiaMes & Format(ldFechaMovimiento, "/mm/yyyy") & " 23:59:59")
End Function

Function DevuelveUltimaFechaDelMesAnteriorDelMovimiento(ldFechaMovimiento As Date) As Date
    ldFechaMovimiento = DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaMovimiento)
    DevuelveUltimaFechaDelMesAnteriorDelMovimiento = CDate("01" & Format(ldFechaMovimiento, "/mm/yyyy hh:mm:ss")) - 1
End Function

Property Get DevuelveHoraSoloFormato_HM() As String
    DevuelveHoraSoloFormato_HM = "hh:mm"
End Property
Property Get DevuelveHoraSoloFormato_HMS() As String
    DevuelveHoraSoloFormato_HMS = "hh:mm:ss"
End Property
Property Get DevuelveFechaSoloFormato_DMY() As String
    DevuelveFechaSoloFormato_DMY = "dd/mm/yyyy"
End Property

Property Get DevuelveFechaSoloFormato_DMY_HM() As String
    DevuelveFechaSoloFormato_DMY_HM = "dd/mm/yyyy hh:mm"
End Property

Property Get DevuelveFechaSoloFormato_DMY_HMS() As String
    DevuelveFechaSoloFormato_DMY_HMS = "dd/mm/yyyy hh:mm:ss"
End Property
Property Get DevuelveFechaSoloFormato_DMYHMS() As String
    DevuelveFechaSoloFormato_DMYHMS = "ddmmyyyyhhmmss"
End Property
Property Get HORA_VACIA_HMS() As String
    HORA_VACIA_HMS = "__:__:__"
End Property

Function EsBisiesto(Anio As Integer) As Boolean
  If Anio Mod 4 = 0 Then
    If Anio Mod 100 = 0 Then
      If Anio Mod 400 = 0 Then
        EsBisiesto = True
      Else
        EsBisiesto = False
      End If
    Else
      EsBisiesto = True
    End If
  Else
    EsBisiesto = False
  End If
End Function

Property Get DevuelveFechaSoloFormato_YMD_SIS() As String
    DevuelveFechaSoloFormato_YMD_SIS = "yyyymmdd"
End Property

'Property Get DevuelveFechaSegunFormato_YMD_SIS(lcFechaSIS As String) As String
'    If IsNull(lcFechaSIS) Or lcFechaSIS = "" Or lcFechaSIS = FECHA_VACIA_DMY Then
'       DevuelveFechaSegunFormato_YMD_SIS = FECHA_VACIA_DMY
'    Else
'       DevuelveFechaSegunFormato_YMD_SIS = CDate(Right(lcFechaSIS, 2) & "/" & Mid(lcFechaSIS, 5, 2) & "/" & Left(lcFechaSIS, 4))
'    End If
'End Property

Property Get DevuelveFechaSegunFormato_YMD_SIS(lcFechaSIS As String) As String
    If IsNull(lcFechaSIS) Or lcFechaSIS = "" Or lcFechaSIS = FECHA_VACIA_DMY Then
       DevuelveFechaSegunFormato_YMD_SIS = FECHA_VACIA_DMY
    Else
        If IsDate(lcFechaSIS) Then 'Actualizado 16092014
            DevuelveFechaSegunFormato_YMD_SIS = CDate(lcFechaSIS)
        Else
            DevuelveFechaSegunFormato_YMD_SIS = CDate(Right(lcFechaSIS, 2) & "/" & Mid(lcFechaSIS, 5, 2) & "/" & Left(lcFechaSIS, 4))
        End If
    End If
End Property

Function DevuelveRangoEdadesDiaMesAnio(lnMinEdad As String, lnMaxEdad As String) As String
    Dim lcMensaje1 As String
    lcMensaje1 = "La Edad es entre: " & lnMinEdad & " y " & lnMaxEdad & " días"
    lcMensaje1 = lcMensaje1 & " (en meses: " & IIf(Val(lnMinEdad) = 0, "0", Trim(str(Round(Val(lnMinEdad) / 30, 0)))) & " , " & Trim(str(Round(Val(lnMaxEdad) / 30, 0))) & ")"
    lcMensaje1 = lcMensaje1 & " (en años: " & IIf(Val(lnMinEdad) = 0, "0", Trim(str(Round(Val(lnMinEdad) / 365, 0)))) & " , " & Trim(str(Round(Val(lnMaxEdad) / 365, 0))) & ")"
    DevuelveRangoEdadesDiaMesAnio = lcMensaje1
End Function

Function PresionDevuelveVacia() As String
     PresionDevuelveVacia = "___/___"
End Function

Sub PresionDevuelveSistolicaDiastolica(lcPresionSistolicaDiastolica As String, ByRef lcSistolica As String, ByRef lcDiastolica As String)
    Dim lnFor As Integer, lbPasoSlash As Boolean
    lbPasoSlash = False
    lcSistolica = ""
    For lnFor = 1 To Len(lcPresionSistolicaDiastolica)
        Select Case Mid(lcPresionSistolicaDiastolica, lnFor, 1)
        Case "/"
            lbPasoSlash = True
        Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
            If lbPasoSlash = True Then
               lcDiastolica = lcDiastolica & Mid(lcPresionSistolicaDiastolica, lnFor, 1)
            Else
               lcSistolica = lcSistolica & Mid(lcPresionSistolicaDiastolica, lnFor, 1)
            End If
        End Select
    Next
End Sub
Function Quita_DePresion(lcTxtSPpa As String) As String
    Dim lcDx As String, lnLen As Integer
    lcDx = ""
    For lnLen = 1 To Len(lcTxtSPpa)
        If Mid(lcTxtSPpa, lnLen, 1) <> "_" Then
           lcDx = lcDx & Mid(lcTxtSPpa, lnLen, 1)
        End If
    Next
    Quita_DePresion = lcDx
      
End Function

Function PresionVerificaSiTieneDatosYsiEstaOK(lcPresionSistolicaDiastolica As String) As Boolean
        PresionVerificaSiTieneDatosYsiEstaOK = False
        Dim lcSistolica As String, lcDiastolica As String
        If lcPresionSistolicaDiastolica <> PresionDevuelveVacia Then
            lcSistolica = Replace(Left(lcPresionSistolicaDiastolica, InStr(lcPresionSistolicaDiastolica, "/") - 1), "_", "")
            lcDiastolica = Replace(Trim(Mid(lcPresionSistolicaDiastolica, InStr(lcPresionSistolicaDiastolica, "/") + 1, 10)), "_", "")
            If Val(lcSistolica) > 0 And Val(lcDiastolica) > 0 Then
               If Val(lcSistolica) <= Val(lcDiastolica) Then
                  MsgBox "Si registra la PRESION: SISTOLICA/DIASTOLICA, el valor SISTOLICA debe ser mayor a la DIASTOLICA", vbInformation, "mensaje"
               Else
                  PresionVerificaSiTieneDatosYsiEstaOK = True
               End If
            Else
               MsgBox "Si registra la PRESION, debe ser en el siguiente formato SISTOLICA/DIASTOLICA, recuerde que el valor SISTOLICA debe ser mayor a la DIASTOLICA", vbCritical, "mensaje"
            End If
        End If
End Function

Function PresionJuntaSistolicaDiastolica(lcSistolica As String, lcDiastolica As String) As String
    PresionJuntaSistolicaDiastolica = Right("___" & lcSistolica, 3) & "/" & Left(lcDiastolica & "___", 3)
End Function

'JVG - Calculo de Numero de dias trnascurriods en un Año, pasandole como parametro la fecha
Function DevuelveDiasTranscurridos(ldfechaInicial As Date, ldfechaFinal As Date) As Long
    DevuelveDiasTranscurridos = DateDiff("d", ldfechaInicial, ldfechaFinal)
End Function

Property Get DevuelveFechaSoloFormato_AAMMDD() As String
    DevuelveFechaSoloFormato_AAMMDD = "yymmdd"
End Property
'mgaray
Public Function calcularEdadDisgregada(daFechaNacimiento As Date, _
                    daFechaReferencia As Date) As Edad
    'calcular años transcurridod
    Dim oEdad As Edad
    Dim anios As Integer, meses As Integer, dias As Integer
    Dim fechaTemp As Date
    
    anios = EdadActual(daFechaNacimiento, daFechaReferencia)
    
    fechaTemp = DateAdd("yyyy", anios, daFechaNacimiento)
    If fechaTemp <> daFechaReferencia Then
        meses = DateDiff("m", fechaTemp, daFechaReferencia)
        If day(fechaTemp) > day(daFechaReferencia) Then
            meses = meses - 1
        End If
        
    End If
    fechaTemp = DateAdd("m", meses, fechaTemp)
    If fechaTemp <> daFechaReferencia Then
        dias = DateDiff("d", fechaTemp, daFechaReferencia)
    End If
    oEdad.EdadAnio = anios
    oEdad.EdadMes = meses
    oEdad.EdadDia = dias
    
    calcularEdadDisgregada = oEdad
End Function

Property Get DevuelveFechaSoloFormato_AAAAMMDD() As String
    DevuelveFechaSoloFormato_AAAAMMDD = "yyyymmdd"
End Property


Function UltimaFechaDDMMYYDelMesActual1(ldHoy As Date) As String
Dim PrimeDiaSiguienteMes As Date
    PrimeDiaSiguienteMes = CDate("01" & "/" & Month(DateAdd("m", 1, ldHoy)) & "/" & Year(DateAdd("m", 1, ldHoy)))
    UltimaFechaDDMMYYDelMesActual1 = CStr(PrimeDiaSiguienteMes - 1)
End Function

Function DevuelveFUM(lnEG As Integer, ldFechaAtencion As Date) As String
    Dim ldFecha1 As Date
    ldFecha1 = ldFechaAtencion - (lnEG * 7)
    DevuelveFUM = Format(ldFecha1, DevuelveFechaSoloFormato_DMY)
End Function



Function CalcularAnioMesDia(FechaInicio As Date, FechaActual As Date, DevuelveAnioMesOdia As sghTipoEdades) As Long
Dim diaActual As Integer, mesActual As Integer, anioActual As Integer
Dim diaInicio As Integer, mesInicio As Integer, anioInicio As Integer
Dim b As Integer, mes As Integer, dias As Integer, anios As Integer, meses As Integer

diaActual = DatePart("d", FechaActual)
mesActual = DatePart("m", FechaActual)
anioActual = DatePart("yyyy", FechaActual)
'**************************************'
diaInicio = DatePart("d", FechaInicio)
mesInicio = DatePart("m", FechaInicio)
anioInicio = DatePart("yyyy", FechaInicio)
 
b = 0
mes = mesInicio - 1
 
' si el mes es febrero
If (mes = 2) Then   ' *
    If ((anioActual / 4 = 0 And anioActual / 100! = 0) Or anioActual / 400 = 0) Then
        b = 29
    Else
        b = 28
    End If
ElseIf (mes <= 7) Then  '*
    If (mes = 0) Then
        b = 31
    ElseIf (mes / 2 = 0) Then
        b = 30
    Else
        b = 31
    End If
    
ElseIf (mes > 7) Then
    If (mes / 2 = 0) Then
        b = 31
    Else
        b = 30
    End If
End If
    
   If ((anioInicio > anioActual) Or (anioInicio = anioActual And mesInicio > mesActual) Or (anioInicio = anioActual And mesInicio = mesActual And diaInicio > diaActual)) Then
        MsgBox "La fecha de inicio ha de ser anterior a la fecha Actual"
   Else
        If (mesInicio <= mesActual) Then
            anios = anioActual - anioInicio
            If (diaInicio <= diaActual) Then
                meses = mesActual - mesInicio
                dias = diaActual - diaInicio
            Else
                If (mesActual = mesInicio) Then
                    anios = anios - 1
                End If
                meses = (mesActual - mesInicio - 1 + 12) / 12
                dias = b - (diaInicio - diaActual)
            End If
        Else
            anios = anioActual - anioInicio - 1
  
        If (diaInicio > diaActual) Then
            meses = mesActual - mesInicio - 1 + 12
            dias = b - (diaInicio - diaActual)
        Else
            meses = mesActual - mesInicio + 12
            dias = diaActual - diaInicio
        End If
  End If
 
End If '*
 
If DevuelveAnioMesOdia = sghAño Then
   CalcularAnioMesDia = anios
ElseIf DevuelveAnioMesOdia = sghMeses Then
   CalcularAnioMesDia = meses
ElseIf DevuelveAnioMesOdia = sghDias Then
   CalcularAnioMesDia = dias
Else
End If

End Function



Function getDescripcionEdad(ml_lnEdadEnDias As Long, ml_lnIdTipoEdad As Long, _
                md_fechaNacimiento As Date, md_fechaIngreso As Date) As String
    Dim ls_edad As String
    ls_edad = Trim(str(ml_lnEdadEnDias)) & " " & EdadDevuelveTipo(ml_lnIdTipoEdad)
    If ml_lnIdTipoEdad = 1 Then
        Dim oEdad As Edad
        oEdad = calcularEdadDisgregada(md_fechaNacimiento, md_fechaIngreso)
        If oEdad.EdadMes > 0 Then
            ls_edad = ls_edad & ", " & oEdad.EdadMes & " " & EdadDevuelveTipo(2)
        End If
        
    End If
    getDescripcionEdad = ls_edad
End Function


