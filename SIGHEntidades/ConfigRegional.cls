VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigRegional"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Configuración Regional
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

Property Get LOCALE_USER_DEFAULT() As Long
    LOCALE_USER_DEFAULT = &H400  'presentar información del usuario
End Property
Property Get LOCALE_SSHORTDATE() As Long
    LOCALE_SSHORTDATE = &H1F   'formato de fecha corta
End Property
Property Get LOCALE_STHOUSAND() As Long
    LOCALE_STHOUSAND = &HF 'sepàrador de miles
End Property
Property Get LOCALE_SDECIMAL() As Long
    LOCALE_SDECIMAL = &HE 'separador decimal
End Property
Property Get LOCALE_STIMEFORMAT() As Long
    LOCALE_STIMEFORMAT = &H1003 'formato de horas
End Property
Property Get LOCALE_SMONDECIMALSEP() As Long
    LOCALE_SMONDECIMALSEP = &H16 'formato decimal de monedas
End Property
Property Get LOCALE_SMONTHOUSANDSEP() As Long
    LOCALE_SMONTHOUSANDSEP = &H17 'separador de miles en las monedas
End Property

Property Get FormatoFechaCorta() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("1F", 16)        'LOCALE_sFormat = &H1F 'formato de fecha corta
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    FormatoFechaCorta = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property
Property Let FormatoFechaCorta(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sValor
End Property

Property Get SeparadorDecimal() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("E", 16)
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    SeparadorDecimal = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property
Property Let SeparadorDecimal(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sValor
End Property

Property Get SeparadorDeMiles() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("F", 16)
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    SeparadorDeMiles = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property

Property Let SeparadorDeMiles(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sValor
End Property

Property Get SeparadorDecimalDeMonedas() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("16", 16)
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    SeparadorDecimalDeMonedas = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property
Property Let SeparadorDecimalDeMonedas(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP, sValor
End Property
Property Get SeparadorDeMilesDeMonedas() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("17", 16)
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    SeparadorDeMilesDeMonedas = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property
Property Let SeparadorDeMilesDeMonedas(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP, sValor
End Property

Property Get FormatoDeHoras() As String
Dim sBuffer As String
Dim lBufferLen As Long
Dim lClaveFecha As Long
Dim sFormat  As String
        
    lClaveFecha = Base2Long("1003", 16)
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    
    GetLocaleInfo LOCALE_USER_DEFAULT, lClaveFecha, sBuffer, lBufferLen
    FormatoDeHoras = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    
End Property
Property Let FormatoDeHoras(sValor As String)
    SetLocaleInfo LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, sValor
End Property

Public Function Base2Long(s As String, ByVal nB As Integer) As Long
Dim s2 As String
Dim i As Long
Dim j As Long
Dim X As Long
Dim n As Boolean
Dim s3 As String

If Len(s) < 1 Then
    Base2Long = 0
    Exit Function
End If
s2 = UCase(s)
If Left$(s2, 1) = "-" Then
    n = True
    s2 = Right$(s2, Len(s2) - 1)
Else
    n = False
End If

j = 1
X = 0

For i = Len(s2) To 1 Step -1
    s3 = Mid$(s2, i, 1)
    Select Case s3
        Case "0" To "9":
            X = X + j * (Asc(s3) - 48)
        Case "A" To "Z":
            X = X + j * (Asc(s3) - 55)
    End Select

    j = j * nB
Next i

If n Then
    X = -X
End If

Base2Long = X
End Function


