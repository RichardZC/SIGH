VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cadena"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Cadenas de Texto
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Function CompletarCadena(ByVal sCadena As String, sCaracter As String, iLongitud As Integer, sLado As String)
    
    ' OBJETIVO : Rellena la cadena enviada con el caracter pasado a la derecha o izquierda
    ' SOLICITA : sCadena    = cadena a transformar
    '            sCaracter  = Caracter de relleno
    '            iLongitud  = Longitud a rellenar
    '            piLado      = Lado a rellenar     (I,D. Por defecto I=izquierdo)
    ' RETORNA  : Valor en string
    
    
    Dim sCadenaTemp As String

    'Seteo el lado
    If sLado <> "D" Then sLado = "I"
    
    'Verifico caracter pasado y longitud de cadena enviada
    If Len(sCadena) >= iLongitud Or sCaracter = "" Then
        CompletarCadena = sCadena
        Exit Function
    End If
       
    sCadenaTemp = String$(iLongitud - Len(sCadena), sCaracter)

    Select Case sLado
    Case "I": CompletarCadena = sCadenaTemp & sCadena
    Case "D": CompletarCadena = sCadena & sCadenaTemp
    End Select


End Function

Function EliminarCadena(ByVal sCadena As String, sSubCadena As String) As String

    'OBJETIVO: Elimina de una cadena pasada una subcadena especifica
    'REQUIERE: sCadena   = Cadena a elminar subcadena
    '          sSubCadena = Subcadena a eliminar de la cadena

Dim iPosIni As Integer
Dim sCadenaNueva As String
Dim iLenSubCadena As Integer


    If Len(sCadena) < Len(sSubCadena) Then Exit Function


    sCadenaNueva = sCadena
    iLenSubCadena = Len(sSubCadena)

    iPosIni = InStr(sCadenaNueva, sSubCadena)

    Do While iPosIni <> 0
        If iPosIni = 1 Then
            sCadenaNueva = Mid(sCadenaNueva, 2)
        Else
            If iPosIni + iLenSubCadena - 1 <> Len(sCadenaNueva) Then
                sCadenaNueva = Left(sCadenaNueva, iPosIni - 1) + Mid(sCadenaNueva, iPosIni + iLenSubCadena)
           Else
                sCadenaNueva = Left(sCadenaNueva, iPosIni - 1)
            End If
        End If

        iPosIni = InStr(sCadenaNueva, sSubCadena)
    Loop

    EliminarCadena = sCadenaNueva

End Function

Function ExtraerCadena(ByVal sCadena As String, iUbicacion As Integer, sSeparador As String) As String
On Error GoTo ManejadorError
    'OBJETIVO: Extraer un item de una cadena separado por un caracter especifico pasado como parametro
    'REQUIERE: sCadena    : Cadena a tratar
    '          piUbicación : Ubicación del item a extraer
    '          sSeparador : Separador de items en la cadena
    'RETORNA : El item pedido o vacío en caso el ietm pedido sea mayor al número de items en la cadena

Dim sArrayCadena() As String

    sArrayCadena = Split(sCadena, sSeparador)
    ExtraerCadena = sArrayCadena(iUbicacion - 1)

Exit Function
ManejadorError:
    MsgBox Err.Description
    ExtraerCadena = ""
End Function

Function Lcero(ByVal psCadena As String, piLongitud As Integer) As String
    
    Lcero = CompletarCadena(psCadena, "0", piLongitud, "I")

End Function

Function OcurrenciasEnCadena(psCadena, psCaracter) As Integer

    Dim iContador   As Integer
    Dim iInicio     As Integer
    
    iContador = 0
    iInicio = 1
    Do While InStr(iInicio, psCadena, psCaracter) > 0
        iContador = iContador + 1
        iInicio = InStr(iInicio, psCadena, psCaracter) + 1
    Loop
    
    OcurrenciasEnCadena = iContador

End Function
Function ReemplazarCadena(sOriginal As String, sCadenaA As String, sCadenaR As String) As String
Dim sTemp As String
Dim lLng As Long
Dim lP As Long

    sTemp = sOriginal
    lLng = Len(sCadenaA)
    lP = InStr(sTemp, sCadenaA)
    
    Do While lP <> 0
        sTemp = Left(sTemp, lP - 1) + sCadenaR + Mid(sTemp, lP + lLng)
        lP = InStr(sTemp, sCadenaA)
    Loop

    ReemplazarCadena = sTemp

End Function

Function DevuelveARROBAS(lcTexto As String) As Integer
    Dim lnFor As Integer, lnNroArrobas As Integer
    lnNroArrobas = 0
    For lnFor = 1 To Len(lcTexto)
        If Mid(lcTexto, lnFor, 1) = "@" Then
           lnNroArrobas = lnNroArrobas + 1
        End If
    Next
    DevuelveARROBAS = lnNroArrobas
End Function

'Actualizado 29092014
'Function DevuelveNumeroRedondeado(lnTotalIngresado As Double) As Double
'        If Parametro387valorInt = 1 Then
'           DevuelveNumeroRedondeado = lnTotalIngresado
'           Exit Function
'        End If
'
'
'        Dim ln_Habia As Double, ln_Suma As Double, lcNumero As String
'        '**** En contra del PACIENTE
''        ln_Habia = Round(lnTotalIngresado, 2)
''        ln_Suma = Round(lnTotalIngresado, 1)
''        If ln_Suma >= ln_Habia Then
''           DevuelveNumeroRedondeado = lnTotalIngresado + ln_Suma - ln_Habia
''        Else
''           DevuelveNumeroRedondeado = Round(lnTotalIngresado + 0.05, 1) ' + ln_Habia - ln_Suma
''        End If
''
'
'
'        '**** A favor del PACIENTE
'        lcNumero = Trim(str(lnTotalIngresado))
'        If InStr(lcNumero, ".") > 0 Then
'            If CCur(Left(lcNumero, InStr(lcNumero, ".") + 1)) + 0.05 > lnTotalIngresado Then
'                DevuelveNumeroRedondeado = CCur(Left(lcNumero, InStr(lcNumero, ".") + 1))
'            Else
'                DevuelveNumeroRedondeado = CCur(Left(lcNumero, InStr(lcNumero, ".") + 1)) + 0.05
'            End If
'        Else
'            DevuelveNumeroRedondeado = lnTotalIngresado
'        End If
'End Function
'



Function RetornaPrimerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaPrimerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaPrimerNombre = Trim(Left(lcPrimerSegundoNombreJuntos, ln))
    Else
       RetornaPrimerNombre = lcPrimerSegundoNombreJuntos
    End If
End Function

Function RetornaSegundoNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaSegundoNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaSegundoNombre = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
       ln = InStr(RetornaSegundoNombre, " ")
       If ln > 0 Then
          RetornaSegundoNombre = Trim(Left(RetornaSegundoNombre, ln))
       End If
    End If
End Function

Function RetornaTercerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer, lcNombre1 As String, lcNombre2 As String, lcNombre3 As String
    RetornaTercerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       lcNombre1 = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
       ln = InStr(lcNombre1, " ")
       If ln > 0 Then
          lcNombre2 = Trim(Left(lcNombre1, ln))
          RetornaTercerNombre = Trim(Mid(lcNombre1, ln + 1, 100))
       End If
    End If
End Function

Public Function RetornaDescFechaNacimientoCalculada(lbFNacimientoCalculada As Boolean, _
            Optional agregarParentisis As Boolean = True)
    Dim lcObservacion As String
    lcObservacion = ""
    If lbFNacimientoCalculada = True Then
        If agregarParentisis = True Then
            lcObservacion = "(FCalculada)"
        Else
            lcObservacion = "FCalculada"
        End If
    End If
    RetornaDescFechaNacimientoCalculada = lcObservacion
End Function



Public Function SISdevuelveCategoriaCS() As String
       SISdevuelveCategoriaCS = "/03/04/11/"
End Function
Public Function SISdevuelveCategoriaPS() As String
       SISdevuelveCategoriaPS = "/01/02/12/"
End Function
Public Function SISdevuelveCategoriaHospital() As String
       SISdevuelveCategoriaHospital = "/05/06/07/08/"
End Function

Public Function getLongitudRuc() As Integer
    getLongitudRuc = 11
End Function

Public Function getLongitudCodigoRENAES() As Integer
    getLongitudCodigoRENAES = 8
End Function

Public Function EsRucCorrecto(sNumeroRuc As String) As Boolean
    If sNumeroRuc = "" Then
        EsRucCorrecto = False
        Exit Function
    End If
    If Not IsNumeric(sNumeroRuc) Then
        EsRucCorrecto = False
        Exit Function
    End If
    If Len(sNumeroRuc) <> getLongitudRuc() Then
        EsRucCorrecto = False
        Exit Function
    End If
    EsRucCorrecto = validarRuc(sNumeroRuc)
End Function

Public Function validarRuc(sNumeroRuc As String) As Boolean
    Dim I As Integer, X As Integer
    Dim suma As Integer, resto As Integer, digito As Integer
    
    sNumeroRuc = Trim(sNumeroRuc)
    If IsNumeric(sNumeroRuc) Then
        If Len(sNumeroRuc) = 8 Then
            suma = 0
            For I = 0 To Len(sNumeroRuc) - 2
                digito = Val(Mid(sNumeroRuc, I + 1, 1))
                If I = 0 Then
                    suma = suma + (digito * 2)
                Else
                    suma = suma + (digito * (Len(sNumeroRuc) - I))
                End If
            Next I
            resto = suma Mod 11
            If resto = 1 Then: resto = 11
            If resto + Val(Mid(sNumeroRuc, Len(sNumeroRuc), 1)) = 11 Then
              validarRuc = True
              Exit Function
            End If
        ElseIf Len(sNumeroRuc) = 11 Then
            suma = 0
            X = 6
            For I = 0 To Len(sNumeroRuc) - 2
            'for (i=0; i<sNumeroRuc.length-1;i++){
                If I = 4 Then: X = 8
                
                digito = Val(Mid(sNumeroRuc, I + 1, 1)) 'sNumeroRuc.charAt(i) - '0';
                X = X - 1
                If I = 0 Then
                    suma = suma + (digito * X)
                Else
                    suma = suma + (digito * X)
                End If
            Next I
            resto = suma Mod 11
            resto = 11 - resto
            
            If resto >= 10 Then: resto = resto - 10
            If resto = Val(Mid(sNumeroRuc, Len(sNumeroRuc), 1)) Then
            'if  resto = sNumeroRuc.charAt( sNumeroRuc.length-1 ) - '0' then
              validarRuc = True
              Exit Function
            End If
        End If
    End If
    validarRuc = False
End Function


Public Function EsCodigoRENAESCorrecto(sNumero As String) As Boolean
    If sNumero = "" Then
        EsCodigoRENAESCorrecto = False
        Exit Function
    End If
    If Not IsNumeric(sNumero) Then
        EsCodigoRENAESCorrecto = False
        Exit Function
    End If
    If Len(sNumero) <> getLongitudCodigoRENAES() Then
        EsCodigoRENAESCorrecto = False
        Exit Function
    End If
    EsCodigoRENAESCorrecto = validarCodigoRENAES(sNumero)
End Function

Public Function digitoControlModulo11(sNumero As String)
    Dim I As Integer
    Dim suma As Integer, factor As Integer, digito As Integer
    Dim resto  As Integer
    suma = 0
    factor = 2
    For I = Len(sNumero) To 1 Step -1
        If factor = 8 Then
            factor = 2
        End If
        digito = Val(Mid(sNumero, I, 1))
        suma = suma + (digito * factor)
        factor = factor + 1
    Next I
    resto = suma Mod 11
    resto = 11 - resto
    Select Case resto
        Case 10:
            resto = 1
        Case 11:
            resto = 0
    End Select
    digitoControlModulo11 = resto
End Function

Public Function validarCodigoRENAES(sNumero As String) As Boolean
    Dim digito As Integer
    
    sNumero = Trim(sNumero)
    If IsNumeric(sNumero) Then
'        If Len(sNumero) = getLongitudCodigoRENAES() Then
            digito = digitoControlModulo11(Mid(sNumero, 1, Len(sNumero) - 1))
            
            If digito = Val(Mid(sNumero, Len(sNumero), 1)) Then
              validarCodigoRENAES = True
              Exit Function
            End If
'        End If
    End If
    validarCodigoRENAES = False
End Function
  
Public Function FormatoCodigoRENAES(sCodigo As String, _
            lFuenteReanes As sghRENAESFuente) As String
    Dim sCodigoFormat As String
    Select Case lFuenteReanes
        Case sghRENAESFuente.GALENHOS:
            sCodigoFormat = Format(sCodigo, "00000")
        Case sghRENAESFuente.SIS:
            sCodigoFormat = Format(sCodigo, "0000000000")
        Case sghRENAESFuente.SUNASA:
            sCodigoFormat = CStr(Val(sCodigo))
        Case sghRENAESFuente.RENAESNORMA:
            sCodigoFormat = Format(sCodigo, "00000000")
    End Select
    FormatoCodigoRENAES = sCodigoFormat
End Function


'debb-02/05/2016
Function DevuelveFuenteDeLineaDelError(lnLinea As Integer, lcProcedimiento As String, lcPrograma As String) As String
    Dim lcCadena1 As String, lbRecienBuscaLinea As Boolean, lcLineaStr As String
    lcLineaStr = Trim(str(lnLinea))
    DevuelveFuenteDeLineaDelError = Chr(13) & "Procedimiento: " & lcProcedimiento & Chr(13) & _
                                              "Linea: " & lcLineaStr & Chr(13)
    lbRecienBuscaLinea = False
    Open App.Path & "\archivos\" & lcPrograma For Input As #1
    Do While Not EOF(1)
        Line Input #1, lcCadena1
        If Left(lcCadena1, Len(lcProcedimiento)) = lcProcedimiento Then
           lbRecienBuscaLinea = True
        End If
        If lbRecienBuscaLinea = True And Left(lcCadena1, Len(lcLineaStr)) = lcLineaStr Then
           DevuelveFuenteDeLineaDelError = DevuelveFuenteDeLineaDelError & "Fuente: " & lcCadena1
           Exit Do
        End If
    Loop
    Close #1
End Function

Function RetornaTextoSinChr13Chr10(lcTexto As String) As String
    Dim lnFor As Integer, lcTextoOK As String
    lcTextoOK = ""
    For lnFor = 1 To Len(lcTexto)
        If Mid(lcTexto, lnFor, 1) <> Chr(13) And Mid(lcTexto, lnFor, 1) <> Chr(10) And Mid(lcTexto, lnFor, 1) <> "|" Then
           lcTextoOK = lcTextoOK & Mid(lcTexto, lnFor, 1)
        End If
    Next
    RetornaTextoSinChr13Chr10 = lcTextoOK
End Function

Function DevuelveSinApellido() As String
    DevuelveSinApellido = "__________"
End Function
