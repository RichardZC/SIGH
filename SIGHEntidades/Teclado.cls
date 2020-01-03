VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Teclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de tipo Teclado
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsLetra
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   una letra.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsLetra(KeyAscii As Integer) As Boolean

    Select Case KeyAscii
    Case vbKeyA + 32 To vbKeyZ + 32 'a to z
        CodigoAsciiEsLetra = True
    Case vbKeyA To vbKeyZ           'A to Z
        CodigoAsciiEsLetra = True
    Case 209, 241               'Ñ,ñ
        CodigoAsciiEsLetra = True
    Case 225, 233, 237, 243, 250 'á,é,í,ó,ú
        CodigoAsciiEsLetra = True
    Case 193, 201, 205, 211, 218 'Á,É,Í,Ó,Ú
        CodigoAsciiEsLetra = True
    Case 220, 252                   'Ü, ü
        CodigoAsciiEsLetra = True
    Case vbKeySpace
        CodigoAsciiEsLetra = True
    Case Else
        CodigoAsciiEsLetra = False
    End Select
    
End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsNumero
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   un numero.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsNumero(KeyAscii As Integer) As Boolean

    CodigoAsciiEsNumero = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack
        CodigoAsciiEsNumero = True
    Case Else
    End Select

End Function
'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsNumeroDeDomicilio
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   un numero.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsNumeroDeDomicilio(KeyAscii As Integer) As Boolean

    CodigoAsciiEsNumeroDeDomicilio = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9
        CodigoAsciiEsNumeroDeDomicilio = True
    Case Asc("S"), Asc("/"), Asc("N")
        CodigoAsciiEsNumeroDeDomicilio = True
    Case Else
    End Select

End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsDinero
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   un caracter valido para valores monetarios.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsDinero(KeyAscii As Integer) As Boolean

    CodigoAsciiEsDinero = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack
        CodigoAsciiEsDinero = True
    Case Asc("."), Asc(",")
        'MZD Ini 25/06/2005
        CodigoAsciiEsDinero = True
        'MZD Fin 25/06/2005
    Case Else
    End Select

End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsEspecial
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   una tecla especial
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsEspecial(KeyAscii As Integer) As Boolean

    CodigoAsciiEsEspecial = True
    
    Select Case KeyAscii
    Case 3  'CTRL C
    Case 22 'CTRL V
    Case 24 'CTRL X
    Case 26 'CTRL Z
    Case vbKeyBack
    Case vbKeyReturn
    'Case vbKeyRight, vbKeyDown
    'Case vbKeyLeft, vbKeyUp
    'Case vbKeyTab
    Case Else
        CodigoAsciiEsEspecial = False
    End Select

End Function
'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsComparador
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   una tecla especial
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsComparador(KeyAscii As Integer) As Boolean

    CodigoAsciiEsComparador = True
    
    Select Case KeyAscii
    Case Asc("="), Asc("<"), Asc(">")
    Case Asc("e"), Asc("n"), Asc("t"), Asc("r")
    Case Asc("E"), Asc("N"), Asc("T"), Asc("R")
    Case Else
        CodigoAsciiEsComparador = False
    End Select

End Function

'------------------------------------------------------------------------------------
'   Función:        RealizarNavegacion
'   Descripción:    Utiliza las teclas de navegacion para pasar de un control a otro
'                   en los TextBoxs y ComboBoxs
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'                   oObjeto:    El TextBox o ComboBox
'------------------------------------------------------------------------------------
Sub RealizarNavegacion(KeyAscii As Integer, oObjeto As Object)

    Select Case KeyAscii
    Case vbKeyBack
'        If TypeOf oObjeto Is TextBox Then
'            If oObjeto.SelStart = 0 Then
'                SendKeys "+{TAB}"
'            End If
'        End If
'        If TypeOf oObjeto Is ComboBox Then
'            If oObjeto.Style <> 2 Then  'Si no es dropdown
'                If oObjeto.SelStart = 0 Then
'                    SendKeys "+{TAB}"
'                End If
'            End If
'        End If
    
    Case vbKeyReturn
        SendKeys "{TAB}"
    Case vbKeyRight
        If TypeOf oObjeto Is TextBox Then
            If oObjeto.SelStart = Len(oObjeto.Text) Then
                SendKeys "{TAB}"
            End If
        End If
        If TypeOf oObjeto Is ComboBox Then
            If oObjeto.Style <> 2 Then
                If oObjeto.SelStart = Len(oObjeto.Text) Then
                    SendKeys "{TAB}"
                End If
            End If
        End If
    Case vbKeyDown
        If TypeOf oObjeto Is TextBox Then
            If oObjeto.SelStart = Len(oObjeto.Text) Then
                SendKeys "{TAB}"
            End If
        End If
    
    Case vbKeyLeft
        If TypeOf oObjeto Is TextBox Then
            If oObjeto.SelStart = 0 Then
                SendKeys "+{TAB}"
            End If
        End If
        If TypeOf oObjeto Is ComboBox Then
            If oObjeto.Style <> 2 Then
                If oObjeto.SelStart = 0 Then
                    SendKeys "+{TAB}"
                End If
            End If
        End If
    Case vbKeyUp
        
        If TypeOf oObjeto Is TextBox Then
            If oObjeto.SelStart = 0 Then
                SendKeys "+{TAB}"
            End If
        End If
    
    Case vbKeyTab
    End Select
    

End Sub

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsValidoParaNombre
'   Descripción:    Valida que el caracter ascii de una tecla es valido para
'                   el nombre de una persona.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsValidoParaNombre(KeyAscii As Integer) As Boolean
    
    CodigoAsciiEsValidoParaNombre = False
    If CodigoAsciiEsLetra(KeyAscii) Or KeyAscii = 32 Or KeyAscii = 180 Or KeyAscii = vbKeySpace Then
        CodigoAsciiEsValidoParaNombre = True
    End If

End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsLetraONumero
'   Descripción:    Valida que el caracter ascii de una tecla es letra o numero.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsLetraONumero(KeyAscii As Integer) As Boolean
    
    CodigoAsciiEsLetraONumero = False
    If CodigoAsciiEsLetra(KeyAscii) Or CodigoAsciiEsNumero(KeyAscii) Then
        CodigoAsciiEsLetraONumero = True
    End If

End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsCIE10
'   Descripción:    Valida que el caracter ascii de una tecla es letra o numero.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsCIE10(KeyAscii As Integer) As Boolean
    
    CodigoAsciiEsCIE10 = False
    If CodigoAsciiEsLetra(KeyAscii) Or CodigoAsciiEsNumero(KeyAscii) Or KeyAscii = Asc("+") Or KeyAscii = Asc(".") Then
        CodigoAsciiEsCIE10 = True
    End If

End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsParaDireccion
'   Descripción:    Valida que el caracter ascii es un tecla valida para una direccion.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsParaDireccion(KeyAscii As Integer) As Boolean
    
    CodigoAsciiEsParaDireccion = False
    If CodigoAsciiEsLetra(KeyAscii) Or CodigoAsciiEsNumero(KeyAscii) Or KeyAscii = 35 Or KeyAscii = 186 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 34 Or KeyAscii = 47 Then
        CodigoAsciiEsParaDireccion = True
    End If

End Function

Function CapitalizarNombres(ByVal sNombre As String) As String
Dim sValor1 As String
Dim sValor2 As String
Dim I As Integer
Dim sLetra As String

        CapitalizarNombres = ""
        
        If Len(sNombre) >= 1 Then
            sValor1 = Trim(sNombre)

            For I = 1 To Len(sValor1)
                
                sLetra = Mid(sValor1, I, 1)
                
                If sLetra = " " Then
                    I = I + 1
                    sLetra = Mid(sValor1, I, 1)
                    sValor2 = sValor2 + " " + UCase(sLetra)
                Else
                    If I = 1 Then
                        sValor2 = sValor2 + UCase(sLetra)
                    Else
                        sValor2 = sValor2 + LCase(sLetra)
                    End If

                End If
            Next

            CapitalizarNombres = sValor2

        End If
End Function

'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsParaHora
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   un dato hora.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsParaHora(KeyAscii As Integer) As Boolean

    CodigoAsciiEsParaHora = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9
        CodigoAsciiEsParaHora = True
    Case Asc(":")
        CodigoAsciiEsParaHora = True
    Case Else
    End Select

End Function
'------------------------------------------------------------------------------------
'   Función:        CodigoAsciiEsParaFecha
'   Descripción:    Valida que el caracter ascii de una tecla correponde a
'                   un dato fecha.
'   Parámetros:     Keyascii:   Es un entero que representa el codigo ascii
'------------------------------------------------------------------------------------
Function CodigoAsciiEsParaFecha(KeyAscii As Integer) As Boolean

    CodigoAsciiEsParaFecha = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9
        CodigoAsciiEsParaFecha = True
    Case Asc("/")
        CodigoAsciiEsParaFecha = True
    Case Else
    End Select

End Function

Function CodigoAsciiEsParaFechaHora(KeyAscii As Integer) As Boolean

    CodigoAsciiEsParaFechaHora = False
    
    Select Case KeyAscii
    Case vbKey0 To vbKey9
        CodigoAsciiEsParaFechaHora = True
    Case Asc("/")
        CodigoAsciiEsParaFechaHora = True
    Case Asc(":")
        CodigoAsciiEsParaFechaHora = True
    Case Asc(" ")
        CodigoAsciiEsParaFechaHora = True
    Case Else
    End Select

End Function

Function TextoEsSoloNumeros(lcTexto As String) As Boolean
    Dim lbOK As Boolean, lnFor As Integer
    lbOK = False
    If Len(lcTexto) > 0 Then
        If lcTexto <> "0" Then
            lbOK = True
            For lnFor = 1 To Len(lcTexto)
                If InStr("1234567890" & vbKeyBack, Mid(lcTexto, lnFor, 1)) = 0 Then
                   lbOK = False
                   Exit For
                End If
            Next
        End If
    Else
       lbOK = False
    End If
    TextoEsSoloNumeros = lbOK
End Function

Function DevuelveImporteRedondeado(lnImporte As Double, lnNroDecimalesRedondeo) As Double
    Dim lcImporte As String, lnDecimales As Integer
    lcImporte = Trim(str(lnImporte))
    lnDecimales = InStr(".", lcImporte)
    If (Len(lcImporte) - lnDecimales) > lnNroDecimalesRedondeo Then
       DevuelveImporteRedondeado = Round(lnImporte, lnNroDecimalesRedondeo)
    Else
       DevuelveImporteRedondeado = lnImporte
    End If
End Function

Function DevuelveTextoSINtildes(lcTexto As String) As String
    If Len(lcTexto) > 0 Then
        Dim lcTextoFinal As String, lnFor As Integer
        For lnFor = 1 To Len(lcTexto)
            Select Case Mid(lcTexto, lnFor, 1)
            Case "Á"
                 lcTextoFinal = lcTextoFinal & "A"
            Case "á"
                 lcTextoFinal = lcTextoFinal & "a"
            Case "É"
                 lcTextoFinal = lcTextoFinal & "E"
            Case "é"
                 lcTextoFinal = lcTextoFinal & "e"
            Case "Í"
                 lcTextoFinal = lcTextoFinal & "I"
            Case "í"
                 lcTextoFinal = lcTextoFinal & "i"
            Case "Ó"
                 lcTextoFinal = lcTextoFinal & "O"
            Case "ó"
                 lcTextoFinal = lcTextoFinal & "o"
            Case "Ú"
                 lcTextoFinal = lcTextoFinal & "U"
            Case "ú"
                 lcTextoFinal = lcTextoFinal & "u"
            Case Else
                 lcTextoFinal = lcTextoFinal & Mid(lcTexto, lnFor, 1)
            End Select
        Next
    End If
    DevuelveTextoSINtildes = lcTextoFinal
End Function



Function TextoAlmenosExisteAlgunaLetra(lcTexto As String) As Boolean
    On Error GoTo TexSal
    TextoAlmenosExisteAlgunaLetra = False
    If Len(lcTexto) > 0 Then
        Dim lcTextoFinal As String, lnFor As Integer
        For lnFor = 1 To Len(lcTexto)
            If CodigoAsciiEsLetra(Asc(Mid(lcTexto, lnFor, 1))) = True Then
               TextoAlmenosExisteAlgunaLetra = True
               Exit For
            End If
        Next
    End If
TexSal:
End Function



