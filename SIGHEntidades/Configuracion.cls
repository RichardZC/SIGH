VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Conexión a la BD
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oReg As New Registro
Const RutaDeRegistroProyecto2000 = "Software\Digital Works Corporation"
Const RutaDeRegistroSIGH = "Software\Digital Works Corporation\SIGH"
'GLCC 02/11/20 CAMBIO36 INICIO
'ANTERIOR: Const wxNueve As String = "9"
'GLCC 02/11/20 CAMBIO36 FIN
'**********************************
'BASE DE DATOS
'**********************************
Property Get CadenaConexion() As String
  On Error Resume Next
  Dim oCrypKey As New CrypKey.Util
  CadenaConexion = oCrypKey.DecryptString(oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexion:"))
  'CadenaConexion = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexion:")
  'Frank
'  CadenaConexion = "Provider=SQLOLEDB.1;User Id=sa;Password=123456;Initial Catalog=sigh;Data Source=SCEIP138758"
End Property

Property Let CadenaConexion(sValor As String)
  oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "CadenaConexion:", sValor
End Property

Property Get CadenaConexionShape() As String
  Dim oCrypKey As New CrypKey.Util
  On Error Resume Next
  CadenaConexionShape = oCrypKey.DecryptString(oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexionShape:"))
End Property

Property Let CadenaConexionShape(sValor As String)
  oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "CadenaConexionShape:", sValor
End Property

Property Get CadenaConexionIntegrada() As String
  Dim oCrypKey As New CrypKey.Util
  CadenaConexionIntegrada = oCrypKey.DecryptString(oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexionIntegrada:"))
  'CadenaConexionIntegrada = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexionIntegrada:")
End Property

Property Let CadenaConexionIntegrada(sValor As String)
  oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "CadenaConexionIntegrada:", sValor
End Property

'**********************************
'SISTEMA
'**********************************
Property Get NombreSistema() As String
  NombreSistema = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Nombre:")
End Property

Property Let NombreSistema(sValor As String)
  oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Nombre:", sValor
End Property

'**********************************
'      USUARIO
'**********************************
Property Get Usuario() As String
  Usuario = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Usuario:")
End Property

Property Let Usuario(sValor As String)
  oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Usuario:", sValor
End Property

Property Get NombreUsuario() As String
  NombreUsuario = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "NombreUsuario:")
End Property

Property Let NombreUsuario(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "NombreUsuario:", sValor
End Property

Property Get GenerarHistoriaClinicaSiempre() As String
    GenerarHistoriaClinicaSiempre = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "GenerarHistoriaClinicaSiempre:")
End Property
Property Let GenerarHistoriaClinicaSiempre(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "GenerarHistoriaClinicaSiempre:", sValor
End Property

Property Get GrillaConFilasBicolor() As String
    GrillaConFilasBicolor = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "GrillaConFilasBicolor:")
End Property
Property Let GrillaConFilasBicolor(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "GrillaConFilasBicolor:", sValor
End Property

Property Get RutaAdobeReader() As String
    RutaAdobeReader = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "RutaAdobeReader:")
End Property
Property Let RutaAdobeReader(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "RutaAdobeReader:", sValor
End Property

Property Get EstanciaMaxHospitalizacion() As String
    EstanciaMaxHospitalizacion = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "EstanciaMaxHospitalizacion:")
    If EstanciaMaxHospitalizacion = "" Then
       oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "EstanciaMaxHospitalizacion:", "28"
    End If
End Property
Property Let EstanciaMaxHospitalizacion(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "EstanciaMaxHospitalizacion:", sValor
End Property

Property Get TipoActualizacionDeCupos() As String
    TipoActualizacionDeCupos = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "TipoActualizacionDeCupos:")
End Property
Property Let TipoActualizacionDeCupos(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "TipoActualizacionDeCupos:", sValor
End Property

Property Get IntervaloActualizacionCupos() As String
    IntervaloActualizacionCupos = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "IntervaloActualizacionCupos:")
End Property
Property Let IntervaloActualizacionCupos(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "IntervaloActualizacionCupos:", sValor
End Property


Private Sub Class_Terminate()
    Set oReg = Nothing
End Sub


Property Get EsCentroSalud() As String
  EsCentroSalud = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "EsCentroSalud:")
End Property
'debb-14/07/2015
Property Get IdFarmaciaParaReceta() As String
    IdFarmaciaParaReceta = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "IdFarmaciaParaReceta:")
End Property
Property Let IdFarmaciaParaReceta(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "IdFarmaciaParaReceta:", sValor
End Property
'debb-12/10/2015
Property Get BuscarSoloIzquierda() As String
    BuscarSoloIzquierda = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "BuscarSoloIzquierda:")
End Property
'debb-12/10/2015
Property Let BuscarSoloIzquierda(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "BuscarSoloIzquierda:", sValor
End Property

'debb-20/10/2015
Property Get TurnoMovimientoHC() As String
    TurnoMovimientoHC = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "TurnoMovimientoHC:")
End Property
'debb-20/10/2015
Property Let TurnoMovimientoHC(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "TurnoMovimientoHC:", sValor
End Property


'debb-19/03/2018
Property Get ConexionCopia() As String
    ConexionCopia = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "ConexionCopia:")
End Property
Property Let ConexionCopia(sValor As String)
    sValor = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "CadenaConexion:")
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "ConexionCopia:", sValor
End Property

Property Get Parametro351() As String
    Parametro351 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro351:")
End Property
Property Let Parametro351(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro351:", sValor
End Property


Function HCigualDNI_DevuelveHistoriaConCerosIzquierda(lcHistoriaOriginal As String, lbEsImpresionExcel As Boolean) As String
           Dim lcParametro351 As String
           lcParametro351 = Parametro351
           If lbEsImpresionExcel = True Then
              HCigualDNI_DevuelveHistoriaConCerosIzquierda = "'" & lcHistoriaOriginal
           Else
              HCigualDNI_DevuelveHistoriaConCerosIzquierda = lcHistoriaOriginal
           End If
           'GLCC 02/11/20 CAMBIO36 INICIO
           'ANTERIOR: If lcParametro351 = "S" And Len(lcHistoriaOriginal) = 9 And Left(lcHistoriaOriginal, 1) = wxNueve Then
           If lcParametro351 = "S" And Len(lcHistoriaOriginal) = 8 Then
          'GLCC 02/11/20 CAMBIO36 FIN
              If lbEsImpresionExcel = True Then
                 HCigualDNI_DevuelveHistoriaConCerosIzquierda = "'" & Right(lcHistoriaOriginal, 8)
              Else
                 HCigualDNI_DevuelveHistoriaConCerosIzquierda = Right(lcHistoriaOriginal, 8)
              End If
           End If
End Function

Function HCigualDNI_AgregaNUEVEaLaHistoria(lcHistoriaOriginal As String) As String
           Dim lcParametro351 As String
           lcParametro351 = Parametro351
           HCigualDNI_AgregaNUEVEaLaHistoria = lcHistoriaOriginal
           If lcParametro351 = "S" And Len(lcHistoriaOriginal) = 8 Then
           'GLCC 02/11/20 CAMBIO36 INICIO
           'ANTERIOR: HCigualDNI_AgregaNUEVEaLaHistoria = wxNueve & lcHistoriaOriginal
              HCigualDNI_AgregaNUEVEaLaHistoria = lcHistoriaOriginal
         'GLCC 02/11/20 CAMBIO36 FIN
           End If
End Function
Property Get Parametro550() As String
    Parametro550 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro550:")
End Property
Property Let Parametro550(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro550:", sValor
End Property
Property Get Parametro551() As String
    Parametro551 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro551:")
End Property
Property Let Parametro551(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro551:", sValor
End Property


Property Get Lx_LabVacio() As String
    Lx_LabVacio = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Lx_LabVacio:")
End Property
Property Let Lx_LabVacio(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Lx_LabVacio:", sValor
End Property


Property Get Parametro387valorInt() As String
    Parametro387valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro387valorInt:")
End Property
Property Let Parametro387valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro387valorInt:", sValor
End Property

Property Get Parametro322valorInt() As String
    Parametro322valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro322valorInt:")
End Property
Property Let Parametro322valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro322valorInt:", sValor
End Property
Property Get Parametro282valorInt() As String
    Parametro282valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro282valorInt:")
End Property
Property Let Parametro282valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro282valorInt:", sValor
End Property
Property Get Parametro378valorInt() As String
    Parametro378valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro378valorInt:")
End Property
Property Let Parametro378valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro378valorInt:", sValor
End Property
'Se agregó la escritura de un parámetro:
Property Let Acceso(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Acceso:", sValor
End Property

Function DevuelveNumeroRedondeado(lnTotalIngresado As Double) As Double
        If Parametro387valorInt = 1 Then
           DevuelveNumeroRedondeado = lnTotalIngresado
           Exit Function
        End If
        
        
        Dim ln_Habia As Double, ln_Suma As Double, lcNumero As String
        '**** En contra del PACIENTE
'        ln_Habia = Round(lnTotalIngresado, 2)
'        ln_Suma = Round(lnTotalIngresado, 1)
'        If ln_Suma >= ln_Habia Then
'           DevuelveNumeroRedondeado = lnTotalIngresado + ln_Suma - ln_Habia
'        Else
'           DevuelveNumeroRedondeado = Round(lnTotalIngresado + 0.05, 1) ' + ln_Habia - ln_Suma
'        End If
'
        
        
        '**** A favor del PACIENTE
        lcNumero = Trim(str(lnTotalIngresado))
        If InStr(lcNumero, ".") > 0 Then
            If CCur(Left(lcNumero, InStr(lcNumero, ".") + 1)) + 0.05 > lnTotalIngresado Then
                DevuelveNumeroRedondeado = CCur(Left(lcNumero, InStr(lcNumero, ".") + 1))
            Else
                DevuelveNumeroRedondeado = CCur(Left(lcNumero, InStr(lcNumero, ".") + 1)) + 0.05
            End If
        Else
            DevuelveNumeroRedondeado = lnTotalIngresado
        End If
        '
End Function


Property Get RutaImagenConPermiso() As String
    RutaImagenConPermiso = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "RutaImagenConPermiso:")
End Property
Property Let RutaImagenConPermiso(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "RutaImagenConPermiso:", sValor
End Property

Property Get NombreCamarita() As String
    NombreCamarita = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "NombreCamarita:")
End Property
Property Let NombreCamarita(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "NombreCamarita:", sValor
End Property


Property Get Parametro556() As String
    Parametro556 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro556:")
End Property
Property Let Parametro556(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro556:", sValor
End Property

Property Get Parametro560() As String
    Parametro560 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro560:")
End Property
Property Let Parametro560(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro560:", sValor
End Property
Property Get Parametro561() As String
    Parametro561 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro561:")
End Property
Property Let Parametro561(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro561:", sValor
End Property
Property Get Parametro562() As String
    Parametro562 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro562:")
End Property
Property Let Parametro562(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro562:", sValor
End Property


Property Get TurnosMananaTarde() As String
    TurnosMananaTarde = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "TurnosMananaTarde:")
End Property
Property Let TurnosMananaTarde(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "TurnosMananaTarde:", sValor
End Property


Property Get Parametro301valorInt() As String
    Parametro301valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro301valorInt:")
End Property
Property Let Parametro301valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro301valorInt:", sValor
End Property


Property Get Parametro503valorInt() As String
    Parametro503valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro503valorInt:")
End Property
Property Let Parametro503valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro503valorInt:", sValor
End Property

Property Get SoftwareImagen() As String
    SoftwareImagen = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "SoftwareImagen:")
End Property
Property Let SoftwareImagen(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "SoftwareImagen:", sValor
End Property


Property Get Pto() As String
    Pto = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Pto:")
End Property
Property Let Pto(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Pto:", sValor
End Property

Property Get Parametro568() As String
    Parametro568 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro568:")
End Property
Property Let Parametro568(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro568:", sValor
End Property
Property Get Parametro569() As String
    Parametro569 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro569:")
End Property
Property Let Parametro569(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro569:", sValor
End Property
Property Get ParaAuditoria() As String
    ParaAuditoria = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "ParaAuditoria:")
End Property
Property Let ParaAuditoria(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "ParaAuditoria:", sValor
End Property

Sub ParaAuditoriaPorCadaDato(lnOpcion As sghAuditoria, lcDato As String)
    On Error GoTo errDatoR
    Dim lcRegEdit As String
    Select Case lnOpcion
    Case sghAudLimpiar
         ParaAuditoria = ""
    Case sghAudGrabaRegEdit
         If lcDato <> "" Then
            lcRegEdit = ParaAuditoria
            ParaAuditoria = lcRegEdit & "/" & lcDato
         End If
    End Select
errDatoR:
End Sub


Property Get Parametro378() As String
    Parametro378 = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro378:")
End Property
Property Let Parametro378(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro378:", sValor
End Property


Sub AbreConexionSIGH(oConexion As Connection)
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open CadenaConexion
End Sub

Property Get Parametro573valorInt() As String
    Parametro573valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro573valorInt:")
End Property
Property Let Parametro573valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro573valorInt:", sValor
End Property



Property Get ImpresoraDefaultDeEstaPC() As String
    ImpresoraDefaultDeEstaPC = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "ImpresoraDefaultDeEstaPC:")
End Property
Property Let ImpresoraDefaultDeEstaPC(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "ImpresoraDefaultDeEstaPC:", sValor
End Property
Property Get Parametro583valorInt() As String
    Parametro583valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro583valorInt:")
End Property
Property Let Parametro583valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro583valorInt:", sValor
End Property
Property Get Parametro584valorInt() As String
    Parametro584valorInt = oReg.ObtenerCadena(oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH, "Parametro584valorInt:")
End Property
Property Let Parametro584valorInt(sValor As String)
    oReg.GuardarCadena oReg.HKEY_LOCAL_MACHINE, RutaDeRegistroSIGH + "\", "Parametro584valorInt:", sValor
End Property


