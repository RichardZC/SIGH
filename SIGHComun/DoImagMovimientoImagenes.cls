VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoImagMovimientoImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla ImagMovimientoImagenes
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdMovimiento As Long
Dim ml_IdOrden As Long
Dim ml_CorrelativoAnual As Long
Dim ml_IdCuentaAtencion As Long
Dim ml_IdComprobantePago As Long
Dim ml_idPersonaTomaImagen As Long
Dim ms_idPersonaRecoge As Long
Dim ms_zonaRayosX As String
Dim mc_PorcInformeRadiolog As Currency
Dim ms_ResultadoFinal As String
Dim ml_EsContraste As Long
Dim ml_EsContrasteIonico As Long
Dim ml_IdDiagnostico As Long
Dim ml_EsDiagnosticoDefinitivo As Long
Dim mda_Eo_FUM As Date
Dim ms_Eo_Gestantes As String
Dim ms_Eo_Partos As String
Dim ml_Eo_EG As Long
Dim mda_FechaNacimiento As Date
Dim ml_IdTipoSexo As Long
Dim ms_Paciente As String
Dim ms_HcRX As String
Property Get HcRX() As String
   HcRX = ms_HcRX
End Property
Property Let HcRX(sValue As String)
   ms_HcRX = sValue
End Property

Property Get Paciente() As String
   Paciente = ms_Paciente
End Property
Property Let Paciente(sValue As String)
   ms_Paciente = sValue
End Property
Property Let IdTipoSexo(lValue As Long)
   ml_IdTipoSexo = lValue
End Property
Property Get IdTipoSexo() As Long
   IdTipoSexo = ml_IdTipoSexo
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property

Property Let Eo_FUM(daValue As Date)
   mda_Eo_FUM = daValue
End Property
Property Get Eo_FUM() As Date
   Eo_FUM = mda_Eo_FUM
End Property
Property Let Eo_Gestantes(sValue As String)
   ms_Eo_Gestantes = sValue
End Property
Property Get Eo_Gestantes() As String
   Eo_Gestantes = ms_Eo_Gestantes
End Property
Property Let Eo_Partos(sValue As String)
   ms_Eo_Partos = sValue
End Property
Property Get Eo_Partos() As String
   Eo_Partos = ms_Eo_Partos
End Property
Property Let Eo_EG(lValue As Long)
   ml_Eo_EG = lValue
End Property
Property Get Eo_EG() As Long
   Eo_EG = ml_Eo_EG
End Property


Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdMovimiento(iValue As Long)
   ml_IdMovimiento = iValue
End Property
Property Get IdMovimiento() As Long
   IdMovimiento = ml_IdMovimiento
End Property
Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
Property Let CorrelativoAnual(iValue As Long)
   ml_CorrelativoAnual = iValue
End Property
Property Get CorrelativoAnual() As Long
   CorrelativoAnual = ml_CorrelativoAnual
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdComprobantePago(iValue As Long)
   ml_IdComprobantePago = iValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property
Property Let idPersonaTomaImagen(iValue As Long)
   ml_idPersonaTomaImagen = iValue
End Property
Property Get idPersonaTomaImagen() As Long
   idPersonaTomaImagen = ml_idPersonaTomaImagen
End Property
Property Let idPersonaRecoge(sValue As Long)
   ms_idPersonaRecoge = sValue
End Property
Property Get idPersonaRecoge() As Long
   idPersonaRecoge = ms_idPersonaRecoge
End Property
Property Let zonaRayosX(sValue As String)
   ms_zonaRayosX = sValue
End Property
Property Get zonaRayosX() As String
   zonaRayosX = ms_zonaRayosX
End Property

Property Let PorcInformeRadiolog(cValue As Currency)
   mc_PorcInformeRadiolog = cValue
End Property
Property Get PorcInformeRadiolog() As Currency
   PorcInformeRadiolog = mc_PorcInformeRadiolog
End Property
Property Let ResultadoFinal(sValue As String)
   ms_ResultadoFinal = sValue
End Property
Property Get ResultadoFinal() As String
   ResultadoFinal = ms_ResultadoFinal
End Property
Property Let EsContraste(lValue As Long)
   ml_EsContraste = lValue
End Property
Property Get EsContraste() As Long
   EsContraste = ml_EsContraste
End Property
Property Let EsContrasteIonico(lValue As Long)
   ml_EsContrasteIonico = lValue
End Property
Property Get EsContrasteIonico() As Long
   EsContrasteIonico = ml_EsContrasteIonico
End Property
Property Let IdDiagnostico(lValue As Long)
   ml_IdDiagnostico = lValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let EsDiagnosticoDefinitivo(lValue As Long)
   ml_EsDiagnosticoDefinitivo = lValue
End Property
Property Get EsDiagnosticoDefinitivo() As Long
   EsDiagnosticoDefinitivo = ml_EsDiagnosticoDefinitivo
End Property


