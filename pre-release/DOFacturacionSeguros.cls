VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFacturacionSeguros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla DOFacturacionSeguros
'        Programado por: Castro W
'        Fecha: Abril 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_NroPlaca As String
Dim ms_CodigoAutorizacion As String
Dim ml_IdTipoFinanciamiento As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdFacturacionSeguro As Long
Dim mc_TotalAsegurado As Currency
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let NroPlaca(sValue As String)
   ms_NroPlaca = sValue
End Property
Property Get NroPlaca() As String
   NroPlaca = ms_NroPlaca
End Property
Property Let CodigoAutorizacion(sValue As String)
   ms_CodigoAutorizacion = sValue
End Property
Property Get CodigoAutorizacion() As String
   CodigoAutorizacion = ms_CodigoAutorizacion
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdFacturacionSeguro(lValue As Long)
   ml_IdFacturacionSeguro = lValue
End Property
Property Get IdFacturacionSeguro() As Long
   IdFacturacionSeguro = ml_IdFacturacionSeguro
End Property
Property Let TotalAsegurado(cValue As Currency)
   mc_TotalAsegurado = cValue
End Property
Property Get TotalAsegurado() As Currency
   TotalAsegurado = mc_TotalAsegurado
End Property


