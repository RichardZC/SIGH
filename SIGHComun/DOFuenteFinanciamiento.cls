VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFuenteFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FuentesFinanciamiento
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdTipoFinanciamiento As Long
Dim ms_Descripcion As String
Dim ml_idFuenteFinanciamiento As Long
Dim ml_idTipoConceptoFarmacia As Long
Dim ml_UtilizadoEn As Long
Dim ml_idAreaTramitaSeguros As Long
Dim ml_CodigoFuenteFinanciamientoSEM As String
Dim mb_EsUsadoEnCaja As Boolean
Dim ml_CodigoHIS As String
Dim ml_idTipoFinanciador As Long
Dim ms_Codigo As String
Dim ml_TieneEPS As Long
Dim ml_usadoEnFUnidosis As Long
Dim ml_CuentaParaUnidosis As Long
Property Let CuentaParaUnidosis(lValue As Long)
   ml_CuentaParaUnidosis = lValue
End Property
Property Get CuentaParaUnidosis() As Long
   CuentaParaUnidosis = ml_CuentaParaUnidosis
End Property


Property Let usadoEnFUnidosis(lValue As Long)
   ml_usadoEnFUnidosis = lValue
End Property
Property Get usadoEnFUnidosis() As Long
   usadoEnFUnidosis = ml_usadoEnFUnidosis
End Property

Property Let TieneEPS(lValue As Long)
   ml_TieneEPS = lValue
End Property
Property Get TieneEPS() As Long
   TieneEPS = ml_TieneEPS
End Property

Property Let CodigoHIS(lValue As String)
   ml_CodigoHIS = lValue
End Property
Property Get CodigoHIS() As String
   CodigoHIS = ml_CodigoHIS
End Property

Property Let EsUsadoEnCaja(lValue As Boolean)
   mb_EsUsadoEnCaja = lValue
End Property
Property Get EsUsadoEnCaja() As Boolean
   EsUsadoEnCaja = mb_EsUsadoEnCaja
End Property

Property Let CodigoFuenteFinanciamientoSEM(lValue As String)
   ml_CodigoFuenteFinanciamientoSEM = lValue
End Property
Property Get CodigoFuenteFinanciamientoSEM() As String
   CodigoFuenteFinanciamientoSEM = ml_CodigoFuenteFinanciamientoSEM
End Property

Property Let idAreaTramitaSeguros(lValue As Long)
   ml_idAreaTramitaSeguros = lValue
End Property
Property Get idAreaTramitaSeguros() As Long
   idAreaTramitaSeguros = ml_idAreaTramitaSeguros
End Property
Property Let UtilizadoEn(lValue As Long)
   ml_UtilizadoEn = lValue
End Property
Property Get UtilizadoEn() As Long
   UtilizadoEn = ml_UtilizadoEn
End Property
Property Let idTipoConceptoFarmacia(lValue As Long)
   ml_idTipoConceptoFarmacia = lValue
End Property
Property Get idTipoConceptoFarmacia() As Long
   idTipoConceptoFarmacia = ml_idTipoConceptoFarmacia
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let idTipoFinanciador(lValue As Long)
   ml_idTipoFinanciador = lValue
End Property
Property Get idTipoFinanciador() As Long
   idTipoFinanciador = ml_idTipoFinanciador
End Property
Property Let Codigo(lValue As String)
   ms_Codigo = lValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property


