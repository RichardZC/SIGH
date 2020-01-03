VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DONotaCreditoDebitoTipoNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: NotaCreditoDebitoTipoNota
'        Autor: Daniel Enrique Barrantes
'        Fecha: 04/06/2015 12:44:13
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_IdTipoNota As String
Dim ms_TipoNota As String
Dim ms_NroSerie As String
Dim ms_NroDocumento As String
Dim ms_NroDocumentoInicial As String
Dim ms_NroDocumentoFinal As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdTipoNota(sValue As String)
   ms_IdTipoNota = sValue
End Property
Property Get IdTipoNota() As String
   IdTipoNota = ms_IdTipoNota
End Property
Property Let TipoNota(sValue As String)
   ms_TipoNota = sValue
End Property
Property Get TipoNota() As String
   TipoNota = ms_TipoNota
End Property
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let NroDocumentoInicial(sValue As String)
   ms_NroDocumentoInicial = sValue
End Property
Property Get NroDocumentoInicial() As String
   NroDocumentoInicial = ms_NroDocumentoInicial
End Property
Property Let NroDocumentoFinal(sValue As String)
   ms_NroDocumentoFinal = sValue
End Property
Property Get NroDocumentoFinal() As String
   NroDocumentoFinal = ms_NroDocumentoFinal
End Property

