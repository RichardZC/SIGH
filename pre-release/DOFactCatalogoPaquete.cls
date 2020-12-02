VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactCatalogoPaquete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoPaquete
'        Programado por: Barrantes D
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactPaquete As Long
Dim ms_codigo As String
Dim ms_Descripcion As String
Dim ml_IdTipoFinanciamiento As Long
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim ml_idEstado As Long
Dim ml_TipoPaquete As Long
Dim ml_EsItemFarmacia As Long
Dim ms_cpt As String
Property Let cpt(sValue As String)
   ms_cpt = sValue
End Property
Property Get cpt() As String
   cpt = ms_cpt
End Property

Property Let EsItemFarmacia(lValue As Long)
   ml_EsItemFarmacia = lValue
End Property
Property Get EsItemFarmacia() As Long
   EsItemFarmacia = ml_EsItemFarmacia
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idFactPaquete(iValue As Long)
   ml_idFactPaquete = iValue
End Property
Property Get idFactPaquete() As Long
   idFactPaquete = ml_idFactPaquete
End Property
Property Let codigo(sValue As String)
   ms_codigo = sValue
End Property
Property Get codigo() As String
   codigo = ms_codigo
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idEstado(iValue As Long)
   ml_idEstado = iValue
End Property
Property Get idEstado() As Long
   idEstado = ml_idEstado
End Property
Property Let TipoPaquete(iValue As Long)
   ml_TipoPaquete = iValue
End Property
Property Get TipoPaquete() As Long
   TipoPaquete = ml_TipoPaquete
End Property

