VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOa_categoriaeess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para capa de estructura de la tabla Forma Farmacéutica
'        Programado por: Garay M.
'        Fecha: Marzo 2015
'
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_cat_IdCategoriaEESS As String
Dim ms_cat_Descripcion As String
Dim ms_cat_Abreviatura As String
Dim ml_cat_Nivel As Long
Dim ms_cat_IdEstado As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let cat_IdCategoriaEESS(sValue As String)
   ms_cat_IdCategoriaEESS = sValue
End Property
Property Get cat_IdCategoriaEESS() As String
   cat_IdCategoriaEESS = ms_cat_IdCategoriaEESS
End Property
Property Let cat_Descripcion(sValue As String)
   ms_cat_Descripcion = sValue
End Property
Property Get cat_Descripcion() As String
   cat_Descripcion = ms_cat_Descripcion
End Property
Property Let cat_Abreviatura(sValue As String)
   ms_cat_Abreviatura = sValue
End Property
Property Get cat_Abreviatura() As String
   cat_Abreviatura = ms_cat_Abreviatura
End Property
Property Let cat_Nivel(iValue As Long)
   ml_cat_Nivel = iValue
End Property
Property Get cat_Nivel() As Long
   cat_Nivel = ml_cat_Nivel
End Property
Property Let cat_IdEstado(sValue As String)
   ms_cat_IdEstado = sValue
End Property
Property Get cat_IdEstado() As String
   cat_IdEstado = ms_cat_IdEstado
End Property





