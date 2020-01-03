VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCatalogoBienesFarmacia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para Mantenimiento de Bienes e Insumos, solo particular
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim mi_Opcion As sghOpciones
Dim ml_IdProducto As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

Dim oFormulario As New CatalogoBienesSoloFarmacia
Dim ml_IdPlanCatalogo As Long
Property Let IdPlanCatalogo(lValue As Long)
   ml_IdPlanCatalogo = lValue
End Property
Property Get IdPlanCatalogo() As Long
   IdPlanCatalogo = ml_IdPlanCatalogo
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_IdProducto
End Property

Sub MostrarFormulario()
    oFormulario.Opcion = mi_Opcion
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.idProducto = ml_IdProducto
    oFormulario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oFormulario.lcNombrePc = mo_lcNombrePc
    oFormulario.IdPlanCatalogo = ml_IdPlanCatalogo
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub
