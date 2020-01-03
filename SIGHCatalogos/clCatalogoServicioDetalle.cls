VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCatalogoServicioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: clase para el detalle de Servicios CPT
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oFormulario As New CatalogoServiciosDetalle
Dim ml_idUsuario As Long
Dim mi_Opcion As sghOpciones
Dim ml_IdProducto As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_TipoCatalogo As Long

Property Let TipoCatalogo(lValue As Long)
    ml_TipoCatalogo = lValue
End Property
Property Get TipoCatalogo() As Long
    TipoCatalogo = ml_TipoCatalogo
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
    oFormulario.TipoCatalogo = ml_TipoCatalogo
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

