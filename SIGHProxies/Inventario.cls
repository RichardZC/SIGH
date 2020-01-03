VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Mantenimiento de Inventarios
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Inventario As New FarmInventario
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdInventario As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormularioInventario()
    mo_Inventario.Opcion = mi_Opcion
    mo_Inventario.idUsuario = ml_idUsuario
    mo_Inventario.IdInventario = ml_IdInventario
    mo_Inventario.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_Inventario.lcNombrePc = mo_lcNombrePc
    mo_Inventario.Show 1
    Set mo_Inventario = Nothing
End Sub


Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property


Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdInventario(lValue As Long)
   ml_IdInventario = lValue
End Property
