VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clFarmHistPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para mantenimiento de Historicos de Precios
'        Programado por: Barrantes D
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_FarmHistPrecio As New FarmHistPreciosDetalle
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdFarmHistPrecio As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
    mo_FarmHistPrecio.Opcion = mi_Opcion
    mo_FarmHistPrecio.idUsuario = ml_idUsuario
    mo_FarmHistPrecio.IdHistPrecio = ml_IdFarmHistPrecio
    mo_FarmHistPrecio.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_FarmHistPrecio.lcNombrePc = mo_lcNombrePc
    mo_FarmHistPrecio.Show 1
    Set mo_FarmHistPrecio = Nothing
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
Property Let IdFarmHistPrecio(lValue As Long)
   ml_IdFarmHistPrecio = lValue
End Property
