VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tomog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Tomografías
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_IdMovimiento As New ImagTomog
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim ml_IdMovimiento As Long
Dim ml_IdTipoVentaSeleccionada As Long
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim ml_SeEligioGridBoleta As Boolean

Property Let SeEligioGridBoleta(lValue As Boolean)
    ml_SeEligioGridBoleta = lValue
End Property
Property Get SeEligioGridBoleta() As Boolean
    SeEligioGridBoleta = ml_SeEligioGridBoleta
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Sub MostrarFormulario()
    mo_IdMovimiento.Opcion = mi_Opcion
    mo_IdMovimiento.idUsuario = ml_idUsuario
    mo_IdMovimiento.IdMovimiento = ml_IdMovimiento
    mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
    mo_IdMovimiento.SeEligioGridBoleta = ml_SeEligioGridBoleta
    mo_IdMovimiento.Show 1
    Set mo_IdMovimiento = Nothing
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
Property Let IdMovimiento(lValue As Long)
   ml_IdMovimiento = lValue
End Property
