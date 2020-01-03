VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BuscaPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para buscar Paquete
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_PaquetesBuscar As New PaquetesBuscar
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_IdFactPaquete As Long
Dim mi_DebeConsiderarPaquete As sghTipoPaquetes

Sub MostrarFormulario()
    mo_PaquetesBuscar.DebeConsiderarPaquete = mi_DebeConsiderarPaquete
    mo_PaquetesBuscar.Show 1
    ml_IdFactPaquete = mo_PaquetesBuscar.IdFactPaquete
    mi_BotonPresionado = mo_PaquetesBuscar.BotonPresionado
    Set mo_PaquetesBuscar = Nothing
End Sub

Property Let IdFactPaquete(lValue As Long)
    ml_IdFactPaquete = lValue
End Property
Property Get IdFactPaquete() As Long
    IdFactPaquete = ml_IdFactPaquete
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property
Property Let DebeConsiderarPaquete(lValue As sghTipoPaquetes)
    mi_DebeConsiderarPaquete = lValue
End Property
