VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MantenimientoHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_FormularioHIS As New frmMantenimientoHIS
Dim mi_Opcion As sghOpciones
Dim ml_IdUsuario As Long
Dim ml_IdEstablecimiento As Long
Dim ml_IdHisCabecera As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property

Property Get IdRegistroHIS() As Long
    IdRegistroHIS = ml_IdHisCabecera
End Property

Property Let IdRegistroHIS(lValue As Long)
   ml_IdHisCabecera = lValue
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property

Sub MostrarFormulario()
    mo_FormularioHIS.Opcion = mi_Opcion
    mo_FormularioHIS.IdUsuario = ml_IdUsuario
    mo_FormularioHIS.IdEstablecimiento = ml_IdEstablecimiento
    mo_FormularioHIS.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    mo_FormularioHIS.lcNombrePc = mo_lcNombrePc
    mo_FormularioHIS.IdHisCabecera = ml_IdHisCabecera
    mo_FormularioHIS.Show 1
    mi_BotonPresionado = mo_FormularioHIS.BotonPresionado
End Sub
