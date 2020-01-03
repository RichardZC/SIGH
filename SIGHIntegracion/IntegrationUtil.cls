VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntegrationUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Metodos utilitarios para la integracion con otros sistemas
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Public Function CreateNodeXml(sTagNode As String, sValueNode As String, oDocXml As MSXML.DOMDocument) As MSXML.IXMLDOMNode
    Dim oNodeXml As MSXML.IXMLDOMNode
'    Set oNodeXml = oDocXml.createNode()
End Function



Public Function CreateElementXml(sTagElement As String, sValueElement As String, _
            oDocXml As MSXML.DOMDocument, Optional oElementParent As MSXML.IXMLDOMElement = Nothing) As MSXML.IXMLDOMElement
            
    Dim oElement As MSXML.IXMLDOMElement
    Set oElement = oDocXml.createElement(sTagElement)
    oElement.Text = sValueElement
    
    If Not (oElementParent Is Nothing) Then
        oElementParent.appendChild oElement
    End If
    
    Set CreateElementXml = oElement
End Function



Public Function ExistsPathFileRepositorySource(sFileName As String) As Boolean
    ExistsPathFileRepositorySource = False
    If Dir$(sFileName, vbDirectory) = "" Then
        MsgBox "Directorio para almacenamiento de Tramas: " & sFileName & ", no existe", vbInformation, "Integración"
        Exit Function
    End If
    ExistsPathFileRepositorySource = True
End Function

