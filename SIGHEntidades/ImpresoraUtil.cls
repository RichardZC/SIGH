VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImpresoraUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function listaImpresorasInstaladas() As ADODB.Recordset
    Dim rsImpresoras As New ADODB.Recordset
    Dim I As Integer
    Dim impresora As Printer
    
    rsImpresoras.Fields.Append "idPrinter", adSmallInt
    rsImpresoras.Fields.Append "printerName", adVarChar, 200
    rsImpresoras.LockType = adLockPessimistic
    rsImpresoras.Open
    
    For I = 0 To Printers.Count - 1
        rsImpresoras.AddNew
        rsImpresoras!idPrinter = I
        rsImpresoras!printerName = Printers(I).DeviceName
        rsImpresoras.Update
    Next
    Set listaImpresorasInstaladas = rsImpresoras
End Function
'mgaray201411h
Public Function ObtenerMensajeErrorParaImpresion(Err As ErrObject) As String
    Dim sMensage As String
    
    Select Case Err.Number
        Case 1004
            sMensage = "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente."
        Case 8555
            sMensage = "No se puede obtener información de la impresora, seleccione Configuración en el menú Inicio de Windows, haga clic en Impresoras y después ponga una impresora como predeterminada. Siga las instrucciones del asistente."
        Case Else
            If Err.Description = "File not found." Then
                sMensage = "No se puede encontrar las plantillas para la impresión"
            Else
                sMensage = Err.Number & " : " & Err.Description
            End If
    End Select
    ObtenerMensajeErrorParaImpresion = sMensage
End Function

Public Sub ImpresoraPredeterminada(lcNombreImpresora As String)
    Dim obj_Impresora As Object
    Set obj_Impresora = CreateObject("WScript.Network")
    obj_Impresora.setdefaultprinter lcNombreImpresora
    Set obj_Impresora = Nothing
End Sub
