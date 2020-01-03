VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormBoletaRpt 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8790
      Top             =   1950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CrvBoleta 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FormBoletaRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CrvBoleta_PrintButtonClicked(UseDefault As Boolean)
On Error GoTo error111
    ' UseDefault = False
'    With CommonDialog1
'        .Flags = cdlPDPrintSetup Or cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDUseDevModeCopies
'        .CancelError = True
'        .ShowPrinter
'
'    End With
    'CrvBoleta.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    'CrvBoleta.ShowWhatsThis
    'CrvBoleta.EnableSelectExpertButton = True
Exit Sub
error111:
    If Err.Number = 32755 Then
        MsgBox "Has seleccionado Cancelar"
    End If
End Sub

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Formato de Boleta
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Private Sub Form_Load()
'    CrvBoleta.ReportSource = crReport
'    CrvBoleta.ViewReport

End Sub

Public Function leerNombreCampo(cadenaBusqueda As String, arrayNombreCampo() As String)
    Dim cadenaABuscar As Variant
    Dim nombreCampo As String
    Dim posicionEncontro As Integer
    
    For Each cadenaABuscar In arrayNombreCampo
        If cadenaABuscar <> "" Then
            posicionEncontro = InStr(cadenaBusqueda, cadenaABuscar)
            If posicionEncontro > 0 Then
                nombreCampo = cadenaABuscar
                Exit For
            End If
        End If
    Next
    leerNombreCampo = nombreCampo
End Function

Public Function obtenerNombresCampos(nombresCampos() As String)
'    Dim nombresCampos(21) As String
    
    nombresCampos(0) = "BoletaNumeroSerie"
    nombresCampos(1) = "BoletaEstado"
    nombresCampos(2) = "BoletaTipo"
    nombresCampos(3) = "RazonSocial"
    nombresCampos(4) = "FechaCobranza"
    nombresCampos(5) = "Servicio"
    nombresCampos(6) = "Observaciones"
    nombresCampos(7) = "Cajero"
    nombresCampos(8) = "nombreCaja"
    nombresCampos(9) = "Adelantos"
    nombresCampos(10) = "TotalPorPagar"
    nombresCampos(11) = "idCuentaAtencion"
    nombresCampos(12) = "Exoneraciones"
    nombresCampos(13) = "TotalEnLetras"
    nombresCampos(14) = "TotalBoletaPorPagar"
    nombresCampos(15) = "SubTotal"
    nombresCampos(16) = "IGV"
    
    nombresCampos(17) = "Codigo"
    nombresCampos(18) = "NombreProducto"
    nombresCampos(19) = "Cantidad"
    nombresCampos(20) = "PrecionUnitario"
    nombresCampos(21) = "TotalPorPagar"
    
    obtenerNombresCampos = nombresCampos
End Function

Private Sub Form_Resize()
    CrvBoleta.Left = 0
    CrvBoleta.Top = 0
    CrvBoleta.Height = ScaleHeight
    CrvBoleta.Width = ScaleWidth
End Sub
