VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmResultadoXitems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label2"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   Icon            =   "frmResultadoXitems.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6585
      Left            =   60
      TabIndex        =   10
      Top             =   1710
      Width           =   13215
      Begin VB.ComboBox cmbCombo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   510
         Visible         =   0   'False
         Width           =   1095
      End
      Begin UltraGrid.SSUltraGrid grdResultados 
         Height          =   6045
         Left            =   60
         TabIndex        =   12
         Top             =   480
         Width           =   13140
         _ExtentX        =   23178
         _ExtentY        =   10663
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BorderStyle     =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "grdResultados"
      End
      Begin VB.Label lblTitulo3 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   8910
         TabIndex        =   15
         Top             =   150
         Width           =   4035
      End
      Begin VB.Label lblTitulo2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   14
         Top             =   150
         Width           =   5050
      End
      Begin VB.Label lblTitulo1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   3765
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   6045
      Begin VB.TextBox txtMedico 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   17
         Top             =   930
         Width           =   4395
      End
      Begin VB.ComboBox cmbResponsable 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   150
         Width           =   4410
      End
      Begin MSMask.MaskEdBox txtFresultado 
         Height          =   315
         Left            =   1590
         TabIndex        =   7
         Top             =   555
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMedicoSolicita 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Médico solicitante"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   16
         Top             =   960
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Realiza Prueba"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   9
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F.Resultado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   105
         TabIndex        =   8
         Top             =   600
         Width           =   945
      End
   End
   Begin VB.Frame fraBoton 
      ForeColor       =   &H00000000&
      Height          =   960
      Left            =   60
      TabIndex        =   3
      Top             =   8340
      Width           =   13215
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmResultadoXitems.frx":0CCA
         DownPicture     =   "frmResultadoXitems.frx":118E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   6765
         Picture         =   "frmResultadoXitems.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprime (F3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   60
         Picture         =   "frmResultadoXitems.frx":1B66
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1365
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmResultadoXitems.frx":203F
         DownPicture     =   "frmResultadoXitems.frx":249F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5205
         Picture         =   "frmResultadoXitems.frx":2914
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
   End
   Begin SIGHLaboratorio.UcPacienteDatos1 UcPacienteDatos1 
      Height          =   1695
      Left            =   6120
      TabIndex        =   4
      Top             =   30
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   2990
   End
End
Attribute VB_Name = "frmResultadoXitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Resultados de varios Items
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const PM_POSITIONCTRL = 1
Private Const PM_MOVEPREVCELL = 2
Private Const PM_MOVENEXTCELL = 3
Private Const PM_EXITEDITMODE = 4
Private Const PM_PROCESSKEY = 5

Private Const VK_TAB = &H9
Private Const VK_SHIFT = &H10
Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_CONTROL = &H11
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const lnColorInabilitado = &HF9EADF

Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim oRsResultados As New Recordset
Dim oRsResultadosCPT As New Recordset
Dim lbGrabo As Boolean
Dim lnIdGrupo As Long
Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_idProductoCpt As Long
Dim ml_NoMuestraBotonGrabar As Boolean
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ml_idTipoSexo As Long
Dim ml_FechaNacimiento As Date
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim lcServicioActualPaciente As String
Dim lnIdPaciente99 As Long
Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let FechaNacimiento(lValue As Date)
    ml_FechaNacimiento = lValue
End Property
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let NoMuestraBotonGrabar(lValue As Boolean)
   ml_NoMuestraBotonGrabar = lValue
   If ml_NoMuestraBotonGrabar = True Then
      cmdGrabar.Visible = False
   End If
End Property

Property Let idProductoCPT(lValue As Long)
   ml_idProductoCpt = lValue
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property





Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    'Case vbKeyReturn
     ' SendKeys "{TAB}"
    Case vbKeyF3
      cmdImprimir_Click
    Case vbKeyEscape
      cmdCancelar_Click
    Case vbKeyF2
      cmdGrabar_Click
  End Select
End Sub







Private Sub cmbCombo_Validate(Cancel As Boolean)
    grdResultados.ActiveCell.Value = cmbCombo.Text
End Sub



Private Sub cmbResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbResponsable
   AdministrarKeyPreview KeyCode
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Function ValidaDatosObligatorios() As Boolean
  ValidaDatosObligatorios = False
  If cmbResponsable.Text = "" Then
    MsgBox "Debe Seleccionar el personal que realizó la prueba", vbInformation, "SIGH "
    cmbResponsable.SetFocus
    Exit Function
  End If
  If Me.txtFresultado.Text = sighentidades.FECHA_VACIA_DMY Then
    MsgBox "Por favor ingresar la Fecha del Resultado", vbInformation, "SIGH "
    Exit Function
  End If
  ValidaDatosObligatorios = True
End Function

Private Sub cmdGrabar_Click()
  lbGrabo = False
  If ValidaDatosObligatorios Then
    If mo_ReglasLaboratorio.LabResultadosPorItemsActualizar(ml_idProductoCpt, ml_idOrden, oRsResultados, _
                                 Val(mo_cmbResponsable.BoundText), ml_idUsuario, txtFresultado.Text, _
                                 lnIdPaciente99) = True Then
       lbGrabo = True
       'debb2014d
       MsgBox "Se Guardó correctamente los resultados Personalizados", vbExclamation, Me.Caption
    End If
  End If
End Sub

'debb-28/03/2016
Private Sub cmdImprimir_Click()
Dim lcTexto As String
Dim lnHwnd As Long
Dim lcEdadEnAtencion As String
lnHwnd = Me.hwnd
    If ValidaDatosObligatorios Then
        If lbGrabo = False Then
           MsgBox "Tiene que GRABAR primero", vbInformation, Me.Caption
           Exit Sub
        End If
        Dim iFila As Long, iColumna As Integer
        Dim lbEsOpenOffice As Boolean
 
        lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
        On Error GoTo ManejadorError
    
    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
        
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\LabResultadoXitem.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LabResultadoXitem.xls")
        oWorkBookPlantilla.Worksheets("Hoja1").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    '
    lcEdadEnAtencion = ""
    If Me.UcPacienteDatos1.DevuelveFechaNacimiento <> sighentidades.DevuelveFechaSoloFormato_DMY Then
       Dim oEdad As Edad
       oEdad = sighentidades.CalcularEdad(CDate(UcPacienteDatos1.DevuelveFechaNacimiento), CDate(txtFresultado.Text))
       lcEdadEnAtencion = oEdad.Edad & " " & oEdad.NombreEdad
    End If
    '
    
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(3, 0).setFormula(Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre & _
                                           " (Sexo: " & Left(Me.UcPacienteDatos1.DevuelveSexo, 2) & ") (" & _
                                           lcEdadEnAtencion & ") " & lcServicioActualPaciente)
            Call Feuille.getcellbyposition(3, 1).setFormula(Me.txtMedico.Text)
            Call Feuille.getcellbyposition(6, 1).setFormula("'" & Me.txtFresultado.Text)
            Call Feuille.getcellbyposition(3, 2).setFormula(Me.cmbResponsable.Text)
            Call Feuille.getcellbyposition(1, 4).setFormula("ANALISIS: " & Me.Caption)
        Else
            oWorkSheet.Cells(1, 4).Value = Me.UcPacienteDatos1.DevuelveHistoriaApellidosYnombre & _
                                           " (Sexo: " & Left(Me.UcPacienteDatos1.DevuelveSexo, 2) & ") (" & _
                                           lcEdadEnAtencion & ") " & lcServicioActualPaciente
            oWorkSheet.Cells(2, 4).Value = Me.txtMedico
            oWorkSheet.Cells(2, 7).Value = "'" & Me.txtFresultado.Text
            oWorkSheet.Cells(3, 4).Value = Me.cmbResponsable.Text
            oWorkSheet.Cells(5, 2).Value = "ANALISIS: " & Me.Caption
        End If
        iFila = 9
        If oRsResultados.RecordCount > 0 Then
           oRsResultados.MoveFirst
           Do While Not oRsResultados.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(CStr(oRsResultados.Fields!Grupo))
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(CStr(oRsResultados.Fields!Item))
                Else
                  oWorkSheet.Cells(iFila, 2).Value = oRsResultados.Fields!Grupo
                  oWorkSheet.Cells(iFila, 3).Value = oRsResultados.Fields!Item
                End If
                lcTexto = ""
                If oRsResultados.Fields!ValorNumero > 0 Then
                  lcTexto = lcTexto & Trim(Str(IIf(IsNull(oRsResultados.Fields!ValorNumero), "", oRsResultados.Fields!ValorNumero))) & "| "
                End If
                If Len(Trim(oRsResultados.Fields!ValorTexto)) > 0 Then
                  lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!ValorTexto), "", oRsResultados.Fields!ValorTexto)) & "| "
                End If
                If Len(Trim(oRsResultados.Fields!ValorCombo)) > 0 Then
                  lcTexto = lcTexto & Trim(IIf(IsNull(oRsResultados.Fields!ValorCombo), "", oRsResultados.Fields!ValorCombo)) & "| "
                End If
                If Not IsNull(oRsResultados.Fields!ValorCheck) Then
                  lcTexto = lcTexto & IIf(oRsResultados.Fields!ValorCheck = True, "x", "")
                End If
                
                lcTexto = Trim(lcTexto)
                If lcTexto <> "" Then
                    If Right(lcTexto, 1) = "|" Then
                        lcTexto = Left(lcTexto, Len(lcTexto) - 1)
                    End If
                End If
                
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcTexto)
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!ValorReferencial), "", oRsResultados.Fields!ValorReferencial))
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(IsNull(oRsResultados.Fields!Metodo), "", oRsResultados.Fields!Metodo))
                Else
                    oWorkSheet.Cells(iFila, 4).Value = lcTexto
            
'              If oRsResultados.Fields!ValorNumero <> 0 Then
'                    oWorkSheet.Cells(iFila, 4).Value = oRsResultados.Fields!ValorNumero
'              End If
'                    oWorkSheet.Cells(iFila, 5).Value = oRsResultados.Fields!ValorTexto
'                    oWorkSheet.Cells(iFila, 6).Value = oRsResultados.Fields!ValorCombo
'                    oWorkSheet.Cells(iFila, 7).Value = oRsResultados.Fields!ValorCheck

                    oWorkSheet.Cells(iFila, 8).Value = oRsResultados.Fields!ValorReferencial
                    oWorkSheet.Cells(iFila, 9).Value = oRsResultados.Fields!Metodo
                End If
              iFila = iFila + 1
              oRsResultados.MoveNext
           Loop
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 9
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.Cells(iFila, 2).Value = "Digitador: " & InicialesDelDigitador
        
            oWorkSheet.Range(oWorkSheet.Cells(iFila, 3), oWorkSheet.Cells(iFila + 2, 100)).Select
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = sighentidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
            oExcel.Visible = True
            
            If Val(lcBuscaParametro.SeleccionaFilaParametro(208)) <> 7637 Then                               'huaral
                    oWorkSheet.PrintPreview
            Else
               ' oWorkSheet.PrintPreview
            End If
        End If
        If lbEsOpenOffice = True Then
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            Set Style = Nothing
            Set Border = Nothing
        Else
            'Liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
    End If
    Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
End Sub

Function InicialesDelDigitador() As String
    Dim oConexion As New Connection
    Dim oReglasCaja As New ReglasCaja
    oConexion.CommandTimeout = 900
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    InicialesDelDigitador = oReglasCaja.SeleccionaDatosCajeroConexion(sighentidades.Usuario, sghIniciales, oConexion)
    oConexion.Close
    Set oConexion = Nothing
    Set oReglasCaja = Nothing
End Function

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
  Me.txtFresultado.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
  CreaTemporal
  
  TemporalLlenarItems
  
End Sub

Sub CreaTemporal()
    With oRsResultados
          .Fields.Append "SoloNumero", adBoolean
          .Fields.Append "SoloTexto", adBoolean
          .Fields.Append "SoloCombo", adBoolean
          .Fields.Append "SoloCheck", adBoolean
          .Fields.Append "ordenXresultado", adInteger, 4, adFldIsNullable
          .Fields.Append "Grupo", adVarChar, 100, adFldIsNullable
          .Fields.Append "Item", adVarChar, 100, adFldIsNullable
          .Fields.Append "idItem", adInteger
          .Fields.Append "ValorNumero", adDouble
          .Fields.Append "ValorTexto", adVarChar, 500, adFldIsNullable
          .Fields.Append "ValorCombo", adVarChar, 100, adFldIsNullable
          .Fields.Append "ValorCheck", adVarChar, 1, adFldIsNullable
          .Fields.Append "ValorReferencial", adVarChar, 100, adFldIsNullable
          .Fields.Append "Metodo", adVarChar, 50, adFldIsNullable
          .Fields.Append "idGrupo", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
    Set Me.grdResultados.DataSource = oRsResultados
    grdResultados.Bands(0).Columns("ordenXresultado").Hidden = True
    grdResultados.Bands(0).Columns("idItem").Hidden = True
    grdResultados.Bands(0).Columns("SoloNumero").Hidden = True
    grdResultados.Bands(0).Columns("SoloTexto").Hidden = True
    grdResultados.Bands(0).Columns("SoloCombo").Hidden = True
    grdResultados.Bands(0).Columns("SoloCheck").Hidden = True
    grdResultados.Bands(0).Columns("Grupo").Width = 1800
    grdResultados.Bands(0).Columns("Item").Width = 1700
    grdResultados.Bands(0).Columns("ValorNumero").Width = 1000
    grdResultados.Bands(0).Columns("ValorTexto").Width = 2000
    grdResultados.Bands(0).Columns("ValorCombo").Width = 1000
    grdResultados.Bands(0).Columns("ValorCheck").Width = 1000
    grdResultados.Bands(0).Columns("ValorReferencial").Width = 3000
    grdResultados.Bands(0).Columns("Metodo").Width = 1000
    grdResultados.Bands(0).Columns("ValorTexto").CellMultiLine = ssCellMultiLineTrue
    grdResultados.Bands(0).Columns("ValorCheck").Style = ssStyleCheckBox
    grdResultados.Bands(0).Columns("idGrupo").Hidden = True
    'mo_Apariencia.ConfigurarFilasBiColores grdResultados, SIGHEntidades.GrillaConFilasBicolor
End Sub

Sub TemporalLlenarItems()
  Dim oReglasFacturacion As New SIGHNegocios.ReglasFacturacion
  Dim oRsTmp1 As New Recordset, lnIdItem As Long, lcPaciente As String
  'Llena Temporal
  Set oRsResultadosCPT = mo_ReglasLaboratorio.LabItemsCptSeleccionarXfiltro("dbo.LabItemsCpt.idProductoCpt=" & ml_idProductoCpt)
  lbGrabo = False
  If oRsResultadosCPT.RecordCount > 0 Then
     oRsResultadosCPT.MoveFirst
     lnIdGrupo = oRsResultadosCPT.Fields!idGrupo
     Me.Caption = oRsResultadosCPT.Fields!Codigo & " - " & oRsResultadosCPT.Fields!Nombre
     Do While Not oRsResultadosCPT.EOF
        lnIdItem = oRsResultadosCPT.Fields!idItem
        oRsResultados.AddNew
        oRsResultados.Fields!ordenXresultado = oRsResultadosCPT.Fields!ordenXresultado
        oRsResultados.Fields!Grupo = oRsResultadosCPT.Fields!Grupo
        oRsResultados.Fields!Item = oRsResultadosCPT.Fields!Item
        oRsResultados.Fields!idItem = oRsResultadosCPT.Fields!idItem
        oRsResultados.Fields!ValorReferencial = oRsResultadosCPT.Fields!ValorReferencial
        oRsResultados.Fields!Metodo = oRsResultadosCPT.Fields!Metodo
        oRsResultados.Fields!SoloNumero = IIf(oRsResultadosCPT.Fields!SoloNumero = True, True, False)
        oRsResultados.Fields!Solotexto = IIf(oRsResultadosCPT.Fields!Solotexto = True, True, False)
        oRsResultados.Fields!SoloCombo = IIf(oRsResultadosCPT.Fields!SoloCombo = True, True, False)
        oRsResultados.Fields!SoloCheck = IIf(oRsResultadosCPT.Fields!SoloCheck = True, True, False)
        oRsResultados.Fields!idGrupo = oRsResultadosCPT.Fields!idItemGrupo
        oRsResultados.Update
        Do While Not oRsResultadosCPT.EOF And lnIdItem = oRsResultadosCPT.Fields!idItem
           oRsResultadosCPT.MoveNext
           If oRsResultadosCPT.EOF Then
              Exit Do
           End If
        Loop
     Loop
     'Barre grid e Inhabilita celdas que no se deben editar
     Dim oRow As UltraGrid.SSRow
     Dim oCell As UltraGrid.SSCell
     Dim SpanBands As Boolean, lnAlto As Long
     Dim lbSoloNumero As Boolean, lbSoloTexto As Boolean, lbSoloCombo As Boolean, lbSoloCheck As Boolean
     SpanBands = False
     oRsResultados.MoveFirst
     
     Set oRow = grdResultados.GetRow(ssChildRowFirst)       'gridSearch.ActiveRow
     Do
        For Each oCell In oRow.Cells
            Select Case oCell.Column.Key
            Case "SoloNumero"
                 lbSoloNumero = oCell.Value
            Case "SoloTexto"
                 lbSoloTexto = oCell.Value
            Case "SoloCombo"
                 lbSoloCombo = oCell.Value
            Case "SoloCheck"
                 lbSoloCheck = oCell.Value
            Case "ValorNumero"
                 If lbSoloNumero = False Then
                    oCell.Appearance.BackColor = lnColorInabilitado
                    oCell.Activation = ssActivationActivateNoEdit
                 End If
            Case "ValorTexto"
                 If lbSoloTexto = False Then
                    oCell.Appearance.BackColor = lnColorInabilitado
                    oCell.Activation = ssActivationActivateNoEdit
                 Else
                     'oRow.Appearance.TextAlign = ssAlignCenter
                     'oRow.Appearance.TextVAlign = ssVAlignMiddle
                   ' oCell.DroppedDown = True
                   'lnAlto = 4000
                   ' oRow.Height = 1000
                 End If
            Case "ValorCombo"
                 If lbSoloCombo = False Then
                    oCell.Appearance.BackColor = lnColorInabilitado
                    oCell.Activation = ssActivationActivateNoEdit
                 End If
            Case "ValorCheck"
                 If lbSoloCheck = False Then
                    oCell.Appearance.BackColor = lnColorInabilitado
                    oCell.Activation = ssActivationActivateNoEdit
                 End If
            Case Else
                 oCell.Appearance.BackColor = lnColorInabilitado
                 oCell.Activation = ssActivationActivateNoEdit
                 
            End Select
        Next
        If Not oRow.HasNextSibling(SpanBands) Then Exit Do
        Set oRow = oRow.GetSibling(ssSiblingRowNext, SpanBands)
        If Err.Number <> 0 Then
            Exit Do
        End If
    Loop
  End If
  '
  CargaDataCombos
  'Medico que ordena
  mo_Formulario.HabilitarDeshabilitar Me.txtMedico, False
  Set oRsTmp1 = mo_ReglasLaboratorio.LabMovimientoLaboratorioSeleccionarXidOrden(ml_idOrden)
  lcPaciente = ""
  If oRsTmp1.RecordCount > 0 Then
     ml_idTipoSexo = IIf(IsNull(oRsTmp1.Fields!idTipoSexo), 1, oRsTmp1.Fields!idTipoSexo)
     
     ml_FechaNacimiento = IIf(IsNull(oRsTmp1.Fields!FechaNacimiento), Date, oRsTmp1.Fields!FechaNacimiento)
     txtMedico.Text = oRsTmp1.Fields!OrdenaPrueba
     lcPaciente = oRsTmp1.Fields!Paciente
  End If
  lcServicioActualPaciente = mo_ReglasLaboratorio.DevuelveDatosParaImpresionResultadoLaboratorio(ml_idOrden)
  'Paciente
  lnIdPaciente99 = 0
  Set oRsTmp1 = oReglasFacturacion.FactOrdenServicioSeleccionarXidOrden(ml_idOrden)
  Me.UcPacienteDatos1.FechaRegistro = Now
  Me.UcPacienteDatos1.DeshabilitarFrames True
  If oRsTmp1.Fields!idPaciente > 0 Then
     lnIdPaciente99 = oRsTmp1.Fields!idPaciente
     Me.UcPacienteDatos1.idPaciente = oRsTmp1.Fields!idPaciente
     Me.UcPacienteDatos1.CargarDatosDePacienteALosControles
     'If Not IsNull(oRsTmp1!IdServicioPaciente) Then
     '   lcServicioActualPaciente = " (S.Act: " & mo_ReglasFacturacion.BuscaServicioActualDelPaciente(oRsTmp1!IdServicioPaciente) & ")"
     'End If
  Else
     Me.UcPacienteDatos1.idPaciente = 0
     Me.UcPacienteDatos1.idTipoSexo = ml_idTipoSexo
     If sighentidades.EsFecha(Format(ml_FechaNacimiento, "dd/mm/yyyy"), "DD/MM/AAAA") Then
        Me.UcPacienteDatos1.FechaNacimiento = ml_FechaNacimiento
     End If
     Me.UcPacienteDatos1.CargaAlgunosDatosDesdeBoleta lcPaciente
     
  End If
  'Llena Resultados Ya grabados
  Set oRsTmp1 = mo_ReglasLaboratorio.LabResultadosPorItemsSeleccionarXfiltro("idOrden=" & ml_idOrden & _
                                                                          " and idProductoCpt=" & ml_idProductoCpt)
                                                                          
  If oRsTmp1.RecordCount > 0 Then
     lbGrabo = True
     oRsTmp1.MoveFirst
     mo_cmbResponsable.BoundText = Trim(Str(oRsTmp1.Fields!realizaAnalisis))
     Me.txtFresultado.Text = Format(oRsTmp1.Fields!fecha, sighentidades.DevuelveFechaSoloFormato_DMY_HM)
     Do While Not oRsTmp1.EOF
        oRsResultados.MoveFirst
        oRsResultados.Find "ordenXresultado=" & oRsTmp1.Fields!ordenXresultado
        If Not oRsResultados.EOF Then
            If Not IsNull(oRsTmp1.Fields!ValorNumero) Then
               oRsResultados.Fields!ValorNumero = oRsTmp1.Fields!ValorNumero
            End If
            If Not IsNull(oRsTmp1.Fields!ValorTexto) Then
            oRsResultados.Fields!ValorTexto = oRsTmp1.Fields!ValorTexto
            End If
            If Not IsNull(oRsTmp1.Fields!ValorCombo) Then
               oRsResultados.Fields!ValorCombo = oRsTmp1.Fields!ValorCombo
            End If
            If Not IsNull(oRsTmp1.Fields!ValorCheck) Then
               oRsResultados.Fields!ValorCheck = oRsTmp1.Fields!ValorCheck
            End If
            oRsResultados.Update
        End If
        oRsTmp1.MoveNext
     Loop
  End If
  '
  If oRsResultados.RecordCount > 0 Then
     oRsResultados.MoveFirst
  End If
  Set oReglasFacturacion = Nothing
  Set oRsTmp1 = Nothing
End Sub

Sub CargaDataCombos()
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  Set mo_cmbResponsable.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & _
                                                            "(select dbo.labGrupos.idCargo from dbo.labGrupos " & _
                                                            "where dbo.labGrupos.idGrupo=" & lnIdGrupo & ")")

    If mo_CabeceraReportes.NOpuedeModificarResponsable(sghAgregar, sighentidades.Usuario, mo_cmbResponsable.RowSource) Then
       mo_cmbResponsable.BoundText = Trim(Str(sighentidades.Usuario))
       mo_Formulario.HabilitarDeshabilitar Me.cmbResponsable, False
    End If

End Sub

Sub LlenaComboParaGrid(lnIdItem As Long)
    cmbCombo.Clear
    oRsResultadosCPT.Filter = "idItem=" & lnIdItem
    If oRsResultadosCPT.RecordCount > 0 Then
       oRsResultadosCPT.MoveFirst
       Do While Not oRsResultadosCPT.EOF
          cmbCombo.AddItem oRsResultadosCPT.Fields!ValorSiEsCombo
          oRsResultadosCPT.MoveNext
       Loop
    End If
End Sub
Private Sub HideControl(Ctrl As Control)
    'hide the control if available
    On Error Resume Next
    If Not Ctrl Is Nothing Then
        Ctrl.Visible = False
    End If
End Sub

Private Sub cmbCombo_LostFocus()

'    grdResultados.ActiveCell.Value = cmbCombo.Text
'    HideControl cmbCombo



    Dim lRet As Long
    
    'check if the tab key is pressed
    lRet = GetAsyncKeyState(VK_TAB)
    
    'hide the dtpickers
    HideControl cmbCombo
    
    If lRet < 0 Then
        If GetAsyncKeyState(VK_LCONTROL) < 0 Or GetAsyncKeyState(VK_RCONTROL) < 0 Or _
            GetAsyncKeyState(VK_CONTROL) < 0 Then
            'check for control key first, if held exit
            Exit Sub
        ElseIf GetAsyncKeyState(VK_LSHIFT) < 0 Or GetAsyncKeyState(VK_RSHIFT) < 0 Or _
            GetAsyncKeyState(VK_SHIFT) < 0 Then
            'if shift is held then move to previous cell
            grdResultados.PostMessage PM_MOVEPREVCELL
        Else
            'no ctrl or shift key is pressed so move to next cell
            grdResultados.PostMessage PM_MOVENEXTCELL
        End If
    End If
End Sub

Private Sub grdResultados_AfterCellUpdate(ByVal Cell As UltraGrid.SSCell)
    If Cell.Column.Key = "ValorTexto" Then
       Dim lcTexto As String, lnFor As Integer, lnPos As Integer
       If Not IsNull(Cell.Row.Cells("ValorTexto").Value) Then
            lcTexto = Cell.Row.Cells("ValorTexto").Value
            If Asc(Left(lcTexto, 1)) = 13 Or Asc(Left(lcTexto, 1)) = 10 Then
                 For lnFor = 1 To Len(lcTexto)
                     If Asc(Mid(lcTexto, lnFor, 1)) <> 13 And Asc(Mid(lcTexto, lnFor, 1)) <> 10 Then
                        lnPos = lnFor
                        Exit For
                     End If
                 Next
                 If lnPos > 0 Then
                    Cell.Row.Cells("ValorTexto").Value = Left(Mid(lcTexto, lnPos, 500), 500)
                 End If
            End If
       End If
    End If
End Sub

Private Sub grdResultados_BeforeCellDeactivate(ByVal Cancel As UltraGrid.SSReturnBoolean)
    HideControl cmbCombo
End Sub

Private Sub grdResultados_BeforeColRegionScroll(ByVal NewState As UltraGrid.SSColScrollRegion, ByVal OldState As UltraGrid.SSColScrollRegion, ByVal Cancel As UltraGrid.SSReturnBoolean)
    HideControl cmbCombo
End Sub

Private Sub grdResultados_BeforeEnterEditMode(ByVal Cancel As UltraGrid.SSReturnBoolean)
    On Error GoTo ErrHandler
    If grdResultados.ActiveCell.Column.DataField = "ValorCombo" Then
       Dim lnItem As Long
       lnItem = oRsResultados.Fields!idItem
       LlenaComboParaGrid lnItem
       cmbCombo.Text = grdResultados.ActiveCell.GetText
       grdResultados.PostMessage PM_POSITIONCTRL, cmbCombo
    End If
ErrHandler:
End Sub

Private Sub grdResultados_BeforeRowRegionScroll(ByVal NewState As UltraGrid.SSRowScrollRegion, ByVal OldState As UltraGrid.SSRowScrollRegion, ByVal Cancel As UltraGrid.SSReturnBoolean)
    HideControl cmbCombo
End Sub



Private Sub grdResultados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdResultados.Caption = ""
    grdResultados.Override.RowSizing = ssRowSizingFree   'grdResultados.Override.RowSizing = ssRowSizingFixed
    

End Sub


Private Sub grdResultados_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3, vbKeyF2
         Dim lnKeyCode As Integer
         lnKeyCode = KeyCode
         AdministrarKeyPreview lnKeyCode
    End Select
End Sub

Private Sub grdResultados_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    Select Case KeyAscii
    Case vbKeyReturn
         grdResultados.PerformAction ssKeyActionExitEditMode
         grdResultados.PerformAction ssKeyActionBelowCell
         grdResultados.PerformAction ssKeyActionEnterEditMode
    End Select
End Sub

Private Sub grdResultados_PostMessageReceived(ByVal MsgID As Long, Optional ByVal MsgData1 As Variant, Optional ByVal MsgData2 As Variant)
    Dim Ctl As Control
    
    If Not IsMissing(MsgData1) Then
        'check if there was data passed in and if so, if it is an object
        If IsObject(MsgData1) Then
            'if an object was passed in, then assign it to the
            ' local control variable declared above
            Set Ctl = MsgData1
        End If
    End If
    
    With Me.grdResultados
        Select Case MsgID
            Case PM_POSITIONCTRL
                'position a control over a cell
                PositionOverCell Ctl, grdResultados
            Case PM_MOVEPREVCELL
                'Move to the previous cell. A control positioned
                ' must have lost focus due to a shift-tab
                .SetFocus
                .PerformAction ssKeyActionPrevCellByTab
            Case PM_MOVENEXTCELL
                'Move to the next cell. A control positioned
                ' must have lost focus due to a tab key
                .SetFocus
                .PerformAction ssKeyActionNextCellByTab
            Case PM_EXITEDITMODE
                'The positioned control signal that the user wants
                ' to exit edit mode
                HideControl Ctl
            Case PM_PROCESSKEY
                'A "special" keystroke was pressed and should be
                ' processed by the UltraGrid.
                .SetFocus
                Select Case MsgData1
                    Case vbKeyUp
                        .PerformAction ssKeyActionAboveCell
                    Case vbKeyDown
                        .PerformAction ssKeyActionBelowCell
                    Case vbKeyPageUp
                        .PerformAction ssKeyActionPageUpCell
                    Case vbKeyPageDown
                        .PerformAction ssKeyActionPageDownCell
                End Select
                .PerformAction ssKeyActionEnterEditMode
        End Select
    End With
End Sub
Private Sub PositionOverCell(CtlToMove As Control, Grid As UltraGrid.SSUltraGrid)
    Dim UIElement As UltraGrid.SSUIElement
    Dim iScale As VBRUN.ScaleModeConstants
    Dim Cell As UltraGrid.SSCell
    
    On Error GoTo ErrHandler
    
    'ensure that there are valid objects to work with
    If Grid Is Nothing Then Exit Sub
    If CtlToMove Is Nothing Then Exit Sub
    
    'must have an activecell in order to know where to position
    ' the control
    If Grid.ActiveCell Is Nothing Then Exit Sub
    
    'get a reference to the activecell
    Set Cell = Grid.ActiveCell
    
    'make sure the cell is in view
    Grid.ActiveColScrollRegion.ScrollCellIntoView Cell, Grid.ActiveRowScrollRegion
    
    'get the uielement for the active cell
    Set UIElement = Cell.GetUIElement(Grid.ActiveRowScrollRegion, Grid.ActiveColScrollRegion)
    If UIElement Is Nothing Then
        Exit Sub
    End If
    
    'find out what kind of scaling is needed
    If TypeOf CtlToMove.Parent Is MDIForm Then
        iScale = vbPixels
    Else
        iScale = CtlToMove.Parent.ScaleMode
    End If
    
    'Use the "visible" positioning for the activecell to position the
    ' control. Theoretically, a cell could be larger than the visible
    ' area of the control.
    With UIElement.RectDisplayed
        If .Left < 0 Or .Top < 0 Or .Width <= 0 Or .Height <= 0 Then
            Exit Sub
        End If
        
        'reposition the control
        If iScale = vbPixels Then
            'CtlToMove.Move .Left + Grid.Left, .Top + Grid.Top, .Width, .Height
            CtlToMove.Move .Left + Grid.Left, .Top + Grid.Top
        Else
            CtlToMove.Move CtlToMove.Parent.ScaleX(.Left - 1, vbPixels, iScale) + _
                           Grid.Left, CtlToMove.Parent.ScaleY(.Top - 1, vbPixels, iScale) + _
                           Grid.Top
'            CtlToMove.Move CtlToMove.Parent.ScaleX(.Left - 1, vbPixels, iScale) + _
'                Grid.Left, CtlToMove.Parent.ScaleY(.Top - 1, vbPixels, iScale) + _
'                Grid.Top, CtlToMove.Parent.ScaleX(.Width + 1, vbPixels, iScale), _
'                CtlToMove.Parent.ScaleY(.Height + 2, vbPixels, iScale)
        End If
        
        'show it, bring it to the front, and give it focus
        CtlToMove.Visible = True
        CtlToMove.ZOrder 0
        CtlToMove.SetFocus
    End With
    
ErrHandler:

End Sub




Private Sub txtFresultado_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFresultado
   AdministrarKeyPreview KeyCode
End Sub



Private Sub txtMedico_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtMedico
   AdministrarKeyPreview KeyCode
End Sub
