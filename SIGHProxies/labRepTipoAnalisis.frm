VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form labRepTipoAnalisis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Pruebas Registradas por Grupos"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "labRepTipoAnalisis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   15
      TabIndex        =   10
      Top             =   1890
      Width           =   7665
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "labRepTipoAnalisis.frx":0CCA
         DownPicture     =   "labRepTipoAnalisis.frx":118E
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
         Left            =   3953
         Picture         =   "labRepTipoAnalisis.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "labRepTipoAnalisis.frx":1B66
         DownPicture     =   "labRepTipoAnalisis.frx":1FC6
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
         Left            =   2453
         Picture         =   "labRepTipoAnalisis.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   7665
      Begin Threed.SSOption optTodos 
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   1140
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   397
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Todos"
         Value           =   -1
      End
      Begin VB.ComboBox cmbGrupo 
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
         ItemData        =   "labRepTipoAnalisis.frx":28B0
         Left            =   1290
         List            =   "labRepTipoAnalisis.frx":28BA
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   210
         Width           =   6330
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   570
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   5115
         TabIndex        =   3
         Top             =   540
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2730
         TabIndex        =   2
         Top             =   570
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   6495
         TabIndex        =   4
         Top             =   540
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##:##"
         PromptChar      =   " "
      End
      Begin Threed.SSOption optHospitalizacion 
         Height          =   225
         Left            =   1110
         TabIndex        =   13
         Top             =   1140
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   397
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Solo Hospitalización"
      End
      Begin Threed.SSOption optSoloEmergencia 
         Height          =   210
         Left            =   3270
         TabIndex        =   14
         Top             =   1140
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   370
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Solo Emergencia"
      End
      Begin Threed.SSOption optSoloCE 
         Height          =   225
         Left            =   5100
         TabIndex        =   15
         Top             =   1140
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   397
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Solo Consultorios Externos"
      End
      Begin Threed.SSOption optExternos 
         Height          =   225
         Left            =   1110
         TabIndex        =   16
         Top             =   1485
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   397
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Externos"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
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
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         Left            =   60
         TabIndex        =   9
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   4605
         TabIndex        =   8
         Top             =   570
         Width           =   435
      End
   End
End
Attribute VB_Name = "labRepTipoAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Tipo de análisis
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbGrupo As New sighentidades.ListaDespleglable
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario  As Long
Dim FI As String, FF As String, HI As String, HF As String
Dim rsReporte As New Recordset
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Private Sub ConfiguraFecha()
  If Len(Trim(txtFdesde.Text)) < 10 Or Len(Trim(txtFhasta.Text)) < 10 Then Exit Sub
  If txtHrInicio.Text <> "" Then
    HI = " " & txtHrInicio.Text '& ":00"
  Else
    HI = " 00:00:00"
  End If
  If txtHrFin.Text <> "" Then
    HF = " " & txtHrFin.Text '& ":59"
  Else
    HF = " 23:59:59"
  End If
  FI = txtFdesde.Text & HI
  FF = txtFhasta.Text & HF
End Sub

Private Function Verifica() As Boolean
  Verifica = False
  If cmbGrupo.Text = "" Or Not (IsDate(txtFdesde.Text)) Or Not (IsDate(txtFhasta.Text)) Then
    Verifica = False
  Else
    Verifica = True
  End If
End Function

Private Sub btnAceptar_Click()

If wxFranklin = "*" Then Exit Sub

 If cmbGrupo.Text = "" Then
    MsgBox "Debe escoger un grupo.", vbInformation, "SIGH "
    cmbGrupo.SetFocus
    Exit Sub
  End If
  ConfiguraFecha
  If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
  End If
  
  'Set rsReporte = mo_ReglasLaboratorio.SacarPruebasPorGrupo(Left(Right(cmbGrupo.Text, 4), 3))
  
  Dim iFila As Long, iCol As Integer
  Dim II As Integer
  Dim TCant As Long, TCant1 As Long
  Dim lbEsOpenOffice As Boolean
  Dim mo_ReporteUtil As New ReporteUtil
  Dim lcBuscaParametro As New SIGHDatos.Parametros
  Dim lcSql As String
  Dim lnHwnd As Long
  Dim lnIdTipoServicio As Long, lcTitulo99 As String
  lnHwnd = Me.hwnd
  
  lnIdTipoServicio = IIf(Me.optTodos.Value, 0, _
                   IIf(Me.optHospitalizacion.Value, 3, _
                   IIf(Me.optSoloCE.Value, 1, _
                   IIf(Me.optSoloEmergencia.Value, 2, 99))))
  lcTitulo99 = "   (" & IIf(Me.optTodos.Value, "TODOS", _
                   IIf(Me.optHospitalizacion.Value, "SOLO HOSPITALIZACION", _
                   IIf(Me.optSoloCE.Value, "SOLO CONSULTAS EXTERNAS", _
                   IIf(Me.optSoloEmergencia.Value, "SOLO EMERGENCIA", "EXTERNOS")))) & ")"
  
  lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
  
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
  MousePointer = 11
  
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\LabRepPruebasRegistradas.ods"
        
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
        Set oExcel = GalenhosExcelApplication()
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LabRepPruebasRegistradas.xls")
        oWorkBookPlantilla.Worksheets("PruebasRegistradas").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        '------- Pruebas por Grupo
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 1).setFormula("GRUPO " & UCase(cmbGrupo.Text))
        Call Feuille.getcellbyposition(1, 2).setFormula("Fecha Inicio " & FI & "  -  Fecha Fin: " & FF & lcTitulo99)
    Else
        'Inicio de Impresion
        oWorkSheet.Cells(2, 2).Value = "GRUPO " & UCase(cmbGrupo.Text)
        oWorkSheet.Cells(3, 2).Value = "Fecha Inicio " & FI & "  -  Fecha Fin: " & FF & lcTitulo99
    End If
  iFila = 6
  iCol = 2
  If rsReporte.State = adStateOpen And Not (rsReporte.EOF = True And rsReporte.BOF = True) Then
  rsReporte.MoveFirst
  II = 0: TCant = 0
  Do While Not rsReporte.EOF
    II = II + 1
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(Trim(rsReporte!codigoCPT))
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(IIf(IsNull(rsReporte!Nombre), "", rsReporte!Nombre))
    Else
        oWorkSheet.Cells(iFila, iCol).Value = II
        oWorkSheet.Cells(iFila, iCol + 1).Value = Trim(rsReporte!codigoCPT)
        oWorkSheet.Cells(iFila, iCol + 2).Value = rsReporte!Nombre
    End If
    TCant1 = mo_ReglasLaboratorio.CuentaPruebas(rsReporte!idProducto, FI, FF, lnIdTipoServicio)
    If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCant1)
    Else
        oWorkSheet.Cells(iFila, iCol + 3).Value = TCant1
    End If
    TCant = TCant + TCant1
    rsReporte.MoveNext
    iFila = iFila + 1
  Loop
  If lbEsOpenOffice = True Then
  Else
    oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 5)).Borders.LineStyle = 1
  End If
  iFila = iFila + 1
    If lbEsOpenOffice = True Then
      Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
      Call Feuille.getcellbyposition(4, iFila - 1).setFormula(TCant)
    Else
      oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
      oWorkSheet.Cells(iFila, 5).Value = TCant ' "=suma(E6:E" & iFila - 2 & ")"
    End If
    If lbEsOpenOffice = True Then
        Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":E" & CStr(iFila))
        mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
    Else
        oWorkSheet.Range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 5)).Borders.LineStyle = 1
    End If
  End If
  If lbEsOpenOffice = True Then
    Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
    PrintArea(0).Sheet = 0
    PrintArea(0).startcolumn = 1
    PrintArea(0).StartRow = 0
    PrintArea(0).EndColumn = 5
    PrintArea(0).EndRow = iFila
    Call Feuille.SetPrintAreas(PrintArea())
    Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
    Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
    MsgBox "El reporte se generó en forma exitosa:" & lcArchivoExcel, vbInformation, "Reporte"
  Else
  oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
    If oWorkSheet.PageSetup.PrintArea <> "" Then
        oWorkSheet.PageSetup.PrintArea = "$A$1:$F$" & (iFila + 2)
    End If
  oExcel.Visible = True
  oWorkSheet.PrintPreview
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
    'encabezado de pagina
    Set PageStyles = Nothing
    Set Sheet = Nothing
    Set StyleFamilies = Nothing
    Set DefPage = Nothing
    Set Htext = Nothing
    Set Hcontent = Nothing
  Else
    'Liberar memoria
    Set oExcel = Nothing
    Set oWorkBookPlantilla = Nothing
    Set oWorkBook = Nothing
    Set oWorkSheet = Nothing
  End If
  MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
  Me.Visible = False
  LimpiarVariablesDeMemoria
  Unload Me
End Sub

Private Sub cmbGrupo_Click()
  If cmbGrupo.Text = "" Then Exit Sub
  Dim rs As New ADODB.Recordset
  Set rsReporte = mo_ReglasLaboratorio.SacarPruebasPorGrupo(Left(Right(cmbGrupo.Text, 4), 3))
  Exit Sub
  
  rs.MoveFirst
  Do While Not rs.EOF
    
    rs.MoveNext
  Loop
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Initialize()
  Set mo_cmbGrupo.MiComboBox = cmbGrupo
End Sub

Private Sub Form_Load()
  txtFdesde.Text = Date
  txtFhasta.Text = Date
  txtHrInicio.Text = "07:00:00"
  txtHrFin.Text = "18:59:59"
  
  mo_cmbGrupo.BoundColumn = "idGrupo"
  mo_cmbGrupo.ListField = "Nombre"
  Set mo_cmbGrupo.RowSource = mo_ReglasLaboratorio.SacarGruposLaboratorio()
End Sub

Private Sub txtFdesde_Change()
  If Verifica = False Then Exit Sub
  cmbGrupo_Click
End Sub



Private Sub txtFdesde_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFdesde_LostFocus()
  If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
      MsgBox "La Fecha Inicial ingresada no es válida", vbInformation, "SIGH "
      txtFdesde.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
      txtFdesde.SetFocus
    End If
  End If
End Sub

Private Sub txtFhasta_Change()
  If Verifica = False Then Exit Sub
  cmbGrupo_Click
End Sub



Private Sub txtFhasta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFhasta_LostFocus()
  If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
      MsgBox "La Fecha Final ingresada no es válida", vbInformation, "SIGH "
      txtFhasta.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
      txtFhasta.SetFocus
    End If
  End If
End Sub

Sub LimpiarVariablesDeMemoria()
  On Error Resume Next
  Set mo_ReglasLaboratorio = Nothing
  Set mo_Teclado = Nothing
  Set mo_cmbGrupo = Nothing
  Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_Change()
  If Verifica = False Then Exit Sub
  cmbGrupo_Click
End Sub


Private Sub txtHrFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrFin_LostFocus()
  If txtHrFin.Text <> "__:__:__" Then
    If Not IsDate(txtHrFin.Text) Then
      MsgBox "La Hora Final ingresada no es válida.", vbInformation, "SIGH "
      txtHrFin.Text = "18:59:59"
      txtHrFin.SetFocus
    End If
  End If
End Sub

Private Sub txtHrInicio_Change()
  If Verifica = False Then Exit Sub
  cmbGrupo_Click
End Sub



Private Sub txtHrInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrInicio_LostFocus()
  If txtHrInicio.Text <> "__:__:__" Then
    If Not IsDate(txtHrInicio.Text) Then
      MsgBox "La Hora Inicial ingresada no es válida.", vbInformation, "SIGH "
      txtHrInicio.Text = "07:00:00"
      txtHrInicio.SetFocus
    End If
  End If
End Sub
