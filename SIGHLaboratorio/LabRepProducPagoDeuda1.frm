VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form LabRepProducPagoDeuda1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productividad de Laboratorio: Consolidado"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "LabRepProducPagoDeuda1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1080
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   7755
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   6960
         Picture         =   "LabRepProducPagoDeuda1.frx":0CCA
         TabIndex        =   10
         Top             =   645
         Visible         =   0   'False
         Width           =   660
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
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
         Left            =   5100
         TabIndex        =   2
         Top             =   210
         Visible         =   0   'False
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
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
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
         Left            =   6480
         TabIndex        =   3
         Top             =   210
         Visible         =   0   'False
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   330
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
         Left            =   4590
         TabIndex        =   9
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F. Movimiento"
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
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   30
      TabIndex        =   5
      Top             =   1125
      Width           =   7740
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LabRepProducPagoDeuda1.frx":0FDC
         DownPicture     =   "LabRepProducPagoDeuda1.frx":143C
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
         Left            =   2370
         Picture         =   "LabRepProducPagoDeuda1.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LabRepProducPagoDeuda1.frx":1D26
         DownPicture     =   "LabRepProducPagoDeuda1.frx":21EA
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
         Left            =   3900
         Picture         =   "LabRepProducPagoDeuda1.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "LabRepProducPagoDeuda1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Productividad
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim rsTmp As New Recordset
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idUsuario(lValue As Long)
  ml_idUsuario = lValue
End Property

Private Sub btnAceptar_Click()
  If ValidaDatosObligatorios = True Then
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporte1 As New ADODB.Recordset
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim mda_FechaInicio As String ' Date
    Dim mda_FechaFin As String 'Date
    Dim mo_ReporteUtil As New ReporteUtil
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lnHwnd As Long
    lnHwnd = Me.hwnd
    Dim lcSql As String
    cmbYear.Enabled = False
    txtFdesde.Enabled = False
    txtFhasta.Enabled = False
    txtHrInicio.Enabled = False
    txtHrFin.Enabled = False
    Me.MousePointer = 11
    mda_FechaInicio = txtFdesde.Text & " " & txtHrInicio.Text
    mda_FechaFin = txtFhasta.Text & " " & txtHrFin.Text
    
    Set rsReporte = mo_ReglasLaboratorio.SacarPruebasTodas()
    If rsReporte.RecordCount > 0 Then
      Dim iFila As Long, iCol As Integer, II As Integer
      Dim TEx As Double, TCE As Double, THosp As Double, TEmer As Double
      Dim TEx1 As Double, TCE1 As Double, THosp1 As Double, TEmer1 As Double
      Dim TEx2 As Double, TCE2 As Double, THosp2 As Double, TEmer2 As Double
      Dim MTEx As Double, MTCE As Double, MTHosp As Double, MTEmer As Double
      Dim MTEx1 As Double, MTCE1 As Double, MTHosp1 As Double, MTEmer1 As Double
      Dim MTEx2 As Double, MTCE2 As Double, MTHosp2 As Double, MTEmer2 As Double
      Dim lbEsOpenOffice As Boolean
      
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
        lcArchivoExcel = App.Path + "\Plantillas\LaboratorioProductividadConsolidado.ods"
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
        Set oExcel = GalenhosExcelApplication()
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LaboratorioProductividadConsolidado.xls")
        oWorkBookPlantilla.Worksheets("ProductividadConsolidado").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
      'Inicio de Impresion
      Dim J As Integer, Anio
      Anio = cmbYear.Text
      If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(1, 1).setFormula("AÑO " & Anio)
      Else
        oWorkSheet.Cells(2, 2).Value = "AÑO " & Anio
      End If
      iFila = 6
      iCol = 2
      II = 0
      TEx2 = 0: MTEx2 = 0: TCE2 = 0: MTCE2 = 0: THosp2 = 0: MTHosp2 = 0: TEmer2 = 0: MTEmer2 = 0
      
      For J = 1 To 12
        TEx1 = 0: MTEx1 = 0: TCE1 = 0: MTCE1 = 0: THosp1 = 0: MTHosp1 = 0: TEmer1 = 0: MTEmer1 = 0
        rsReporte.MoveFirst
        If J = 1 Then
          mda_FechaInicio = "01/01/" & Anio & " 00:00:00"
          mda_FechaFin = "31/01/" & Anio & " 23:59:59"
        ElseIf J = 2 Then
          Dim TT As String
          TT = "28"
          If sighentidades.EsBisiesto(CInt(Anio)) = True Then TT = "29"
          mda_FechaInicio = "01/02/" & Anio & " 00:00:00"
          mda_FechaFin = TT & "/02/" & Anio & " 23:59:59"
        ElseIf J = 3 Then
          mda_FechaInicio = "01/03/" & Anio & " 00:00:00"
          mda_FechaFin = "31/03/" & Anio & " 23:59:59"
        ElseIf J = 4 Then
          mda_FechaInicio = "01/04/" & Anio & " 00:00:00"
          mda_FechaFin = "30/04/" & Anio & " 23:59:59"
        ElseIf J = 5 Then
          mda_FechaInicio = "01/05/" & Anio & " 00:00:00"
          mda_FechaFin = "31/05/" & Anio & " 23:59:59"
        ElseIf J = 6 Then
          mda_FechaInicio = "01/06/" & Anio & " 00:00:00"
          mda_FechaFin = "30/06/" & Anio & " 23:59:59"
        ElseIf J = 7 Then
          mda_FechaInicio = "01/07/" & Anio & " 00:00:00"
          mda_FechaFin = "31/07/" & Anio & " 23:59:59"
        ElseIf J = 8 Then
          mda_FechaInicio = "01/08/" & Anio & " 00:00:00"
          mda_FechaFin = "31/08/" & Anio & " 23:59:59"
        ElseIf J = 9 Then
          mda_FechaInicio = "01/09/" & Anio & " 00:00:00"
          mda_FechaFin = "30/09/" & Anio & " 23:59:59"
        ElseIf J = 10 Then
          mda_FechaInicio = "01/10/" & Anio & " 00:00:00"
          mda_FechaFin = "31/10/" & Anio & " 23:59:59"
        ElseIf J = 11 Then
          mda_FechaInicio = "01/11/" & Anio & " 00:00:00"
          mda_FechaFin = "30/11/" & Anio & " 23:59:59"
        ElseIf J = 12 Then
          mda_FechaInicio = "01/12/" & Anio & " 00:00:00"
          mda_FechaFin = "31/12/" & Anio & " 23:59:59"
        End If
        Do While Not rsReporte.EOF
          II = II + 1
          TEx = 0: MTEx = 0
          Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeExternos(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
          If rsReporte1.State = adStateOpen Then
            If rsReporte1.RecordCount > 0 Then
              rsReporte1.MoveFirst
              Do While Not rsReporte1.EOF
                TEx = TEx + rsReporte1!Cantidad
                MTEx = MTEx + rsReporte1!Cantidad * rsReporte1!precio
                rsReporte1.MoveNext
              Loop
            End If
          End If
          Set rsReporte1 = Nothing
          TEx1 = TEx1 + TEx
          MTEx1 = MTEx1 + MTEx
        
          TCE = 0: MTCE = 0
          Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeConsultoriosExternos(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
          If rsReporte1.State = adStateOpen Then
            If rsReporte1.RecordCount > 0 Then
              rsReporte1.MoveFirst
              Do While Not rsReporte1.EOF
                TCE = TCE + rsReporte1!Cantidad
                MTCE = MTCE + rsReporte1!Cantidad * rsReporte1!precio
                rsReporte1.MoveNext
              Loop
            End If
          End If
          Set rsReporte1 = Nothing
          TCE1 = TCE1 + TCE
          MTCE1 = MTCE1 + MTCE
        
          TEmer = 0: MTEmer = 0
          Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeEmergencia(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
          If rsReporte1.State = adStateOpen Then
            If rsReporte1.RecordCount > 0 Then
              rsReporte1.MoveFirst
              Do While Not rsReporte1.EOF
                TEmer = TEmer + rsReporte1!Cantidad
                MTEmer = MTEmer + rsReporte1!Cantidad * rsReporte1!precio
                rsReporte1.MoveNext
              Loop
            End If
          End If
          Set rsReporte1 = Nothing
          TEmer1 = TEmer1 + TEmer
          MTEmer1 = MTEmer1 + MTEmer
        
          THosp = 0: MTHosp = 0
          Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeHospitalizacion(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
          If rsReporte1.State = adStateOpen Then
            If rsReporte1.RecordCount > 0 Then
              rsReporte1.MoveFirst
              Do While Not rsReporte1.EOF
                THosp = THosp + rsReporte1!Cantidad
                MTHosp = MTHosp + rsReporte1!Cantidad * rsReporte1!precio
                rsReporte1.MoveNext
              Loop
            End If
          End If
          Set rsReporte1 = Nothing
          THosp1 = THosp1 + THosp
          MTHosp1 = MTHosp1 + MTHosp
          
          rsReporte.MoveNext
        Loop
      
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(TEx1)
            Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(MTEx1)
            Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCE1)
            Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(MTCE1)
            Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TEmer1)
            Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(MTEmer1)
            Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(THosp1)
            Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(MTHosp1)
            Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(TEx1 + TCE1 + THosp1 + TEmer1)
            Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(MTEx1 + MTCE1 + MTHosp1 + MTEmer1)
        Else
            oWorkSheet.Cells(iFila, iCol + 1).Value = TEx1
            oWorkSheet.Cells(iFila, iCol + 2).Value = MTEx1
            oWorkSheet.Cells(iFila, iCol + 3).Value = TCE1
            oWorkSheet.Cells(iFila, iCol + 4).Value = MTCE1
            oWorkSheet.Cells(iFila, iCol + 5).Value = TEmer1
            oWorkSheet.Cells(iFila, iCol + 6).Value = MTEmer1
            oWorkSheet.Cells(iFila, iCol + 7).Value = THosp1
            oWorkSheet.Cells(iFila, iCol + 8).Value = MTHosp1
            oWorkSheet.Cells(iFila, iCol + 9).Value = TEx1 + TCE1 + THosp1 + TEmer1
            oWorkSheet.Cells(iFila, iCol + 10).Value = MTEx1 + MTCE1 + MTHosp1 + MTEmer1
            'oWorkSheet.Cells(iFila, 5).Value = TCant ' "=suma(E6:E" & iFila - 2 & ")"
        End If
        If lbEsOpenOffice = True Then
        Else
            oWorkSheet.Range(oWorkSheet.Cells(iFila, 2), oWorkSheet.Cells(iFila, 12)).Borders.LineStyle = 1
        End If
        TEx2 = TEx2 + TEx1
        MTEx2 = MTEx2 + MTEx1
        TCE2 = TCE2 + TCE1
        MTCE2 = MTCE2 + MTCE1
        TEmer2 = TEmer2 + TEmer1
        MTEmer2 = MTEmer2 + MTEmer1
        THosp2 = THosp2 + THosp1
        MTHosp2 = MTHosp2 + MTHosp1
      Next J
      iFila = iFila + 1
      iFila = iFila + 1
      If lbEsOpenOffice = True Then
        Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula("TOTAL")
        Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(TEx2)
        Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula(MTEx2)
        Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TCE2)
        Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(MTCE2)
        Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TEmer2)
        Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(MTEmer2)
        Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(THosp2)
        Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(MTHosp2)
        Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(TEx2 + TCE2 + THosp2 + TEmer2)
        Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(MTEx2 + MTCE2 + MTHosp2 + MTEmer2)
      Else
        oWorkSheet.Cells(iFila, iCol).Value = "TOTAL"
        oWorkSheet.Cells(iFila, iCol + 1).Value = TEx2
        oWorkSheet.Cells(iFila, iCol + 2).Value = MTEx2
        oWorkSheet.Cells(iFila, iCol + 3).Value = TCE2
        oWorkSheet.Cells(iFila, iCol + 4).Value = MTCE2
        oWorkSheet.Cells(iFila, iCol + 5).Value = TEmer2
        oWorkSheet.Cells(iFila, iCol + 6).Value = MTEmer2
        oWorkSheet.Cells(iFila, iCol + 7).Value = THosp2
        oWorkSheet.Cells(iFila, iCol + 8).Value = MTHosp2
        oWorkSheet.Cells(iFila, iCol + 9).Value = TEx2 + TCE2 + THosp2 + TEmer2
        oWorkSheet.Cells(iFila, iCol + 10).Value = MTEx2 + MTCE2 + MTHosp2 + MTEmer2
      End If
      If lbEsOpenOffice = True Then
        If lbEsOpenOffice = True Then
            Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":L" & CStr(iFila))
            mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
        Else
            oWorkSheet.Range(oWorkSheet.Cells(iFila, iCol), oWorkSheet.Cells(iFila, 12)).Borders.LineStyle = 1
        End If
      End If
      MousePointer = 0
      If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 12
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
      Else
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
              oWorkSheet.PageSetup.PrintArea = "$A$1:$L$" & (iFila + 2)
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
    Else
      MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
    End If
    txtFdesde.Enabled = True
    txtFhasta.Enabled = True
    txtHrInicio.Enabled = True
    txtHrFin.Enabled = True
    cmbYear.Enabled = True
    Me.MousePointer = 1
    cmbYear.SetFocus
  End If
  
  
  
  Exit Sub
  If ValidaDatosObligatorios Then
    Dim lnIdProducto As Long, lcCodigo As String, lcNombre As String
    Dim lnSalidas As Long, lnPrecio As Long
    Dim lnSalidasImg As Long, lnSaldoInicial As Long, lnSaldofinal As Long
    If rsReporte.RecordCount > 0 Then
      If rsTmp.State <> adStateClosed Then Set rsTmp = Nothing
      With rsTmp
        .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
        .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
        .Fields.Append "Buenos", adInteger, 4, adFldIsNullable
        .Fields.Append "Fallados", adInteger, 4, adFldIsNullable
        .Fields.Append "Repetidos", adInteger, 4, adFldIsNullable
        .Fields.Append "Total", adInteger, 4, adFldIsNullable
        .Fields.Append "Importe", adDouble
        .LockType = adLockOptimistic
        .Open
      End With
      rsReporte.MoveFirst
      Do While Not rsReporte.EOF
        lnIdProducto = rsReporte.Fields!idProductoCPT
        lcCodigo = rsReporte.Fields!Codigo
        lcNombre = rsReporte.Fields!Nombre
        '*******Saldo Inicial********
        lnSalidas = 0: lnPrecio = 0
        lnSalidasImg = 0: lnSaldoInicial = 0: lnSaldofinal = 0
        Do While Not rsReporte.EOF And lnIdProducto = rsReporte.Fields!idProductoCPT
          lnPrecio = lnPrecio + rsReporte.Fields!Total
          lnSalidas = lnSalidas + rsReporte.Fields!Cantidad
          'externos
          If IsNull(rsReporte.Fields!idTipoServicio) Then lnSalidasImg = lnSalidasImg + rsReporte.Fields!Cantidad
          'CE
          If rsReporte.Fields!idTipoServicio = 1 Then lnSaldoInicial = lnSaldoInicial + rsReporte.Fields!Cantidad
          'hosp/emerg
          If rsReporte.Fields!idTipoServicio > 1 Then lnSaldofinal = lnSaldofinal + rsReporte.Fields!Cantidad
          rsReporte.MoveNext
          If rsReporte.EOF Then Exit Do
        Loop
        rsTmp.AddNew
        rsTmp.Fields!Codigo = lcCodigo
        rsTmp.Fields!Nombre = lcNombre
        rsTmp.Fields!buenos = lnSalidasImg      'Externos
        rsTmp.Fields!fallados = lnSaldoInicial  'CE
        rsTmp.Fields!repetidos = lnSaldofinal   'Hosp/emerg
        rsTmp.Fields!Total = lnSalidas
        rsTmp.Fields!Importe = lnPrecio
        rsTmp.Update
      Loop
    End If
  End If
  If rsTmp.State = adStateClosed Then
    MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
    Exit Sub
  End If
  If rsTmp.EOF = True And rsTmp.BOF = True Then
    MsgBox "No hay datos para mostrar", vbInformation, "SIGH "
  Else
    Me.MousePointer = 11
    Dim oRptClaseCry As New frmCrystalR
    oRptClaseCry.Excel = IIf(chkExcel.Value = 1, True, False)
    oRptClaseCry.Archivo = "LabProductividad"
    oRptClaseCry.Tabla = rsTmp
    oRptClaseCry.Show vbModal
    Set oRptClaseCry = Nothing
    Set rsTmp = Nothing
    Me.MousePointer = 1
  End If

End Sub

Function ValidaDatosObligatorios() As Boolean
  If cmbYear.Text = "" Then ' txtFdesde.Text = "" Or txtFhasta.Text = "" Then
    MsgBox "Escoja un año", vbInformation, "SIGH " '"Ingrese Fechas de Inicio y Fecha de Fin", vbInformation, "SIGH "
    cmbYear.SetFocus
    ValidaDatosObligatorios = False
  Else
    ValidaDatosObligatorios = True
  End If
End Function

Private Sub btnCancelar_Click()
  Me.Visible = False
  LimpiarVariablesDeMemoria
End Sub

Private Sub Form_Initialize()
  'Set mo_cmbIdPuntoCarga.MiComboBox = cmbIdPuntoDeCarga
End Sub

Private Sub Form_Load()
  Dim I
  txtFdesde.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
  txtFhasta.Text = Date
  txtHrInicio.Text = "00:00:00"
  txtHrFin.Text = "23:59:59"
  For I = 2009 To Year(Now)
    cmbYear.AddItem I
  Next I
  '
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_GotFocus()
  SeleccionaMask txtFdesde
End Sub

Private Sub txtFdesde_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFdesde_LostFocus()
  If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
      MsgBox "La Fecha Inicial ingresada no es válida", vbInformation, "SIGH "
      txtFdesde = sighentidades.PrimerFechaDDMMYYDelMesActual 'Format(Now, "dd/mm/yyyy")
      txtFdesde.SetFocus
    End If
  End If
End Sub

Private Sub txtFhasta_GotFocus()
  SeleccionaMask txtFhasta
End Sub

Private Sub txtFhasta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFhasta_LostFocus()
  If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
      MsgBox "La Fecha Final ingresada no es válida", vbInformation, "SIGH "
      txtFhasta.Text = Format(Now, "dd/mm/yyyy")
      txtFhasta.SetFocus
    End If
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
  On Error Resume Next
  Set mo_ReglasFarmacia = Nothing
  Set mo_Teclado = Nothing
  Set mo_ReglasFacturacion = Nothing
  Set mo_reglasComunes = Nothing
  Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_GotFocus()
  SeleccionaMask txtHrFin
End Sub

Private Sub txtHrFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrFin_LostFocus()
  If txtHrFin.Text <> "__:__:__" Then
    If Not IsDate(txtHrFin.Text) Then
      MsgBox "La Hora Final ingresada no es válida.", vbInformation, "SIGH "
      txtHrFin.Text = "23:59:59"
      txtHrFin.SetFocus
    End If
  End If
End Sub

Private Sub txtHrInicio_GotFocus()
  SeleccionaMask txtHrInicio
End Sub

Private Sub txtHrInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHrInicio_LostFocus()
  If txtHrInicio.Text <> "__:__:__" Then
    If Not IsDate(txtHrInicio.Text) Then
      MsgBox "La Hora Inicial ingresada no es válida.", vbInformation, "SIGH "
      txtHrInicio.Text = "00:00:00"
      txtHrInicio.SetFocus
    End If
  End If
End Sub
