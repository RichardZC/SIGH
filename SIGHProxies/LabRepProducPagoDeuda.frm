VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form LabRepProducPagoDeuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio: Producción, Pagos y Deudas  por Fechas"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   Icon            =   "LabRepProducPagoDeuda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
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
      Height          =   1170
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   7755
      Begin VB.CheckBox chkSoloGestantes 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo GESTANTES"
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
         Left            =   5490
         Picture         =   "LabRepProducPagoDeuda.frx":0CCA
         TabIndex        =   12
         Top             =   645
         Width           =   2085
      End
      Begin VB.CheckBox chkSoloMovimiento 
         Caption         =   "Solo los que tienen MOVIMIENTO"
         Height          =   255
         Left            =   105
         TabIndex        =   11
         Top             =   645
         Value           =   1  'Checked
         Width           =   3315
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
         Left            =   3825
         Picture         =   "LabRepProducPagoDeuda.frx":0FDC
         TabIndex        =   10
         Top             =   675
         Visible         =   0   'False
         Width           =   765
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   240
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
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   15
      TabIndex        =   5
      Top             =   1215
      Width           =   7740
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LabRepProducPagoDeuda.frx":12EE
         DownPicture     =   "LabRepProducPagoDeuda.frx":174E
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
         Picture         =   "LabRepProducPagoDeuda.frx":1BC3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LabRepProducPagoDeuda.frx":2038
         DownPicture     =   "LabRepProducPagoDeuda.frx":24FC
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
         Picture         =   "LabRepProducPagoDeuda.frx":29E8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "LabRepProducPagoDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de producción Pago y Deuda
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

If wxFranklin = "*" Then Exit Sub

  If ValidaDatosObligatorios = True Then
    Dim rsReporte As New ADODB.Recordset
    Dim rsReporte1 As New ADODB.Recordset
    Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mda_FechaInicio As String ' Date
    Dim mda_FechaFin As String 'Date
    Dim mo_ReporteUtil As New ReporteUtil
    Dim lnHwnd As Long
    lnHwnd = Me.hwnd
    Dim lcSql As String
    
    mda_FechaInicio = txtFdesde.Text & " " & txtHrInicio.Text
    mda_FechaFin = txtFhasta.Text & " " & txtHrFin.Text
    
    Set rsReporte = mo_ReglasLaboratorio.SacarPruebasTodas()
    If rsReporte.RecordCount > 0 Then
      Dim iFila As Long, iCol As Integer, II As Integer, lcCodigoCPT As String
      Dim TEx As Double, TCE As Double, THosp As Double, TEmer As Double
      Dim TEx1 As Double, TCE1 As Double, THosp1 As Double, TEmer1 As Double
      Dim MTEx As Double, MTCE As Double, MTHosp As Double, MTEmer As Double
      Dim MTEx1 As Double, MTCE1 As Double, MTHosp1 As Double, MTEmer1 As Double
      Dim lbEsOpenOffice As Boolean, lbContinuar As Boolean
      
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
      rsReporte.MoveFirst
      
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\LaboratorioProductividad.ods"
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
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\LaboratorioProductividad.xls")
        oWorkBookPlantilla.Worksheets("Productividad").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
     End If
        If lbEsOpenOffice = True Then
             Call Feuille.getcellbyposition(1, 1).setFormula("Fecha Inicio " & mda_FechaInicio & "  -  Fecha Fin: " & mda_FechaFin & _
                                            IIf(chkSoloGestantes.Value = 1, "   (" & chkSoloGestantes.Caption & ")", ""))
        Else
        'Inicio de Impresion
             oWorkSheet.Cells(2, 2).Value = "Fecha Inicio " & mda_FechaInicio & "  -  Fecha Fin: " & mda_FechaFin & _
                                            IIf(chkSoloGestantes.Value = 1, "   (" & chkSoloGestantes.Caption & ")", "")
        End If
      iFila = 6
      iCol = 2
      II = 0 ': TCant = 0
      TEx1 = 0: MTEx1 = 0: TCE1 = 0: MTCE1 = 0: THosp1 = 0: MTHosp1 = 0: TEmer1 = 0: MTEmer1 = 0
      Do While Not rsReporte.EOF
        


        TEx = 0: MTEx = 0
        Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeExternos(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
        If rsReporte1.State = adStateOpen Then
            If chkSoloGestantes.Value = 1 Then
               rsReporte1.Filter = "Eo_EG>0"
            End If
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
        
        TCE = 0: MTCE = 0
        Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeConsultoriosExternos(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
        If rsReporte1.State = adStateOpen Then
        If chkSoloGestantes.Value = 1 Then
           rsReporte1.Filter = "Eo_EG>0"
        End If
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
        
        TEmer = 0: MTEmer = 0
        Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeEmergencia(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
        If rsReporte1.State = adStateOpen Then
        If chkSoloGestantes.Value = 1 Then
           rsReporte1.Filter = "Eo_EG>0"
        End If
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

        THosp = 0: MTHosp = 0
        Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeHospitalizacion(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
        If rsReporte1.State = adStateOpen Then
        If chkSoloGestantes.Value = 1 Then
           rsReporte1.Filter = "Eo_EG>0"
        End If
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
        
        lbContinuar = True
        If chkSoloMovimiento.Value = 1 Then
           If (TEx + TCE + THosp + TEmer) = 0 Then
               lbContinuar = False
           End If
        End If
        If lbContinuar = True Then
                II = II + 1
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol - 1, iFila - 1).setFormula(II)
                    Call Feuille.getcellbyposition(iCol + 0, iFila - 1).setFormula(Trim(rsReporte!codigoCPT))
                    Call Feuille.getcellbyposition(iCol + 1, iFila - 1).setFormula((IIf(IsNull(rsReporte!Nombre), "", rsReporte!Nombre)))
                Else
                    oWorkSheet.Cells(iFila, iCol).Value = II
                    oWorkSheet.Cells(iFila, iCol + 1).Value = Trim(rsReporte!codigoCPT)
                    oWorkSheet.Cells(iFila, iCol + 2).Value = rsReporte!Nombre
                    'oWorkSheet.Cells(iFila, iCol + 3).Value = rsReporte!idProducto
                End If
        
        
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TEx)
                    Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(MTEx)
                Else
                    oWorkSheet.Cells(iFila, iCol + 3).Value = TEx
                    oWorkSheet.Cells(iFila, iCol + 4).Value = MTEx
                End If
                TEx1 = TEx1 + TEx
                MTEx1 = MTEx1 + MTEx
                
'                TCE = 0: MTCE = 0
'                Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeConsultoriosExternos(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
'                If rsReporte1.State = adStateOpen Then
'                If rsReporte1.RecordCount > 0 Then
'                  rsReporte1.MoveFirst
'                  Do While Not rsReporte1.EOF
'                    TCE = TCE + rsReporte1!Cantidad
'                    MTCE = MTCE + rsReporte1!Cantidad * rsReporte1!precio
'                    rsReporte1.MoveNext
'                  Loop
'                End If
'                End If
'                Set rsReporte1 = Nothing
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TCE)
                    Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(MTCE)
                Else
                    oWorkSheet.Cells(iFila, iCol + 5).Value = TCE
                    oWorkSheet.Cells(iFila, iCol + 6).Value = MTCE
                End If
                TCE1 = TCE1 + TCE
                MTCE1 = MTCE1 + MTCE
                
'                TEmer = 0: MTEmer = 0
'                Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeEmergencia(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
'                If rsReporte1.State = adStateOpen Then
'                If rsReporte1.RecordCount > 0 Then
'                  rsReporte1.MoveFirst
'                  Do While Not rsReporte1.EOF
'                    TEmer = TEmer + rsReporte1!Cantidad
'                    MTEmer = MTEmer + rsReporte1!Cantidad * rsReporte1!precio
'                    rsReporte1.MoveNext
'                  Loop
'                End If
'                End If
'                Set rsReporte1 = Nothing
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(TEmer)
                    Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(MTEmer)
                Else
                    oWorkSheet.Cells(iFila, iCol + 7).Value = TEmer
                    oWorkSheet.Cells(iFila, iCol + 8).Value = MTEmer
                End If
                TEmer1 = TEmer1 + TEmer
                MTEmer1 = MTEmer1 + MTEmer
                
'                THosp = 0: MTHosp = 0
'                Set rsReporte1 = mo_ReglasLaboratorio.AveriguaConsumosDeHospitalizacion(mda_FechaInicio, mda_FechaFin, rsReporte!idProducto)
'                If rsReporte1.State = adStateOpen Then
'                If rsReporte1.RecordCount > 0 Then
'                  rsReporte1.MoveFirst
'                  Do While Not rsReporte1.EOF
'                    THosp = THosp + rsReporte1!Cantidad
'                    MTHosp = MTHosp + rsReporte1!Cantidad * rsReporte1!precio
'                    rsReporte1.MoveNext
'                  Loop
'                End If
'                End If
'                Set rsReporte1 = Nothing
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(THosp)
                    Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(MTHosp)
                Else
                    oWorkSheet.Cells(iFila, iCol + 9).Value = THosp
                    oWorkSheet.Cells(iFila, iCol + 10).Value = MTHosp
                End If
                THosp1 = THosp1 + THosp
                MTHosp1 = MTHosp1 + MTHosp
                
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(TEx + TCE + THosp + TEmer)
                    Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(MTEx + MTCE + MTHosp + MTEmer)
                Else
                    oWorkSheet.Cells(iFila, iCol + 11).Value = TEx + TCE + THosp + TEmer
                    oWorkSheet.Cells(iFila, iCol + 12).Value = MTEx + MTCE + MTHosp + MTEmer
                End If
                
                iFila = iFila + 1
        End If
        lcCodigoCPT = rsReporte!codigoCPT
        Do While Not rsReporte.EOF And lcCodigoCPT = rsReporte!codigoCPT
           rsReporte.MoveNext
           If rsReporte.EOF Then
              Exit Do
           End If
        Loop
      Loop
        If lbEsOpenOffice = True Then
        Else
'            oWorkSheet.Range(oWorkSheet.Cells(iFila - II, 2), oWorkSheet.Cells(iFila - 1, 14)).Borders.LineStyle = 1
        End If
      

            iFila = iFila + 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula("TOTAL")
                Call Feuille.getcellbyposition(iCol + 2, iFila - 1).setFormula(TEx1)
                Call Feuille.getcellbyposition(iCol + 3, iFila - 1).setFormula(MTEx1)
                Call Feuille.getcellbyposition(iCol + 4, iFila - 1).setFormula(TCE1)
                Call Feuille.getcellbyposition(iCol + 5, iFila - 1).setFormula(MTCE1)
                Call Feuille.getcellbyposition(iCol + 6, iFila - 1).setFormula(TEmer1)
                Call Feuille.getcellbyposition(iCol + 7, iFila - 1).setFormula(MTEmer1)
                Call Feuille.getcellbyposition(iCol + 8, iFila - 1).setFormula(THosp1)
                Call Feuille.getcellbyposition(iCol + 9, iFila - 1).setFormula(MTHosp1)
                Call Feuille.getcellbyposition(iCol + 10, iFila - 1).setFormula(TEx1 + TCE1 + THosp1 + TEmer1)
                Call Feuille.getcellbyposition(iCol + 11, iFila - 1).setFormula(MTEx1 + MTCE1 + MTHosp1 + MTEmer1)
            Else
                oWorkSheet.Cells(iFila, 4).Value = "TOTAL"
                oWorkSheet.Cells(iFila, iCol + 3).Value = TEx1
                oWorkSheet.Cells(iFila, iCol + 4).Value = MTEx1
                oWorkSheet.Cells(iFila, iCol + 5).Value = TCE1
                oWorkSheet.Cells(iFila, iCol + 6).Value = MTCE1
                oWorkSheet.Cells(iFila, iCol + 7).Value = TEmer1
                oWorkSheet.Cells(iFila, iCol + 8).Value = MTEmer1
                oWorkSheet.Cells(iFila, iCol + 9).Value = THosp1
                oWorkSheet.Cells(iFila, iCol + 10).Value = MTHosp1
                oWorkSheet.Cells(iFila, iCol + 11).Value = TEx1 + TCE1 + THosp1 + TEmer1
                oWorkSheet.Cells(iFila, iCol + 12).Value = MTEx1 + MTCE1 + MTHosp1 + MTEmer1
                  
                'oWorkSheet.Cells(iFila, 5).Value = TCant ' "=suma(E6:E" & iFila - 2 & ")"
            End If

        If lbEsOpenOffice = True Then
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("D" & CStr(iFila) & ":N" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                oWorkSheet.Range(oWorkSheet.Cells(iFila, 4), oWorkSheet.Cells(iFila, 14)).Borders.LineStyle = 1
            End If
        End If
      MousePointer = 0
      If lbEsOpenOffice = True Then
        Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
        PrintArea(0).Sheet = 0
        PrintArea(0).startcolumn = 1
        PrintArea(0).StartRow = 0
        PrintArea(0).EndColumn = 14
        PrintArea(0).EndRow = iFila
        Call Feuille.SetPrintAreas(PrintArea())
        Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
        MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
      Else
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
               oWorkSheet.PageSetup.PrintArea = "$A$1:$O$" & (iFila + 2)
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
  If txtFdesde.Text = "" Or txtFhasta.Text = "" Then
    MsgBox "Ingrese Fechas de Inicio y Fecha de Fin", vbInformation, "SIGH "
    txtFdesde.SetFocus
    ValidaDatosObligatorios = False
  Else
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
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

Sub InicializaFechaHora()
  txtFdesde.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
  txtFhasta.Text = Date
  txtHrInicio.Text = "00:00:00"
  txtHrFin.Text = "23:59:59"

End Sub
Private Sub Form_Load()
  InicializaFechaHora
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



Private Sub txtFdesde_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFdesde_LostFocus()
  If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
    If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
      MsgBox "La Fecha Inicial ingresada no es válida", vbInformation, "SIGH "
      txtFdesde = sighentidades.PrimerFechaDDMMYYDelMesActual 'Format(Now, sighEntidades.DevuelveFechaSoloFormato_DMY)
      txtFdesde.SetFocus
    End If
  End If
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
