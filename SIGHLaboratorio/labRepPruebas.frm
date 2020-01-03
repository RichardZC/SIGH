VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form labRepPruebas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laboratorio: Reporte de Pruebas Realizadas"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "labRepPruebas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7710
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
      Height          =   2385
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   7650
      Begin VB.Frame Frame1 
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   120
         TabIndex        =   12
         Top             =   630
         Width           =   7305
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   15
            Top             =   1110
            Value           =   -1  'True
            Width           =   855
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
            Left            =   420
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   600
            Visible         =   0   'False
            Width           =   6720
         End
         Begin VB.OptionButton optIndividual 
            Caption         =   "Individual"
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
            Left            =   150
            TabIndex        =   13
            Top             =   300
            Width           =   1215
         End
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1470
         TabIndex        =   0
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   4740
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
         Left            =   2850
         TabIndex        =   1
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
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   6120
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
         Left            =   4170
         TabIndex        =   10
         Top             =   270
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
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   60
      TabIndex        =   7
      Top             =   2400
      Width           =   7650
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "Exportar En Excel"
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
         Left            =   240
         Picture         =   "labRepPruebas.frx":0CCA
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Visualizar(F2)"
         DisabledPicture =   "labRepPruebas.frx":0FDC
         DownPicture     =   "labRepPruebas.frx":143C
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
         Left            =   2408
         Picture         =   "labRepPruebas.frx":18B1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "labRepPruebas.frx":1D26
         DownPicture     =   "labRepPruebas.frx":21EA
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
         Left            =   3938
         Picture         =   "labRepPruebas.frx":26D6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdAuditoria 
      Height          =   285
      Left            =   -4170
      TabIndex        =   6
      Top             =   3030
      Visible         =   0   'False
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   503
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Reporte de Pruebas realizadas por el personal, en un rango de fechas"
   End
End
Attribute VB_Name = "labRepPruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte de Pruebas con Resultados
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Datos As New SIGHDatos.CatalogoServicios
Dim mo_cmbIdPuntoCarga As New sighentidades.ListaDespleglable
Dim mo_cmbUsuario As New sighentidades.ListaDespleglable
Dim mo_cmbResponsable As New sighentidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim rsReporte As New ADODB.Recordset
Dim rsReporte1 As New ADODB.Recordset
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long
Dim rsTmp As New ADODB.Recordset
Dim rsTmp1 As New ADODB.Recordset
Dim FI As Date, FF As Date, HI As String, HF As String

Dim lcNombreTablaCab As String, lcNombreTablaDet As String
Dim lcTexto1 As String, lnTotal As Double
Dim mrs_Cab As New Recordset, mrs_Det As New Recordset, mrs_Shape As New Recordset
    
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Private Function Verifica() As Boolean
  Verifica = False
  If optTodos.Value = True Then
    If Not (IsDate(FI)) Or Not (IsDate(FF)) Or Len(Trim(txtFdesde.Text)) < 10 Or Len(Trim(txtFhasta.Text)) < 10 Then
      Verifica = False
    Else
      Verifica = True
    End If
  ElseIf optIndividual.Value = True Then
    If cmbResponsable.Text = "" Or Not (IsDate(txtFdesde.Text)) Or Not (IsDate(txtFhasta.Text)) Then
      Verifica = False
    Else
      Verifica = True
    End If
  Else
    MsgBox "Debe escoger el Tipo de Personal: Todos/Individual", vbInformation, "SIGH "
  End If
  If IsDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) And IsDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
  End If
End Function

Private Sub btnAceptar_Click()
  If lcNombreTablaCab = "" Or lcNombreTablaDet = "" Then Exit Sub
  If Verifica = False Then Exit Sub
  If optTodos.Value = True Then
    Set mrs_Cab = mo_ReglasLaboratorio.ReporteSeleccionaTodoContenido(lcNombreTablaCab)
    Set mrs_Det = mo_ReglasLaboratorio.ReporteSeleccionaTodoContenido(lcNombreTablaDet)
    If mrs_Cab.EOF = True And mrs_Cab.BOF = True Then
         MsgBox "No existen datos para mostrar", vbInformation, "SIGH "
         Exit Sub
    End If
    mrs_Cab.Close
    mrs_Det.Close
    Set mrs_Shape = mo_ReglasLaboratorio.ReporteMuestraResultado
    With RepLabPruebaTodos
      .Orientation = rptOrientPortrait
      .Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
      .Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
      .Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
      .Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
      .Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
      
      .Sections("Cabecera").Controls("FIni").Caption = FI
      .Sections("Cabecera").Controls("FFin").Caption = FF
      Set .Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
      Set .DataSource = mrs_Shape
      .DataMember = ""
      With .Sections("CabAgrupa")
        .Controls("txtCodigoP").DataMember = ""
        .Controls("txtCodigoP").DataField = "NroCuenta"
        .Controls("txtNombreP").DataMember = ""
        .Controls("txtNombreP").DataField = "Paciente"
      End With
      With .Sections("DetGrupo")
        .Controls("txtCPrueba").DataMember = "Hijo"
        .Controls("txtCPrueba").DataField = "idUsuario"
        .Controls("txtNPrueba").DataMember = "Hijo"
        .Controls("txtNPrueba").DataField = "ConsumoDescripcion"
        .Controls("txtCantidadPr").DataMember = "Hijo"
        .Controls("txtCantidadPr").DataField = "ConsumoImporte"
      End With
      With .Sections("PieAgrupa")
        .Controls("fncSuma").DataMember = "Hijo"
        .Controls("fncSuma").DataField = "ConsumoImporte"
      End With
      .Sections("PiePagina").Controls("lblTotal").Caption = lnTotal
      .Show 1
    End With
    'debb-27/05/2015
    Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
    mo_reglasComunes.grabaTablaAuditoria ("RepLabPruebaTodos: " & _
                                   FI & " " & FF)
    Set mo_reglasComunes = Nothing
    '
    
    
    
    
    
    Exit Sub
    Set RepLabPruebaTodos.DataSource = rsTmp
    RepLabPruebaTodos.DataMember = ""
    RepLabPruebaTodos.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    RepLabPruebaTodos.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
    RepLabPruebaTodos.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    RepLabPruebaTodos.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
    RepLabPruebaTodos.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
    RepLabPruebaTodos.Sections("Cabecera").Controls("FIni").Caption = FI
    RepLabPruebaTodos.Sections("Cabecera").Controls("FFin").Caption = FF
    Set RepLabPruebaTodos.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
    'RepLabPruebaTodos.Sections("CabAgrupa").Controls("txtCodigoP").DataMember = "Hijo"
    RepLabPruebaTodos.Sections("CabAgrupa").Controls("txtCodigoP").DataField = "CodigoP"
    If chkExcel.Value = 1 Then mo_ReglasReportes.ExportarRecordSetAexcel rsTmp, "Pruebas todos los Empleados", "PRUEBAS REALIZADAS POR TODO EL PERSONAL DE LABORATORIO", "Nro Pruebas: " & Trim(Str(rsTmp.RecordCount)), Me.hwnd
    RepLabPruebaTodos.Show vbModal
    'debb-27/05/2015
    mo_reglasComunes.grabaTablaAuditoria ("RepLabPruebas: " & _
                                   FI & " " & FF)
  Else
    If rsTmp.State = adStateClosed Then
      MsgBox "No existen datos para mostrar", vbInformation, "SIGH "
      Exit Sub
    Else
      If (rsTmp.EOF = True And rsTmp.BOF = True) Then
        MsgBox "No existen datos para mostrar", vbInformation, "SIGH "
        Exit Sub
      End If
    End If
    Set RepLabPruebas.DataSource = rsTmp
    RepLabPruebas.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
    RepLabPruebas.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
    RepLabPruebas.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
    RepLabPruebas.Sections("cabecera").Controls("lblhora").Caption = lcBuscaParametro.RetornaHoraServidorSQL
    RepLabPruebas.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaServidorSQL
    RepLabPruebas.Sections("Cabecera").Controls("NombreP").Caption = cmbResponsable.Text
    RepLabPruebas.Sections("Cabecera").Controls("FIni").Caption = FI
    RepLabPruebas.Sections("Cabecera").Controls("FFin").Caption = FF
    Set RepLabPruebas.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
    If chkExcel.Value = 1 Then mo_ReglasReportes.ExportarRecordSetAexcel rsTmp, "Pruebas por Empleado", "PRUEBAS REALIZADAS POR EL PERSONAL DE LABORATORIO", "Nro Pruebas: " & Trim(Str(rsTmp.RecordCount)), Me.hwnd
    RepLabPruebas.Show vbModal
  End If
  
  'Dim oRptClaseCry As New frmCrystalR
  'oRptClaseCry.Excel = IIf(chkExcel.Value = 1, True, False)
  'If optTodos.Value = True Then
  '  oRptClaseCry.Archivo = "LabRepPruebasTodos"
  'Else
  '  oRptClaseCry.Archivo = "LabRepPruebas"
  'End If
  'oRptClaseCry.Tabla = rsTmp
  'oRptClaseCry.Show vbModal
  'Set oRptClaseCry = Nothing

  'Set rsTmp = Nothing
End Sub

Function ValidaDatosObligatorios() As Boolean
    If optIndividual.Value = False Then Exit Function
    sMensaje = ""
    If cmbResponsable.Text = "" Then
        sMensaje = sMensaje + "- Elija algun Empleado del Laboratorio" + Chr(13)
        cmbResponsable.SetFocus
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, "SIGH "
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function

Private Sub btnCancelar_Click()
  Me.Visible = False
  LimpiarVariablesDeMemoria
  Unload Me
End Sub

Private Sub ConfiguraFecha()
  If Len(Trim(txtFdesde.Text)) < 10 Or Not IsDate(txtHrInicio.Text) Or Len(Trim(txtFhasta.Text)) < 10 Or Not IsDate(txtHrFin.Text) Then Exit Sub
  If txtHrInicio.Text <> "" Then
    HI = " " & txtHrInicio.Text '& ":00"
  Else
    HI = " 00:00:00"
  End If
  If txtHrFin.Text <> "" Then
    HF = " " & txtHrFin.Text ' & ":59"
  Else
    HF = " 23:59:59"
  End If
  FI = CDate(txtFdesde.Text & HI)
  FF = CDate(txtFhasta.Text & HF)
End Sub

Private Sub cmbResponsable_Click()
  ConfiguraFecha
  Set grdAuditoria.DataSource = Nothing
  Set rsTmp = Nothing
  If optIndividual.Value = False Or optTodos.Value = True Then Exit Sub
  If Verifica Then 'btnAceptar_Click
    Dim ml_IdPuntoCarga As Long
    If ValidaDatosObligatorios Then
      Me.MousePointer = 11
      ConfiguraFecha
      ml_IdPuntoCarga = Val(mo_cmbResponsable.BoundText)
      Set rsReporte = mo_ReglasLaboratorio.LaboratorioPruebasPorEmpleado(FI, FF, ml_IdPuntoCarga)
      If rsReporte.EOF = True And rsReporte.BOF = True Then
        Me.MousePointer = 1
        Exit Sub
      End If
      rsReporte.MoveFirst
      If rsTmp.State = adStateOpen Then Set rsTmp = Nothing
      With rsTmp
        .Fields.Append "NombreP", adVarChar, 100, adFldIsNullable
        .Fields.Append "CodigoP", adVarChar, 20, adFldIsNullable
        .Fields.Append "FIni", adVarChar, 20, adFldIsNullable
        .Fields.Append "FFin", adVarChar, 20, adFldIsNullable
        .Fields.Append "CPrueba", adVarChar, 10, adFldIsNullable
        .Fields.Append "NPrueba", adVarChar, 100, adFldIsNullable
        .Fields.Append "CantidadPr", adInteger
        .LockType = adLockOptimistic
        .Open
      End With
      
      Do While Not rsReporte.EOF
        rsTmp.AddNew
        rsTmp!nombrep = Left(cmbResponsable.Text, 100)
        rsTmp!Codigop = mo_cmbResponsable.BoundText
        rsTmp!FIni = Format(FI, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) 'txtFdesde.Text & " " & txtHrInicio.Text
        rsTmp!FFin = Format(FF, sighentidades.DevuelveFechaSoloFormato_DMY_HMS) 'txtFhasta.Text & " " & txtHrFin.Text
        rsTmp!cprueba = rsReporte!Codigo
        rsTmp!nPrueba = rsReporte!Nombre
        rsTmp!cantidadpr = rsReporte!Cantidad
        rsTmp.Update
        rsReporte.MoveNext
      Loop
      
      Set grdAuditoria.DataSource = rsReporte
      mo_Apariencia.ConfigurarFilasBiColores grdAuditoria, sighentidades.GrillaConFilasBicolor
      Me.MousePointer = 1
    End If
  End If
End Sub

Private Sub Form_Initialize()
  Set mo_cmbResponsable.MiComboBox = cmbResponsable
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
  txtFdesde.Text = Date
  txtFhasta.Text = Date
  txtHrInicio.Text = "07:00:00"
  txtHrFin.Text = "18:59:59"
  mo_cmbResponsable.BoundColumn = "idEmpleado"
  mo_cmbResponsable.ListField = "ApNom"
  Set mo_cmbResponsable.RowSource = mo_ReglasLaboratorio.TodosEmpleadosDeLab()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
  Select Case KeyCode
    Case vbKeyEscape
      btnCancelar_Click
    Case vbKeyF2
      btnAceptar_Click
  End Select
End Sub

Private Sub optIndividual_Click()
  Set grdAuditoria.DataSource = Nothing
  Set rsTmp = Nothing
  ConfiguraFecha
  If optIndividual.Value = True Then
    cmbResponsable.Visible = True
  Else
    cmbResponsable.Visible = False
  End If
End Sub

Private Sub optTodos_Click()
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
  oConexion.CursorLocation = adUseClient
  oConexion.CommandTimeout = 300
  oConexion.Open sighentidades.CadenaConexionShape
  Set grdAuditoria.DataSource = Nothing
  Set rsTmp = Nothing
  lnTotal = 0
  ConfiguraFecha
  If optTodos.Value = False Or optIndividual.Value = True Then Exit Sub
  If optTodos.Value = True Then
    cmbResponsable.Visible = False
    If Verifica = False Then Exit Sub
    Set rsTmp1 = mo_ReglasLaboratorio.TodosEmpleadosDeLab
    If rsTmp1.EOF = True And rsTmp1.BOF = True Then Exit Sub
    lcNombreTablaCab = "reporte_cabecera"
    lcNombreTablaDet = "reporte_detalle"
    If mrs_Cab.State = adStateOpen Then Set mrs_Cab = Nothing
    If mrs_Det.State = adStateOpen Then Set mrs_Det = Nothing


'aqui me quede




    mo_ReglasLaboratorio.ReporteEliminaTodoContenido (lcNombreTablaCab)
    mo_ReglasLaboratorio.ReporteEliminaTodoContenido (lcNombreTablaDet)
    rsTmp1.MoveFirst
    Do While Not rsTmp1.EOF
      Set rsReporte1 = mo_ReglasLaboratorio.LaboratorioPruebasPorEmpleado(FI, FF, rsTmp1!idEmpleado)
      If rsReporte1.RecordCount > 0 Then
         'Cabecera
         With oCommand
              .CommandType = adCmdStoredProc
              Set .ActiveConnection = oConexion
              .CommandTimeout = 150
              .CommandText = "reporte_cabeceraInsertar"
              Set oParameter = .CreateParameter("@nroCuenta", adInteger, adParamInput, 0, rsTmp1!idEmpleado): .Parameters.Append oParameter
              Set oParameter = .CreateParameter("@Paciente", adVarChar, adParamInput, 100, Left(rsTmp1!apNom, 100)): .Parameters.Append oParameter
              .Execute
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
          
        rsReporte1.MoveFirst
        Do While Not rsReporte1.EOF
          lnTotal = lnTotal + rsReporte1!Cantidad
            'Detalle
             With oCommand
                  .CommandType = adCmdStoredProc
                  Set .ActiveConnection = oConexion
                  .CommandTimeout = 150
                  .CommandText = "reporte_detalleInsertar"
                  Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, rsReporte1!idProducto): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@nroCuenta", adInteger, adParamInput, 0, rsTmp1!idEmpleado): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@ConsumoDescripcion", adVarChar, adParamInput, 100, rsReporte1!Nombre): .Parameters.Append oParameter
                  Set oParameter = .CreateParameter("@ConsumoImporte", adCurrency, adParamInput, 0, rsReporte1!Cantidad): .Parameters.Append oParameter
                  .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
            rsReporte1.MoveNext
        Loop
      End If
      rsTmp1.MoveNext
    Loop
    mo_Apariencia.ConfigurarFilasBiColores grdAuditoria, sighentidades.GrillaConFilasBicolor
  Else
    cmbResponsable.Visible = True
  End If
   Set oConexion = Nothing
   Set oCommand = Nothing
End Sub

Private Sub txtFdesde_Change()
  If Verifica = False Then Exit Sub
  cmbResponsable_Click
  optTodos_Click
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
      txtFdesde.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
      txtFdesde.SetFocus
    End If
  End If
End Sub

Private Sub txtFhasta_Change()
  txtFdesde_Change
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
      txtFhasta.Text = Format(Now, sighentidades.DevuelveFechaSoloFormato_DMY)
      txtFhasta.SetFocus
    End If
  End If
End Sub

Sub LimpiarVariablesDeMemoria()
  On Error Resume Next
  Set mo_ReglasFarmacia = Nothing
  Set mo_Teclado = Nothing
  Set mo_cmbIdPuntoCarga = Nothing
  Set mo_cmbUsuario = Nothing
  Set mo_ReglasFacturacion = Nothing
  Set mo_reglasComunes = Nothing
  Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_Change()
  txtFdesde_Change
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
      txtHrFin.Text = "18:59:59"
      txtHrFin.SetFocus
    End If
  End If
End Sub

Private Sub txtHrInicio_Change()
  txtFdesde_Change
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
      txtHrInicio.Text = "07:00:00"
      txtHrInicio.SetFocus
    End If
  End If
End Sub
