VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucFarmVentasLista 
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12780
   LockControls    =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   12780
   Begin VB.Frame fraBusqueda 
      Caption         =   "Busqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   0
      TabIndex        =   7
      Top             =   570
      Width           =   12705
      Begin VB.CommandButton bntReporte 
         Height          =   315
         Left            =   11160
         Picture         =   "ucFarmVentasLista.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Imprimir LISTA"
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox chkSoloBoletas 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo DOCUMENTOS emitidos por PREVENTAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   18
         Top             =   885
         Width           =   4185
      End
      Begin VB.TextBox txtNCuenta 
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
         Left            =   4950
         MaxLength       =   9
         TabIndex        =   1
         Top             =   870
         Width           =   1545
      End
      Begin VB.TextBox txtNDocumento 
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
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   0
         Top             =   870
         Width           =   1665
      End
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   8460
         Picture         =   "ucFarmVentasLista.ctx":04D9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton btnLimpiar 
         Height          =   315
         Left            =   9810
         Picture         =   "ucFarmVentasLista.ctx":3122
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1275
      End
      Begin Threed.SSOption optPreventa 
         Height          =   255
         Left            =   6690
         TabIndex        =   2
         Top             =   450
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   445
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
         Caption         =   "PreVenta"
      End
      Begin Threed.SSOption optVenta 
         Height          =   285
         Left            =   6690
         TabIndex        =   3
         Top             =   720
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   508
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
         Caption         =   "Venta_Directa"
         Value           =   -1
      End
      Begin MSDataListLib.DataCombo cmbFuenteFinanciamiento 
         Height          =   315
         Left            =   9900
         TabIndex        =   14
         Top             =   585
         Width           =   2715
         _ExtentX        =   4784
         _ExtentY        =   529
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BoundColumn     =   ""
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbAlmacenV 
         Height          =   330
         Left            =   150
         TabIndex        =   15
         Top             =   480
         Width           =   2925
         _ExtentX        =   5144
         _ExtentY        =   550
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtFinicio 
         Height          =   315
         Left            =   3090
         TabIndex        =   16
         Top             =   480
         Width           =   1815
         _ExtentX        =   3196
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFfinal 
         Height          =   315
         Left            =   4935
         TabIndex        =   17
         Top             =   480
         Width           =   1830
         _ExtentX        =   3239
         _ExtentY        =   550
         _Version        =   393216
         MaxLength       =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6690
         TabIndex        =   13
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label lblTfinanc 
         AutoSize        =   -1  'True
         Caption         =   "IAFA del Paciente"
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
         Left            =   8430
         TabIndex        =   12
         Top             =   615
         Width           =   1455
      End
      Begin VB.Label lblNcuenta 
         Alignment       =   1  'Right Justify
         Caption         =   "N° Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3810
         TabIndex        =   11
         Top             =   885
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "N° Documento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   $"ucFarmVentasLista.ctx":5CFE
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   8145
      End
   End
   Begin UltraGrid.SSUltraGrid grdLista 
      Height          =   4560
      Left            =   0
      TabIndex        =   6
      Top             =   1860
      Width           =   12705
      _ExtentX        =   22416
      _ExtentY        =   8043
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista "
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Ventas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   45
      Width           =   12750
   End
End
Attribute VB_Name = "ucFarmVentasLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para buscar una VENTA
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
Dim ml_idRegistroSeleccionado As Long
Dim ml_TipoBusqueda As sghTipoBusquedaPrestamoHistoria

Dim mo_Teclado As New sighEntidades.Teclado
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsTipoFinanciamiento As New ADODB.Recordset
Dim oRsBusqueda As New ADODB.Recordset
Dim oRsFuentesFinanciamiento As New ADODB.Recordset
Dim ml_IdTipoVentaSeleccionada As Long
Dim ml_idUsuario As Long
Dim lcSerie As String
Dim lcMensajeError As String
Dim lbBotonBuscar As Boolean
Dim lcSubTitulo As String

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdLista.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdLista.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoBusqueda(lValue As sghTipoBusquedaPrestamoHistoria)
    ml_TipoBusqueda = lValue
End Property
Property Get TipoBusqueda() As sghTipoBusquedaPrestamoHistoria
    TipoBusqueda = ml_TipoBusqueda
End Property
Property Let TipoVentaSeleccionada(lValue As Long)
    ml_IdTipoVentaSeleccionada = lValue
End Property
Property Get TipoVentaSeleccionada() As Long
    TipoVentaSeleccionada = ml_IdTipoVentaSeleccionada
End Property


Private Sub bntReporte_Click()
    On Error GoTo errores
    UserControl.MousePointer = 11
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    mo_ReglasReportes.ExportarRecordSetAexcel oRsBusqueda, "VENTAS", lcSubTitulo, "", 1, True, True
    Set mo_ReglasReportes = Nothing
errores:
    UserControl.MousePointer = 1
End Sub

Private Sub btnBuscar_Click()
    If CDate(UserControl.txtFinicio.Text) > CDate(UserControl.txtFfinal.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lbBotonBuscar = True
    RealizarBusqueda
    lbBotonBuscar = False
    Screen.MousePointer = vbDefault
    UserControl_Resize
End Sub

Public Sub RealizarBusqueda()
        If cmbAlmacenV.Text = "" Then
            If lbBotonBuscar = True Then
                MsgBox "Por favor elija el Almacén", vbInformation, "Busqueda"
                Exit Sub
            End If
        End If

        Dim lcFilter As String, lnTotal As Double, lnIdAlmacen As Long, ldFechaIni As Date, ldFechaFin As Date
        lnTotal = 0
        If optPreventa.Value = False Then
            lcFilter = ""
            Set oRsBusqueda = mo_ReglasFarmacia.FarmDevuelveCabeceraDeVentasOpreventa("D", Val(cmbAlmacenV.BoundText), _
                                    "S", CDate(Format(txtFinicio.Text & ":00", sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                                    CDate(Format(txtFfinal.Text & ":59", sighEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
'            If txtNDocumento.Text <> "" And txtNDocumento.Text <> lcSerie Then
'               lcFilter = "dalmacen='" & Trim(txtNDocumento.Text) & "'"
'            End If
            If mo_Teclado.TextoEsSoloNumeros(txtNcuenta.Text) Then
               lcFilter = "idCuentaAtencion=" & Val(txtNcuenta.Text)
            ElseIf txtNdocumento.Text <> "" And txtNdocumento.Text <> lcSerie Then
               lcFilter = "dalmacen='" & txtNdocumento.Text & "'"
            End If
            If chkSoloBoletas.Value = 1 Then
               If lcFilter = "" Then
                  lcFilter = "idPreventa>0"
               Else
                  lcFilter = lcFilter & " and idPreventa>0"
               End If
            ElseIf Val(cmbFuenteFinanciamiento.BoundText) > 0 Then
               If lcFilter = "" Then
                  lcFilter = "idFuenteFinanciamiento=" & cmbFuenteFinanciamiento.BoundText
               Else
                  lcFilter = lcFilter & " and idFuenteFinanciamiento=" & cmbFuenteFinanciamiento.BoundText
               End If
            End If
'            If Val(cmbFuenteFinanciamiento.BoundText) > 0 Then
'               If lcFilter = "" Then
'                  If Val(cmbFuenteFinanciamiento.BoundText) = 1 Then
'                     lcFilter = "idPreventa>0"
'                  Else
'                     lcFilter = "idFuenteFinanciamiento=" & cmbFuenteFinanciamiento.BoundText
'                  End If
'               Else
'                  If Val(cmbFuenteFinanciamiento.BoundText) = 1 Then
'                     lcFilter = lcFilter & " and idPreventa>0"
'                  Else
'                     lcFilter = lcFilter & " and idFuenteFinanciamiento=" & cmbFuenteFinanciamiento.BoundText
'                  End If
'               End If
'            End If
            
            
            If lcFilter <> "" Then
               oRsBusqueda.Filter = lcFilter
            End If
        Else
            lnIdAlmacen = Val(cmbAlmacenV.BoundText)
            ldFechaIni = CDate(txtFinicio.Text)
            ldFechaFin = CDate(txtFfinal.Text)
            Set oRsBusqueda = mo_ReglasFarmacia.FarmDevuelveCabeceraDeVentasOpreventa("P", lnIdAlmacen, "S", ldFechaIni, ldFechaFin)
            If mo_Teclado.TextoEsSoloNumeros(txtNdocumento.Text) Then
               oRsBusqueda.Filter = "movNumero=" & Val(txtNdocumento.Text)
            End If
        End If
        If oRsBusqueda.RecordCount > 0 Then
           oRsBusqueda.MoveFirst
           Do While Not oRsBusqueda.EOF
              If oRsBusqueda.Fields!idEstadoMovimiento = 1 Then
                 lnTotal = lnTotal + oRsBusqueda.Fields!Total
              End If
              oRsBusqueda.MoveNext
           Loop
        End If
        lblTotal.Caption = Format(lnTotal, "#,###,###.#0")
        Set grdLista.DataSource = oRsBusqueda
       ' mo_Apariencia.ConfigurarFilasBiColores grdLista, sighentidades.GrillaConFilasBicolor
        lcSubTitulo = "Almac: " & Trim(cmbAlmacenV.Text) & "  (Fechas: " & txtFinicio.Text & " al " & txtFfinal.Text & _
                      ") (Tipo:" & IIf(optPreventa.Value = True, optPreventa.Caption, optVenta.Caption) & ") " & _
                      IIf(txtNdocumento.Text = lcSerie, "", "(" & Label1.Caption & ": " & txtNdocumento.Text & ")") & _
                      IIf(txtNcuenta.Text = "", "", "(" & lblNcuenta.Caption & ": " & txtNcuenta.Text & ")") & _
                      IIf(cmbFuenteFinanciamiento.Text = "", "", "(" & lblTfinanc.Caption & ": " & cmbFuenteFinanciamiento.Text & ")") & _
                      IIf(chkSoloBoletas.Value = 0, "", "(" & chkSoloBoletas.Caption & ")")
                      
End Sub

Private Sub btnLimpiar_Click()
    LimpiarFiltro
End Sub
Public Sub LimpiarFiltro()
        txtFinicio.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 00:00"
        txtFfinal.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
        If optPreventa.Value = False Then
           UserControl.txtNdocumento.Text = lcSerie
           txtNcuenta.Text = ""
           cmbFuenteFinanciamiento.Visible = True
           lblTfinanc.Visible = True
           chkSoloBoletas.Visible = True
           
        Else
           UserControl.txtNdocumento.Text = ""
           txtNcuenta.Text = ""
           cmbFuenteFinanciamiento.Visible = False
           lblTfinanc.Visible = False
           chkSoloBoletas.Visible = False
        End If
        cmbFuenteFinanciamiento.BoundText = ""
        chkSoloBoletas.Value = 0
End Sub







Private Sub chkSoloBoletas_Click()
    If chkSoloBoletas.Value = 1 Then
       cmbFuenteFinanciamiento.Text = ""
    End If
End Sub

Private Sub cmbAlmacenV_Click(Area As Integer)
   lcMensajeError = ""
End Sub

Private Sub cmbAlmacenV_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacenV
    AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbFuenteFinanciamiento_Click(Area As Integer)
     chkSoloBoletas.Value = 0
End Sub


Private Sub cmbTipoFinanciamiento_Click()
     cmbFuenteFinanciamiento.Text = ""
End Sub

Private Sub grdLista_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = Val(rsRecordset("MovNumero"))
    
End Sub

Private Sub grdLista_Click()
Dim rsRecordset As ADODB.Recordset

    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdLista.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = Val(rsRecordset("MovNumero"))
    
End Sub


Private Sub grdLista_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    
    grdLista.Bands(0).Columns("idTipoFinanciamiento").Hidden = True
    grdLista.Bands(0).Columns("IdEstadoMovimiento").Hidden = True
    If optPreventa.Value = True Then
       'debb-16/02/2011
       grdLista.Bands(0).Columns("MovNumero1").Header.Caption = "N° Documento"
       grdLista.Bands(0).Columns("MovNumero").Hidden = False
       grdLista.Bands(0).Columns("MovNumero").Header.Caption = "N° Documento"
       grdLista.Bands(0).Columns("Dalmacen").Header.Caption = "Almacén"
       'debb-16/02/2011
    Else
       grdLista.Bands(0).Columns("MovNumero").Hidden = True
       grdLista.Bands(0).Columns("MovNumero").Header.Caption = ""
       grdLista.Bands(0).Columns("Dalmacen").Header.Caption = "N° Documento"
       grdLista.Bands(0).Columns("Paciente").Width = 2400
       grdLista.Bands(0).Columns("NroHistoriaClinica").Width = 900
    End If
    grdLista.Bands(0).Columns("Estado").Width = 1000
    grdLista.Bands(0).Columns("MovNumero").Width = 900
    grdLista.Bands(0).Columns("Dalmacen").Width = 1000
    grdLista.Bands(0).Columns("Descripcion").Header.Caption = "Producto/Plan"
    grdLista.Bands(0).Columns("Descripcion").Width = 1000
    grdLista.Bands(0).Columns("fechaCreacion").Header.Caption = "Fecha"
    grdLista.Bands(0).Columns("fechaCreacion").Width = 1000
    grdLista.Bands(0).Columns("fechaCreacion").Format = "dd/mm/yyyy hh:mm:ss"
    grdLista.Bands(0).Columns("idCuentaAtencion").Header.Caption = "N° Cuenta"
    grdLista.Bands(0).Columns("idCuentaAtencion").Width = 1000
    grdLista.Bands(0).Columns("Total").Width = 1100
    grdLista.Bands(0).Columns("Total").Format = "#0.00"

End Sub







Private Sub grdLista_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoMovimiento").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
        End If

End Sub

Private Sub optPreventa_Click(Value As Integer)
   If optPreventa.Value = True Then
      ml_IdTipoVentaSeleccionada = 1
      txtNdocumento.Text = ""
      lblNcuenta.Visible = False
      txtNcuenta.Visible = False
      cmbFuenteFinanciamiento.Visible = False
      lblTfinanc.Visible = False
      btnBuscar_Click
   End If
End Sub

Private Sub optVenta_Click(Value As Integer)
    If optVenta.Value = True Then
       ml_IdTipoVentaSeleccionada = 0
       txtNdocumento.Text = lcSerie
       lblNcuenta.Visible = True
       txtNcuenta.Visible = True
       cmbFuenteFinanciamiento.Visible = True
       lblTfinanc.Visible = True
       btnBuscar_Click
    End If
End Sub






Private Sub txtFfinal_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFfinal
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtFfinal_LostFocus()
    If Not IsDate(txtFfinal.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFfinal.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If
End Sub

Private Sub txtFinicio_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtFinicio
    AdministrarKeyPreview KeyCode

End Sub





Private Sub txtFinicio_LostFocus()
    If Not IsDate(txtFinicio.Text) Then
        MsgBox "La fecha ingresada no es válida", vbInformation, ""
        On Error Resume Next
        txtFinicio.Text = sighEntidades.FECHA_VACIA_DMY
        Exit Sub
    End If

End Sub

Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtNcuenta.Text <> "" Then
       btnBuscar_Click
    End If

End Sub

Private Sub txtNDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtNdocumento.Text <> "" Then
       btnBuscar_Click
    End If
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
   
   fraBusqueda.Width = UserControl.Width - 110
   lblNombre.Width = UserControl.Width
   grdLista.Width = fraBusqueda.Width
   grdLista.Height = UserControl.Height - UserControl.lblTotal.Height - 1600
End Sub



Sub CargaComboBox()
        On Error GoTo ErrFarm
        Dim rsIdAlmacen As Recordset
        Dim lcSql As String
        Dim oRsTmp As New Recordset
        Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
        Set oBuscaDondeLabora = Nothing
        Set oRsAlmacenes = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1")
        Set cmbAlmacenV.RowSource = Nothing
        'SCCQ 02/06/2020 Cambio23  Inicio
        If rsIdAlmacen.RecordCount > 0 Then 'Solo filtra farmacias asignadas
         Set cmbAlmacenV.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1 and idAlmacen in (select idLaboraSubArea from EmpleadosLugarDeTrabajo where idLaboraArea=" + CStr(sghAlmacenFarmacia) + " and idEmpleado=" + CStr(ml_idUsuario) + ")")
        Else 'Muestra todas las farmacias como lo hacía antes
            Set cmbAlmacenV.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='F' and idEstado=1") 'oRsAlmacenes
        End If
        'SCCQ 02/06/2020 Cambio23  Fin
        cmbAlmacenV.ListField = "descripcion"
        cmbAlmacenV.BoundColumn = "idAlmacen"
        'SCCQ 02/06/2020 Cambio23  Inicio
        If rsIdAlmacen.RecordCount = 1 Then
           cmbAlmacenV.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
           'cmbAlmacenV.Enabled = False
        End If
        'SCCQ 02/06/2020 Cambio23  Fin
        Set oRsFuentesFinanciamiento = mo_ReglasComunes.FuentesFinanciamientoSeleccionarTodos
        Set cmbFuenteFinanciamiento.RowSource = oRsFuentesFinanciamiento
        cmbFuenteFinanciamiento.ListField = "Descripcion"
        cmbFuenteFinanciamiento.BoundColumn = "idFuenteFinanciamiento"
        '
        Set oRsTmp = mo_ReglasCaja.CajaTiposComprobanteFarmacia
        lcSerie = ""
        If oRsTmp.RecordCount > 0 Then
           lcSerie = oRsTmp.Fields!nroSerie & "-"
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
        txtNdocumento.Text = lcSerie
        lbBotonBuscar = False
ErrFarm:
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighEntidades.Parametro282valorInt = "1" Then
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdLista, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdLista, sighEntidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Sub Inicializar()
    SkinConfigura
    CargaComboBox
    txtFinicio.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 00:00"
    txtFfinal.Text = Format(Date, sighEntidades.DevuelveFechaSoloFormato_DMY) & " 23:59"
    ml_IdTipoVentaSeleccionada = 0
    
    optVenta.Value = True
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
        btnBuscar_Click
     Case vbKeyF7
        btnLimpiar_Click
     Case vbKeyF8
    End Select
       
End Sub







