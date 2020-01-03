VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.UserControl UcHISCalidad 
   ClientHeight    =   9870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17475
   ScaleHeight     =   9870
   ScaleWidth      =   17475
   Begin VB.Frame frmGenerados 
      Caption         =   "REGISTROS PARA LA DOBLE DIGITACIÓN DEL LOTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   5160
      TabIndex        =   5
      Top             =   480
      Width           =   12255
      Begin VB.CommandButton btnGenerar 
         DisabledPicture =   "UcHISCalidad.ctx":0000
         DownPicture     =   "UcHISCalidad.ctx":03E9
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
         Left            =   5160
         Picture         =   "UcHISCalidad.ctx":07F5
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Generar Registros Aleatorios (F7)"
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtTotalHojas 
         Alignment       =   2  'Center
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
         Left            =   10170
         MaxLength       =   3
         TabIndex        =   13
         Top             =   300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtEstablecimientoSeleccionado 
         Alignment       =   2  'Center
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtNroRegistros 
         Alignment       =   2  'Center
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
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   8
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
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
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   7
         Top             =   600
         Width           =   885
      End
      Begin UltraGrid.SSUltraGrid grdGenerados 
         Height          =   8235
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   14526
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Registros aleatorios"
      End
      Begin VB.Label Label2 
         Caption         =   "(F7)"
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
         Index           =   0
         Left            =   5760
         TabIndex        =   17
         Top             =   600
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Total Hojas"
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
         Index           =   2
         Left            =   10170
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Establecimiento"
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
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label2 
         Caption         =   "Total Registros"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   11
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
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
         Index           =   0
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Búsqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9285
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5055
      Begin VB.CommandButton btnBuscar 
         Height          =   315
         Left            =   3660
         Picture         =   "UcHISCalidad.ctx":0C01
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   510
         Width           =   1305
      End
      Begin VB.ComboBox cmbEstablecimiento 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   520
         Width           =   3495
      End
      Begin UltraGrid.SSUltraGrid grdCalidad 
         Height          =   8265
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   14579
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lotes a verificar"
      End
      Begin VB.Label Label1 
         Caption         =   "Establecimiento"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calidad de Lotes"
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
      TabIndex        =   0
      Top             =   0
      Width           =   17535
   End
End
Attribute VB_Name = "UcHISCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para CALIDAD HIS
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_LoginPC As String
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_idLoteSeleccionado As Long
Dim ml_Registrado As Integer
Dim lblNombre As String
Dim ml_IdEstablecimiento As Long: Dim ml_IdEstablecimientoSeleccionado As Long
Dim ml_idRegistroSeleccionado As Long
Dim ml_idUsuario As Long
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mo_DatosParametros As New SIGHDatos.Parametros
Dim oRcs_DatosEstablecimiento As New ADODB.Recordset
Dim oRcs_HisLotes As New ADODB.Recordset
Dim oListaEstablecimientos As ADODB.Recordset
Dim mo_cmbEstablecimiento As New sighEntidades.ListaDespleglable
Dim mo_Formulario As New sighEntidades.Formulario

'========================================= PROPIEDADES ==============================
Property Set DataSource(oValue As ADODB.Recordset)
    Set grdCalidad.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = grdCalidad.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property

Property Let Registrado(lValue As Integer)
    ml_Registrado = lValue
End Property
Property Get Registrado() As Integer
    Registrado = ml_Registrado
End Property

Property Let IdEstablecimiento(lValue As Long)
    ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
    IdEstablecimiento = ml_IdEstablecimiento
End Property
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property
Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property

Private Sub btnAceptar_Click()
Dim oTablaDOHIS_Lote As New DOHIS_Lotes
    If oRcs_HisLotes.RecordCount > 0 Then
        oRcs_HisLotes.MoveFirst
        Do While Not oRcs_HisLotes.EOF
            If oRcs_HisLotes.Fields!Revisado = True Then
                oTablaDOHIS_Lote.IdEstadoLote = 2
                oTablaDOHIS_Lote.IdHisLote = oRcs_HisLotes.Fields!IdHisLote
                If mo_ReglasHIS.ModificarRegistroLoteHIS(oTablaDOHIS_Lote) Then
                End If
            End If
            oRcs_HisLotes.MoveNext
        Loop
        oRcs_HisLotes.MoveFirst
    End If
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    RealizarBusqueda
    Screen.MousePointer = vbDefault
End Sub

Public Sub RealizarBusqueda()
    If mo_cmbEstablecimiento.BoundText = "" Then
        Exit Sub
    End If
    
    Dim oRcs_Consultas As New Recordset
    Dim oRcsLotes As New Recordset
    With oRcsLotes
        .Fields.Append "IdHisLote", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "IdEstablecimiento", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Nombre", adVarChar, 200, adFldIsNullable + adFldUpdatable
        .Fields.Append "Anio", adVarChar, 4, adFldIsNullable + adFldUpdatable
        .Fields.Append "Mes", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .Fields.Append "Lote", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .Fields.Append "NroHojas", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .Fields.Append "HojasRegistradas", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .Fields.Append "TotalRegistros", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .Fields.Append "Estado", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .Fields.Append "DobleDigitacion", adVarChar, 20, adFldIsNullable + adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With

    
    ml_idLoteSeleccionado = -1
    ml_IdEstablecimientoSeleccionado = 0
    UserControl.txtEstablecimientoSeleccionado.Text = ""
    UserControl.txtLote.Text = ""
    UserControl.txtTotalHojas.Text = ""
    UserControl.txtNroRegistros.Text = ""
    
    Set oRcs_Consultas = mo_ReglasHIS.His_LotesConsultarFiltro(mo_cmbEstablecimiento.BoundText, 3)
    If oRcs_Consultas.RecordCount <> 0 Then
        oRcs_Consultas.MoveFirst
        Do While Not oRcs_Consultas.EOF
            oRcsLotes.AddNew
            oRcsLotes.Fields!IdHisLote = oRcs_Consultas!IdHisLote
            oRcsLotes.Fields!IdEstablecimiento = oRcs_Consultas!IdEstablecimiento
            oRcsLotes.Fields!Nombre = oRcs_Consultas!Nombre
            oRcsLotes.Fields!Anio = oRcs_Consultas!Anio
            oRcsLotes.Fields!Mes = oRcs_Consultas!Mes
            oRcsLotes.Fields!Lote = oRcs_Consultas!Lote
            oRcsLotes.Fields!NroHojas = oRcs_Consultas!NroHojas
            oRcsLotes.Fields!HojasRegistradas = mo_ReglasHIS.His_ConsultarHojasRegistradas(mo_cmbEstablecimiento.BoundText, oRcs_Consultas!IdHisLote).RecordCount
            oRcsLotes.Fields!TotalRegistros = mo_ReglasHIS.His_ConsultarTotalRegistrosLote(mo_cmbEstablecimiento.BoundText, oRcs_Consultas!IdHisLote).RecordCount
            oRcsLotes.Fields!estado = oRcs_Consultas!estado
            oRcsLotes.Fields!DobleDigitacion = oRcs_Consultas!DobleDigitacion
            oRcsLotes.Update
            oRcs_Consultas.MoveNext
        Loop
        oRcs_Consultas.MoveFirst
    End If
    Set grdCalidad.DataSource = oRcsLotes
    mo_Apariencia.ConfigurarFilasBiColores grdCalidad, sighEntidades.GrillaConFilasBicolor
    If mo_ReglasHIS.MensajeError <> "" Then
        MsgBox mo_ReglasHIS.MensajeError, vbCritical, "Filtro Registros HIS"
    End If
End Sub


Private Sub cmbEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub grdCalidad_Click()
    CargarListaGenerados
End Sub

Public Sub CargarListaGenerados()
    Dim rsRecordset As ADODB.Recordset
    Dim rcsTemp As ADODB.Recordset
    
    'Limpiamos la grilla
    Set rcsTemp = mo_ReglasHIS.HIS_ConsultarRegMuestraLotes(0)
    Set grdGenerados.DataSource = rcsTemp
    mo_Apariencia.ConfigurarFilasBiColores grdGenerados, sighEntidades.GrillaConFilasBicolor
        
    ml_idLoteSeleccionado = -1
    ml_IdEstablecimientoSeleccionado = 0
    UserControl.txtEstablecimientoSeleccionado.Text = ""
    UserControl.txtLote.Text = ""
    UserControl.txtTotalHojas.Text = ""
    UserControl.txtNroRegistros.Text = ""
    UserControl.btnGenerar.Enabled = False
    Set rsRecordset = grdCalidad.DataSource
    On Error Resume Next
    If rsRecordset.RecordCount = 0 Then Exit Sub
    ml_idLoteSeleccionado = IIf(IsNull(rsRecordset("IdHisLote")), -1, rsRecordset("IdHisLote"))
    If ml_idLoteSeleccionado = -1 Then Exit Sub
    Dim oTablaDOHIS_Lote As New DOHIS_Lotes
    oTablaDOHIS_Lote.IdHisLote = ml_idLoteSeleccionado
    Set oTablaDOHIS_Lote = mo_ReglasHIS.ConsultarRegistroLoteHIS(oTablaDOHIS_Lote)
    UserControl.txtEstablecimientoSeleccionado.Text = rsRecordset!Nombre
    ml_IdEstablecimientoSeleccionado = rsRecordset!IdEstablecimiento
    UserControl.txtLote.Text = rsRecordset!Lote
    UserControl.txtTotalHojas.Text = rsRecordset!HojasRegistradas
    UserControl.txtNroRegistros.Text = rsRecordset!TotalRegistros
    If oTablaDOHIS_Lote.DobleDigitacion = 0 Then
        UserControl.btnGenerar.Enabled = True
    Else
        Set rcsTemp = mo_ReglasHIS.HIS_ConsultarRegMuestraLotes(ml_idLoteSeleccionado)
        Set grdGenerados.DataSource = rcsTemp
        mo_Apariencia.ConfigurarFilasBiColores grdGenerados, sighEntidades.GrillaConFilasBicolor
    End If
End Sub

Private Sub grdCalidad_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With grdCalidad.Bands(0)
        .Columns("IdHisLote").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("Nombre").Hidden = True
        .Columns("Anio").Hidden = True
        .Columns("Mes").Hidden = True
        .Columns("Lote").Header.Caption = "Lote"
        .Columns("Lote").Width = 600
        .Columns("Lote").Activation = ssActivationActivateNoEdit
        .Columns("NroHojas").Hidden = True
        .Columns("HojasRegistradas").Header.Caption = "Hojas registradas"
        .Columns("HojasRegistradas").Width = 1520
        .Columns("HojasRegistradas").Activation = ssActivationActivateNoEdit
        .Columns("TotalRegistros").Header.Caption = "Total Registros"
        .Columns("TotalRegistros").Width = 1320
        .Columns("TotalRegistros").Activation = ssActivationActivateNoEdit
        .Columns("Estado").Header.Caption = "Estado"
        .Columns("Estado").Width = 1300
        .Columns("Estado").Activation = ssActivationActivateNoEdit
        .Columns("DobleDigitacion").Hidden = True
    End With
End Sub

'=========================================== METODOS ================================
Public Function inicializar()
    ml_IdEstablecimientoSeleccionado = 0
    ml_idLoteSeleccionado = -1
    mo_Formulario.HabilitarDeshabilitar UserControl.txtEstablecimientoSeleccionado, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtLote, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtTotalHojas, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtNroRegistros, False
    mo_Formulario.HabilitarDeshabilitar UserControl.txtNroRegistros, False
    UserControl.btnGenerar.Enabled = False
    Set mo_cmbEstablecimiento.MiComboBox = UserControl.cmbEstablecimiento
    CargarComboBoxes
'    RealizarBusqueda
End Function

Private Sub CargarComboBoxes()
    Dim orsTemp As New ADODB.Recordset
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set orsTemp = ListadoEstablecimientos()
    Set mo_cmbEstablecimiento.RowSource = orsTemp
    If orsTemp.RecordCount = 0 Then
        MsgBox "No tiene establecimientos ni servicios configurados", vbExclamation, "HIS"
    End If
    If orsTemp.RecordCount > 0 Then
        cmbEstablecimiento.ListIndex = 0
    End If
    Set orsTemp = Nothing
End Sub

Private Function ListadoEstablecimientos() As Recordset
    Dim oTabla As New DOEstablecimiento
    Set oRcs_DatosEstablecimiento = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    If oRcs_DatosEstablecimiento.RecordCount > 0 Then
        oRcs_DatosEstablecimiento.MoveFirst
    End If
    Set ListadoEstablecimientos = oRcs_DatosEstablecimiento
End Function


Private Sub grdCalidad_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub grdGenerados_AfterRowActivate()
    Dim rsRecordset As ADODB.Recordset
    ml_idRegistroSeleccionado = -1
    ml_Registrado = -1
    Set rsRecordset = grdGenerados.DataSource
    On Error Resume Next
    ml_idRegistroSeleccionado = IIf(IsNull(rsRecordset("IdHisDetalle")), -1, rsRecordset("IdHisDetalle"))
    ml_Registrado = IIf(IsNull(rsRecordset("Registrado")), -1, rsRecordset("Registrado"))
End Sub

Private Sub grdGenerados_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
With grdGenerados.Bands(0)
        .Columns("NroRegistroLote").Header.Caption = "ID"
        .Columns("NroRegistroLote").Width = 800
        .Columns("NroRegistroLote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisLote").Hidden = True
        .Columns("Lote").Header.Caption = "Lote"
        .Columns("Lote").Width = 1400
        .Columns("Lote").Activation = ssActivationActivateNoEdit
        .Columns("IdHisCabecera").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("NroHojaHis").Header.Caption = "Nro Hoja"
        .Columns("NroHojaHis").Width = 1400
        .Columns("NroHojaHis").Activation = ssActivationActivateNoEdit
        .Columns("IdHisDetalle").Hidden = True
        .Columns("NroRegistroHoja").Header.Caption = "Registro Hoja"
        .Columns("NroRegistroHoja").Width = 1400
        .Columns("NroRegistroHoja").Activation = ssActivationActivateNoEdit
        .Columns("DiaAtencion").Header.Caption = "Día"
        .Columns("DiaAtencion").Width = 1400
        .Columns("DiaAtencion").Activation = ssActivationActivateNoEdit
        .Columns("IdTipoAtencion").Hidden = True
        .Columns("HC_FF_COD").Header.Caption = "HC_FF_COD"
        .Columns("HC_FF_COD").Width = 1400
        .Columns("HC_FF_COD").Activation = ssActivationActivateNoEdit
        .Columns("IdPais").Hidden = True
        .Columns("Codigo").Hidden = True
        .Columns("IdTipoDocumento").Hidden = True
        .Columns("Documento").Hidden = True
        .Columns("NroDocIdentidad").Hidden = True
        .Columns("NroHijo").Hidden = True
        .Columns("IdTipoFinanciamiento").Hidden = True
        .Columns("Financiamiento").Hidden = True
        .Columns("IdEtnia").Hidden = True
        .Columns("Etnia").Hidden = True
        .Columns("IdDistrito").Hidden = True
        .Columns("Distrito").Hidden = True
        .Columns("Edad").Hidden = True
        .Columns("IdTipoEdad").Hidden = True
        .Columns("TipoEdad").Hidden = True
        .Columns("Sexo").Hidden = True
        .Columns("Peso").Hidden = True
        .Columns("Talla").Hidden = True
        .Columns("IdEstadoaEstablec").Hidden = True
        .Columns("IdEstadoaServicio").Hidden = True
        .Columns("Registrado").Hidden = True
        .Columns("Coincide").Hidden = True
        .Columns("EsRegistrado").Header.Caption = "Registrado"
        .Columns("EsRegistrado").Width = 2000
        .Columns("EsRegistrado").Activation = ssActivationActivateNoEdit
    End With
End Sub

Private Sub txtEstablecimientoSeleccionado_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtNroRegistros_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtTotalHojas_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
    UserControl.lblNombre.Width = UserControl.Width - 200
    fraBusqueda.Height = UserControl.Height - 510
    grdCalidad.Height = fraBusqueda.Height - 1060
    frmGenerados.Left = fraBusqueda.Left + fraBusqueda.Width + 50
    frmGenerados.Height = fraBusqueda.Height
    frmGenerados.Width = UserControl.Width - fraBusqueda.Width - 200
    UserControl.grdGenerados.Width = frmGenerados.Width - 200
    grdGenerados.Height = grdCalidad.Height
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
         GenerarRegistrosAleatorios
     Case vbKeyF8
    End Select
End Sub

Private Sub btnGenerar_Click()
    GenerarRegistrosAleatorios
    grdCalidad_Click
End Sub

Sub GenerarRegistrosAleatorios()
    If UserControl.btnGenerar.Enabled Then
        If ml_idLoteSeleccionado = -1 Then Exit Sub
        Dim mo_HISLotes As New SIGHhisDigitacion.MantGenRegAleLotes
        mo_HISLotes.Opcion = sghAgregar
        mo_HISLotes.idUsuario = ml_idUsuario
        mo_HISLotes.IdRegistroLote = ml_idLoteSeleccionado
        mo_HISLotes.IdEstablecimiento = ml_IdEstablecimientoSeleccionado
        mo_HISLotes.MostrarFormulario
    End If
End Sub
