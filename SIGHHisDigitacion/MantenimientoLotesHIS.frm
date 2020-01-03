VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form MantenimientoLotesHIS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HIS - Lotes"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "MantenimientoLotesHIS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6855
      Begin VB.ComboBox cmbEstadoLote 
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
         TabIndex        =   8
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtTotalRegistros 
         Height          =   330
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox chkPasarConcluido 
         Caption         =   "¿Se concluyó la digitación?"
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
         Top             =   1800
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox txtHojasDigitadas 
         Height          =   330
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
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
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtNroPag 
         Height          =   330
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cmbMes 
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
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtLote 
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
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskfechaAnio 
         Height          =   330
         Left            =   5400
         TabIndex        =   3
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMensajeLote 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label7 
         Caption         =   "Total de registros"
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
         Left            =   4680
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Hojas digitadas"
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
         Left            =   3240
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Total de Hojas"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   840
         Width           =   1215
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Mes"
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
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   6855
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MantenimientoLotesHIS.frx":000C
         DownPicture     =   "MantenimientoLotesHIS.frx":046C
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
         Left            =   1800
         Picture         =   "MantenimientoLotesHIS.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MantenimientoLotesHIS.frx":0D56
         DownPicture     =   "MantenimientoLotesHIS.frx":121A
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
         Left            =   3360
         Picture         =   "MantenimientoLotesHIS.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "MantenimientoLotesHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica en donde se ingresara los lotes del HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_DatosParametros As New SIGHDatos.Parametros
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_cmbEstablecimiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim mo_cmbEstado As New SIGHEntidades.ListaDespleglable
Dim oTablaDOHIS_Lote As New DOHIS_Lotes
Dim mi_Opcion As sghOpciones
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_IdEstablecimiento As Long
Dim ml_NroTotalHojas As Integer
Dim ml_HojasRegistradas As Integer
Dim ml_IdUsuario As Long
Dim ml_IdLote As Long
Dim ms_fechaactual As String
Dim lcBuscaParametro As New SIGHDatos.Parametros

'========================================== PROPIEDADES ========================================
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdLote() As Long
    IdLote = ml_IdLote
End Property
Property Let IdLote(lValue As Long)
   ml_IdLote = lValue
End Property

Private Sub chkPasarConcluido_Click()
    If oTablaDOHIS_Lote.IdEstadoLote = 0 Then
        If Val(Me.txtNroPag.Text) = Val(Me.txtHojasDigitadas.Text) Then
            If Me.chkPasarConcluido.Value Then
                  Call MsgBox("Al cambiar el Lote al estado CONCLUIDO, las hojas no podrán ser editadas", vbExclamation, Me.Caption)
'                 If MsgBox("¿Esta seguro de pasar el lote al estado 'Concluido'?", vbYesNo, Me.Caption) = vbNo Then
'                    Me.chkPasarConcluido.Value = False
'                 End If
            End If
        End If
    End If
End Sub

Private Sub cmbEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.cmbMes
    AdministrarKeyPreview KeyCode
End Sub

'========================================== EVENTOS ========================================
Private Sub Form_Load()
    Dim oRcsTemp As ADODB.Recordset
    Set mo_cmbMes.MiComboBox = Me.cmbMes
    Set mo_cmbEstablecimiento.MiComboBox = Me.cmbEstablecimiento
    Set mo_cmbEstado.MiComboBox = Me.cmbEstadoLote
    ml_NroTotalHojas = 0
    CargarComboBoxes
    mo_Formulario.HabilitarDeshabilitar Me.txtHojasDigitadas, False
    mo_Formulario.HabilitarDeshabilitar Me.txtTotalRegistros, False
    mo_Formulario.HabilitarDeshabilitar Me.cmbEstadoLote, False
    Select Case mi_Opcion
        Case sghOpciones.sghAgregar
            Me.Caption = "Ingresar Lote"
            mo_cmbEstado.BoundText = 0
        Case sghOpciones.sghModificar, sghOpciones.sghConsultar, sghOpciones.sghEliminar
            Me.Caption = "Modificar Lote"
            oTablaDOHIS_Lote.IdHisLote = ml_IdLote
            Set oTablaDOHIS_Lote = mo_ReglasHIS.ConsultarRegistroLoteHIS(oTablaDOHIS_Lote)
            ml_NroTotalHojas = mo_ReglasHIS.ObtenerDatosLoteNroHojaLibre(ml_IdLote)
            mo_cmbEstado.BoundText = oTablaDOHIS_Lote.IdEstadoLote
            Set oRcsTemp = mo_ReglasHIS.His_ConsultarHojasRegistradas(mo_cmbEstablecimiento.BoundText, ml_IdLote)
            Me.txtHojasDigitadas.Text = oRcsTemp.RecordCount
            If CInt(Me.txtHojasDigitadas.Text) > 0 Then mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
            Set oRcsTemp = mo_ReglasHIS.His_ConsultarTotalRegistrosLote(mo_cmbEstablecimiento.BoundText, ml_IdLote)
            Me.txtTotalRegistros.Text = oRcsTemp.RecordCount
            mo_Formulario.HabilitarDeshabilitar Me.cmbEstablecimiento, False
            mo_Formulario.HabilitarDeshabilitar Me.cmbMes, False
            mo_Formulario.HabilitarDeshabilitar Me.mskfechaAnio, False
            If mi_Opcion = sghConsultar Or mi_Opcion = sghEliminar Then
                mo_Formulario.HabilitarDeshabilitar Me.txtNroPag, False
                mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
                Me.btnAceptar.Enabled = False
                If mi_Opcion = sghEliminar And CInt(Me.txtHojasDigitadas.Text) = 0 Then Me.btnAceptar.Enabled = True
            End If
    End Select
    CargarDatosAlFormulario
End Sub

Private Sub btnAceptar_Click()
    Select Case mi_Opcion
        Case sghOpciones.sghAgregar
        If ValidarDatosObligatorios Then
            If ValidarReglas Then
                If AgregarDatos Then
                Call MsgBox("El Lote fue registrado satisfactoriamente.", vbInformation, Me.Caption)
                Me.Hide
                Else
                Call MsgBox("No se pudo registrar el lote, Verificar Error.", vbExclamation, Me.Caption)
                Exit Sub
                End If
            End If
        End If
        Case sghOpciones.sghModificar
        If ValidarDatosObligatorios Then
            If ValidarReglas Then
                If ModificarDatos Then
                Call MsgBox("El lote fue modificado satisfactoriamente.", vbInformation, Me.Caption)
                Me.Hide
                Else
                Call MsgBox("No se pudo modificar el lote, Verificar Error.", vbExclamation, Me.Caption)
                Exit Sub
                End If
            End If
        End If
        Case sghOpciones.sghEliminar
        If ValidarReglas Then
            If EliminarDatos Then
            Call MsgBox("El lote fue eliminado satisfactoriamente.", vbInformation, Me.Caption)
            Me.Hide
            Else
            Call MsgBox("No se pudo eliminar el lote, Verificar Error.", vbExclamation, Me.Caption)
            Exit Sub
            End If
        End If
    End Select
End Sub
Private Sub btnCancelar_Click()
    Me.Hide
End Sub

Private Sub txtNroPag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
            Me.btnAceptar.SetFocus
        End If
    Else
        AdministrarKeyPreview KeyCode
    End If
End Sub

Private Sub txtNroPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If mi_Opcion <> sghOpciones.sghAgregar Then
        If oTablaDOHIS_Lote.IdEstadoLote = 0 Then
            If Val(Me.txtNroPag.Text) > Val(Me.txtHojasDigitadas.Text) Then
                Me.chkPasarConcluido.Value = False
                Me.chkPasarConcluido.Visible = False
            Else
                If Val(Me.txtNroPag.Text) = Val(Me.txtHojasDigitadas.Text) Then
                    Me.chkPasarConcluido.Visible = True
                End If
            End If
        End If
    End If
End Sub

Private Sub txtTotalRegistros_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48) Or KeyAscii > 57 Then
    If KeyAscii = 8 Then
        KeyAscii = 8
    Else
        KeyAscii = 1
    End If
End If
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, cmbMes
End Sub

Private Sub cmbMes_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, mskfechaAnio
End Sub

Private Sub mskfechaAnio_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
    mo_Teclado.RealizarNavegacion KeyCode, btnAceptar
End Sub

'========================================== METODOS ========================================
Sub CargarDatosAlFormulario()
    ms_fechaactual = mo_DatosParametros.RetornaFechaServidorSQL
    If mi_Opcion = sghAgregar Then
        Me.mskfechaAnio.Text = CStr(Year(CDate(ms_fechaactual)))
        mo_cmbMes.BoundText = CStr(Month(CDate(ms_fechaactual)))
    Else
        Me.txtLote.Text = oTablaDOHIS_Lote.Lote
        Me.txtNroPag.Text = CStr(oTablaDOHIS_Lote.NroHojas)
        mo_cmbMes.BoundText = oTablaDOHIS_Lote.Mes
        Me.mskfechaAnio.Text = oTablaDOHIS_Lote.Anio
        
        If oTablaDOHIS_Lote.IdEstadoLote = 0 Then 'Digitandose'
            If Me.txtNroPag.Text = Me.txtHojasDigitadas.Text Then
                Me.chkPasarConcluido.Visible = True
                Me.chkPasarConcluido.Caption = "¿Se concluyó la digitación?"
            End If
        Else
            mo_Formulario.HabilitarDeshabilitar Me.txtLote, False
            mo_Formulario.HabilitarDeshabilitar Me.txtNroPag, False
            Select Case oTablaDOHIS_Lote.IdEstadoLote
                Case 1 'Concluido'
                    Me.lblMensajeLote.Caption = "El Lote esta listo para ser ENVIADO"
                    Me.btnAceptar.Enabled = False
                Case 2 'Enviado'
                    Me.lblMensajeLote.Caption = "El Lote ya fue ENVIADO"
                    Me.btnAceptar.Enabled = False
                Case 3 'Por Verificar'
                    Dim rcsTemp As Recordset
                    Dim lbTodoRegistrado As Boolean
                    Me.btnAceptar.Enabled = False
                    Set rcsTemp = mo_ReglasHIS.HIS_ConsultarRegMuestraLotes(oTablaDOHIS_Lote.IdHisLote)
                    If rcsTemp.RecordCount > 0 Then
                        lbTodoRegistrado = True
                        rcsTemp.MoveFirst
                        Do While Not rcsTemp.EOF
                            If rcsTemp.Fields!Registrado <> 1 Then
                                lbTodoRegistrado = False
                                Exit Do
                            End If
                            rcsTemp.MoveNext
                        Loop
                        If lbTodoRegistrado = True Then
                            Me.lblMensajeLote.Caption = "Se completó la doble digitación"
                            Me.chkPasarConcluido.Visible = True
                            Me.chkPasarConcluido.Caption = "¿Comparar los registros de la doble digitación?"
                            Me.btnAceptar.Enabled = True
                        End If
                    End If
                Case 4 'Concluido Verificado'
                    Me.lblMensajeLote.Caption = "El Lote esta listo para ser ENVIADO"
                    Me.btnAceptar.Enabled = False
                Case 5 'Si falla veriificado'
                    Me.lblMensajeLote.Caption = "Se recomienda pasar al estado 'Digitándose' para revisar las hojas"
                    Me.chkPasarConcluido.Visible = True
                    Me.chkPasarConcluido.Caption = "Volver al estado 'Digitandose'"
            End Select
        End If
    End If
End Sub

Sub CargarComboBoxes()
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set mo_cmbEstablecimiento.RowSource = mo_ReglasHIS.ObtenerListaEstablecimientosMR
    mo_cmbEstablecimiento.BoundText = ml_IdEstablecimiento
    
    mo_cmbEstado.BoundColumn = "idestado"
    mo_cmbEstado.ListField = "descripcion"
    Set mo_cmbEstado.RowSource = mo_ReglasHIS.HIS_ConsultarEstadosLote()
   
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mo_ReglasHIS.ListaMeses
End Sub

Private Function ListadoEstablecimientos() As Recordset
    Dim oTabla As New DOEstablecimiento
    Dim oRcs_Establecimiento As New Recordset
    Set ListadoEstablecimientos = mo_ReglasHIS.ObtenerListaEstablecimientosMR
End Function

Function ValidarDatosObligatorios() As Boolean
    On Error Resume Next
    ValidarDatosObligatorios = False
    If Trim(Me.txtLote.Text) = "" Then
        Call MsgBox("Debe ingresar un código de lote.", vbExclamation Or vbSystemModal, Me.Caption)
        Exit Function
    End If
    If Trim(Me.txtNroPag.Text) = "" Then
        Call MsgBox("Debe ingresar el número de páginas del lote.", vbExclamation Or vbSystemModal, Me.Caption)
        Exit Function
    End If
    If Trim(Me.cmbMes.Text) = "" Then
        Call MsgBox("Debe seleccionar una mes válido.", vbExclamation Or vbSystemModal, Me.Caption)
        Exit Function
    End If
    If Trim(Me.mskfechaAnio.Text) = "____" Then
        Call MsgBox("Debe ingresar una año válido.", vbExclamation Or vbSystemModal, Me.Caption)
        Exit Function
    End If
    If IsNumeric(Me.mskfechaAnio.Text) = False Then
        Call MsgBox("Debe ingresar una año válido.", vbExclamation Or vbSystemModal, Me.Caption)
        Exit Function
    End If
    ValidarDatosObligatorios = True
End Function

Function ValidarReglas() As Boolean
    ValidarReglas = False
    CargaDatosAlObjetosDeDatos
    If mi_Opcion = sghAgregar Then
        If mo_ReglasHIS.ValidarLoteHIS_LoteExiste(oTablaDOHIS_Lote) Then
            Call MsgBox("El código del lote ya existe, elija otro código.", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    If Val(Me.txtNroPag.Text) < 1 Or Val(Me.txtNroPag.Text) > Val(lcBuscaParametro.SeleccionaFilaParametro(271)) Then
        Call MsgBox("El número total de hojas debe estar entre 1 y " & lcBuscaParametro.SeleccionaFilaParametro(271) & ".", vbInformation, Me.Caption)
        Exit Function
    End If
    If mi_Opcion = sghModificar Then
        If Val(Me.txtNroPag.Text) < ml_NroTotalHojas - 1 Then
            Call MsgBox("No se puede modificar el lote, revise el número total de hojas ", vbInformation, Me.Caption)
            Exit Function
        End If
        If oTablaDOHIS_Lote.IdEstadoLote = 1 And Me.chkPasarConcluido.Value Then
            If Val(Me.txtTotalRegistros.Text) = 0 Then
                Call MsgBox("No se puede modificar el lote, las hojas del lote no deben estar vacías", vbInformation, Me.Caption)
                Exit Function
            End If
        End If
    End If
    If mi_Opcion = sghEliminar Then
        If Val(Me.txtHojasDigitadas.Text) > 0 Then
            Call MsgBox("No se puede eliminar, porque ya tiene hojas registradas ", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    ValidarReglas = True
End Function

Function AgregarDatos() As Boolean
    AgregarDatos = mo_ReglasHIS.IngresarRegistroLoteHIS(oTablaDOHIS_Lote)
    LimpiarVariablesDeMemoria
End Function

Function ModificarDatos() As Boolean
    Dim mbActualizoRegistro As Boolean
    Dim oRcsTemp As New ADODB.Recordset
    Dim lnNroRegistro As Long
    
    If oTablaDOHIS_Lote.IdEstadoLote = 1 And Me.chkPasarConcluido.Value Then
        Select Case Val(lcBuscaParametro.SeleccionaFilaParametro(331))
        Case 0
            oTablaDOHIS_Lote.IdEstadoLote = 1
        Case 1, 2
            Call MsgBox("El módulo esta configurado para que los Lotes HIS tengan un grado de calidad. " & vbCrLf & "El lote pasará por verificación, ingrese por favor a la opción 'Calidad de Lotes'", vbInformation, "HIS - Calidad de Lotes")
            oTablaDOHIS_Lote.IdEstadoLote = 3
            'Poner númeracion a los registros
            Set oRcsTemp = mo_ReglasHIS.His_ConsultarTotalRegistrosLote(mo_cmbEstablecimiento.BoundText, ml_IdLote)
            oRcsTemp.MoveFirst
            lnNroRegistro = 0
            Do While Not oRcsTemp.EOF
                lnNroRegistro = lnNroRegistro + 1
                mbActualizoRegistro = mo_ReglasHIS.HisActualizarNroRegistroHisDetalle(CLng(oRcsTemp!IdHisDetalle), lnNroRegistro)
                oRcsTemp.MoveNext
            Loop
        End Select
    End If
    If oTablaDOHIS_Lote.IdEstadoLote = 4 And Me.chkPasarConcluido.Value Then
        Dim rcsTemp As ADODB.Recordset
        Dim lbTodoCoincide As Boolean
        Set rcsTemp = mo_ReglasHIS.HIS_ConsultarRegMuestraLotes(oTablaDOHIS_Lote.IdHisLote)
        If rcsTemp.RecordCount > 0 Then
            lbTodoCoincide = True
            rcsTemp.MoveFirst
            Do While Not rcsTemp.EOF
                If rcsTemp.Fields!Coincide <> 1 Then
                    lbTodoCoincide = False
                    Exit Do
                End If
                rcsTemp.MoveNext
            Loop
            If lbTodoCoincide = True Then
                oTablaDOHIS_Lote.IdEstadoLote = 4
            Else
                Call MsgBox("Verificación: " & vbCrLf & " Los registros del lote tienen inconsistencia con los registros de la doble digitación", vbInformation, "HIS - Calidad de Lotes")
                oTablaDOHIS_Lote.IdEstadoLote = 5
            End If
            If mo_ReglasHIS.EliminaHISDobleDigitacion(oTablaDOHIS_Lote.IdHisLote) = False Then
                Call MsgBox("No se pudo eliminar los datos de la doble digitación", vbInformation, "HIS - Calidad de Lotes")
            End If
        End If
    End If
    ModificarDatos = mo_ReglasHIS.ModificarRegistroLoteHIS(oTablaDOHIS_Lote)
    LimpiarVariablesDeMemoria
End Function

Function EliminarDatos() As Boolean
    EliminarDatos = mo_ReglasHIS.EliminarRegistroLoteHIS(oTablaDOHIS_Lote)
    LimpiarVariablesDeMemoria
End Function

Sub CargaDatosAlObjetosDeDatos()
    oTablaDOHIS_Lote.IdHisLote = ml_IdLote
    If mi_Opcion = sghAgregar Then
        oTablaDOHIS_Lote.IdEstadoLote = 0
    Else
        If oTablaDOHIS_Lote.IdEstadoLote = 0 Then
            If Val(Me.txtNroPag.Text) = Val(Me.txtHojasDigitadas.Text) Then
                If Me.chkPasarConcluido.Value Then
                    oTablaDOHIS_Lote.IdEstadoLote = 1
                End If
            End If
        End If
            Select Case oTablaDOHIS_Lote.IdEstadoLote
                Case 0 'Digitandose
                    If Val(Me.txtNroPag.Text) = Val(Me.txtHojasDigitadas.Text) Then
                        If Me.chkPasarConcluido.Value Then
                            oTablaDOHIS_Lote.IdEstadoLote = 1
                        End If
                    End If
                 Case 3 'Por Verificar'
                    If Me.chkPasarConcluido.Value Then
                        oTablaDOHIS_Lote.IdEstadoLote = 4
                    End If
                 Case 5 'Si Falla Verificado'
                    If Me.chkPasarConcluido.Value Then
                        oTablaDOHIS_Lote.IdEstadoLote = 0
                    End If
            End Select
    End If
    oTablaDOHIS_Lote.IdEstablecimiento = mo_cmbEstablecimiento.BoundText
    oTablaDOHIS_Lote.IdUsuarioAuditoria = ml_IdUsuario
    oTablaDOHIS_Lote.Lote = Me.txtLote.Text
    oTablaDOHIS_Lote.Mes = Val(Me.cmbMes.ItemData(Me.cmbMes.ListIndex))
    oTablaDOHIS_Lote.NroHojas = Me.txtNroPag.Text
    oTablaDOHIS_Lote.Anio = Me.mskfechaAnio.Text
    oTablaDOHIS_Lote.DobleDigitacion = 0
    oTablaDOHIS_Lote.Cerrado = 0    'solo se modifica , no se ha cerrado
End Sub

Sub LimpiarVariablesDeMemoria()
    Set mo_ReglasHIS = Nothing
    Set mo_cmbMes = Nothing
End Sub

Private Sub txtNroPag_KeyPress(KeyAscii As Integer)
    If ((KeyAscii < 48) Or KeyAscii > 57) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            If KeyAscii = 46 Then
                KeyAscii = 46
            Else
                KeyAscii = 1
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           If btnAceptar.Enabled Then btnAceptar_Click
       End Select
End Sub
