VERSION 5.00
Begin VB.Form TipoModalidadSala 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   11610
   ClientTop       =   6240
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7905
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   120
      TabIndex        =   7
      Top             =   1230
      Width           =   7680
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "TipoModalidadSala.frx":0000
         DownPicture     =   "TipoModalidadSala.frx":04C4
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
         Left            =   4080
         Picture         =   "TipoModalidadSala.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "TipoModalidadSala.frx":0E9C
         DownPicture     =   "TipoModalidadSala.frx":12FC
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
         Left            =   2565
         Picture         =   "TipoModalidadSala.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7680
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   5
         TabIndex        =   0
         Top             =   270
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   5505
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
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
         TabIndex        =   6
         Top             =   690
         Width           =   960
      End
      Begin VB.Label lblCodigoCIE2004 
         Caption         =   "Código"
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
         Left            =   135
         TabIndex        =   5
         Top             =   285
         Width           =   1335
      End
   End
End
Attribute VB_Name = "TipoModalidadSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de tipo de modalidad de sala
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_ImagTipoModalidadSala As New DOImagTipoModalidadSala
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdTipoModalidadSala As Long
Dim mo_ReglasImagenes As New ReglasImagenes
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_cmbIdTipoComprobante As New ListaDespleglable
Dim mb_RegistroInactivo As Boolean
Dim mb_ResultadoOperacion As Boolean

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property

Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let IdTipoModalidadSala(lValue As Long)
   ml_IdTipoModalidadSala = lValue
End Property
Property Get IdTipoModalidadSala() As Long
   IdTipoModalidadSala = ml_IdTipoModalidadSala
End Property

Property Let ResultadoOperacion(bValue As Boolean)
   mb_ResultadoOperacion = bValue
End Property

Property Get ResultadoOperacion() As Boolean
   ResultadoOperacion = mb_ResultadoOperacion
End Property

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
   AdministrarKeyPreview KeyCode
End Sub
Private Sub txtCodigo_LostFocus()
    txtCodigo = UCase(txtCodigo)
   mo_Formulario.MarcarComoVacio txtCodigo
End Sub
Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
   AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         Frame1.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       CargaComboBoxes
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Tipo Modalidad Sala"
       Case sghModificar
           Me.Caption = "Modificar Tipo Modalidad Sala"
       Case sghConsultar
           Me.Caption = "Consultar Tipo Modalidad Sala"
       Case sghEliminar
           Me.Caption = "Eliminar Tipo Modalidad Sala"
       End Select
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

Sub CargaComboBoxes()
   
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   mb_ResultadoOperacion = False
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
            mb_RegistroInactivo = False
           If ValidarReglas() Then
'                Dim bSuccess As Boolean
                If mb_RegistroInactivo = False Then
                    mb_ResultadoOperacion = AgregarDatos()
                Else
                    mb_ResultadoOperacion = ModificarDatos()
                End If
               If mb_ResultadoOperacion Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
                   Me.txtCodigo.SetFocus
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   mb_ResultadoOperacion = True
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_ReglasImagenes.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   If Me.txtCodigo.Text = "" Then
       sMensaje = sMensaje + "Ingrese el código " + Chr(13)
   End If
   If Me.txtDescripcion = "" Then
       sMensaje = sMensaje + "Ingrese la descripción" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   'si existe y esta inactivo dejar que siga su curso
   CargaDatosAlObjetosDeDatos
   Dim oRs As ADODB.Recordset
   
    If mi_Opcion = sghAgregar Or mi_Opcion = sghModificar Then
        If mi_Opcion = sghAgregar Then
            mo_ImagTipoModalidadSala.IdTipoModalidadSala = 0
        Else
            mo_ImagTipoModalidadSala.IdTipoModalidadSala = ml_IdTipoModalidadSala
        End If
        Set oRs = mo_ReglasImagenes.ImagTipoModalidadSalaVerificarPorNombre(mo_ImagTipoModalidadSala)
         If oRs Is Nothing Then
             MsgBox mo_ReglasImagenes.MensajeError, vbInformation, "Tipo Modalidad Sala"
             Exit Function
         Else
             Dim lNumeroError As Long
             lNumeroError = ValidarDuplicadoRegistro(mi_Opcion, oRs)
             If lNumeroError > 0 Then
                MsgBox mapearMensageErrorDupplicado(lNumeroError), vbInformation, "Tipo Modalidad Sala"
                Exit Function
             End If
         End If
    End If
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   With mo_ImagTipoModalidadSala
        .Codigo = Me.txtCodigo.Text
        .TipoModalidadSala = UCase(Me.txtDescripcion.Text)
        .IdUsuarioAuditoria = Me.idUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_ReglasImagenes.ImagTipoModalidadSalaAgregar(mo_ImagTipoModalidadSala, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_ReglasImagenes.ImagTipoModalidadSalaModificar(mo_ImagTipoModalidadSala, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_ReglasImagenes.ImagTipoModalidadSalaEliminar(mo_ImagTipoModalidadSala, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    
    Set mo_ImagTipoModalidadSala = mo_ReglasImagenes.ImagTipoModalidadSalaSeleccionarPorId(Me.IdTipoModalidadSala)
    If mo_ReglasImagenes.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_ReglasImagenes.MensajeError, vbInformation, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_ImagTipoModalidadSala Is Nothing Then
        With mo_ImagTipoModalidadSala
            Me.txtCodigo = .Codigo
            Me.txtDescripcion = .TipoModalidadSala

            mb_ExistenDatos = True
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdTipoModalidadSala = 0
    Me.txtCodigo = ""
    Me.txtDescripcion = ""
    Err = 0
End Sub


Private Function ValidarDuplicadoRegistro(iOpcion As sghOpciones, oRs As ADODB.Recordset) As Integer
    Dim lTotalRegistros As Long
    ValidarDuplicadoRegistro = 0
    
    If oRs.RecordCount > 0 Then
        lTotalRegistros = oRs.RecordCount
        
        oRs.MoveFirst
        While oRs.EOF = False
            If oRs.Fields!EsActivo = True Then
                'duplicado activo
                ValidarDuplicadoRegistro = 1
                Exit Function
            Else
                If iOpcion = sghModificar Then
                    'registro duplicado e inactivo si se permite la edicion ocasionaria dos registros uno activo y otro inactivo
                    ValidarDuplicadoRegistro = 2
                    Exit Function
                ElseIf iOpcion = sghAgregar Then
                    'si hay mas de un registro inactivo duplicado no se puede permitir activar uno al azar
                    If lTotalRegistros > 1 Then
                        ValidarDuplicadoRegistro = 3
                        Exit Function
                    Else
'                        ml_IdTipoModalidadSala = oRs.Fields!IdTipoModalidadSala
                        mo_ImagTipoModalidadSala.IdTipoModalidadSala = oRs.Fields!IdTipoModalidadSala
                        mb_RegistroInactivo = True
                        Exit Function
                    End If
                End If
            End If
            oRs.MoveNext
        Wend
    End If
End Function

Private Function mapearMensageErrorDupplicado(lNumeroError As Long) As String
    Dim sMessage As String
    Select Case lNumeroError
        Case 1:
            sMessage = "Existe un registro con la misma descripción o código"
        Case 2:
            sMessage = "Existe un registro con la misma descripción o código que ha sido eliminado"
        Case 3:
            sMessage = "Existe mas de un registro con la misma descripción o código que han sido eliminados"
    End Select
    
    mapearMensageErrorDupplicado = sMessage
End Function
