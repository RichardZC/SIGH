VERSION 5.00
Begin VB.Form SeleccionaAnexoFua2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORMATO FUA VERSION 2"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SeleccionaAnexoFua2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   6855
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "SeleccionaAnexoFua2.frx":0442
         DownPicture     =   "SeleccionaAnexoFua2.frx":08A2
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
         Left            =   1920
         Picture         =   "SeleccionaAnexoFua2.frx":0D17
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "SeleccionaAnexoFua2.frx":118C
         DownPicture     =   "SeleccionaAnexoFua2.frx":1650
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
         Left            =   3480
         Picture         =   "SeleccionaAnexoFua2.frx":1B3C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "FUA VERSION 2: CONFIGURACION PARA EL DISEÑO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame 
         Caption         =   "SELECCIONE EL DISEÑO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   6615
         Begin VB.CommandButton btnVerAnexo1 
            Caption         =   "Vista Previa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   1320
            Picture         =   "SeleccionaAnexoFua2.frx":2028
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Visualizar el tipo de formato FUA configurado"
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton btnVerAnexo2 
            Caption         =   "Vista Previa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4440
            Picture         =   "SeleccionaAnexoFua2.frx":246A
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Visualizar el tipo de formato FUA configurado"
            Top             =   240
            Width           =   1275
         End
         Begin VB.OptionButton opFuaAnexo2 
            Caption         =   "ANEXO 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton opFuaAnexo1 
            Caption         =   "ANEXO 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label Label4 
         Caption         =   "* Vea la configuración en Herramientas -> Exporta/Importa datos SIS, en la     pestaña Datos Generales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   6615
      End
      Begin VB.Label Label3 
         Caption         =   "* El EESS esta configurado para usar los 2 diseños del nuevo formato FUA:      Anexo 1 y Anexo 2 de la RJ 107-2015 (Parametro 359)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label Label2 
         Caption         =   "* El EESS esta configurado con el nuevo formato FUA (Parametro 358)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   6615
      End
   End
End
Attribute VB_Name = "SeleccionaAnexoFua2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - Oficina de Informatica y Telecomunicaciones
'        Aplicativo: SisGalenPlus v.3
'        Programa: Selecciona el anexo para el Fua version 2
'        Programado por: Cachay F
'        Fecha: Setiembre 2015
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim mi_respuesta As Integer
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mo_doServicio As New doServicio
Dim ml_IdServicio As Long

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get Respuesta() As Integer
    Respuesta = mi_respuesta
End Property
Property Let Respuesta(lValue As Integer)
   mi_respuesta = lValue
End Property
Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property

Private Sub btnAceptar_Click()
    If Not (opFuaAnexo1.Value = True Or opFuaAnexo2.Value = True) Then
        MsgBox "Tiene que elegir el diseño", vbInformation, Me.Caption
        Exit Sub
    End If
    Set mo_doServicio = RetornaServicio(ml_IdServicio)
    If opFuaAnexo1.Value = True Then
        mi_respuesta = 1
    ElseIf opFuaAnexo2.Value = True Then
        mi_respuesta = 2
    End If
    If MsgBox("Desea Guardar la configuración del formato FUA para el servicio " + mo_doServicio.nombre, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
        mo_doServicio.FuaTipoAnexo2015 = CLng(mi_respuesta)
        If ModificarServicio(mo_doServicio) = True Then
            MsgBox "Se guardo la configuración del formato FUA para el servicio " + mo_doServicio.nombre + "." + Chr(13) + "Si posteriormente desea modificar la configuración del Servicio acceda al módulo General -> Servicios", vbInformation, Me.Caption
        End If
    End If
    Me.Visible = False
    Set mo_doServicio = Nothing
End Sub

Private Sub btnVerAnexo1_Click()
    VerDisenoFuaVersion2 ("1")
End Sub

Private Sub btnCancelar_Click()
    mi_respuesta = 0
    Me.Visible = False
End Sub

Private Sub btnVerAnexo2_Click()
    VerDisenoFuaVersion2 ("2")
End Sub

Sub VerDisenoFuaVersion2(mc_TipoAnexo As String)
    Dim Ruta As String
    Dim lcFormatoFua As String
    Dim lcTipoAnexo As String
    lcFormatoFua = lcBuscaParametro.SeleccionaFilaParametro(358)
    Ruta = App.Path + "\Imagenes\FUA\" + lcFormatoFua + mc_TipoAnexo + "\" + lcFormatoFua + mc_TipoAnexo + "-1.png"
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", Ruta, "", "", 1)
End Sub

Function ModificarServicio(mo_Servicios As doServicio) As Boolean
    ModificarServicio = mo_AdminServiciosHosp.ServiciosModificar(mo_Servicios, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, mo_Servicios.nombre)
End Function

Function RetornaServicio(ml_IdServ As Long) As doServicio
    Dim oConexion As New Connection
    Dim oDoServicio As New doServicio
    Dim oServicio As New Servicios
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    Set oServicio.Conexion = oConexion
    oDoServicio.IdServicio = ml_IdServ
    If oServicio.SeleccionarPorId(oDoServicio) Then
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oServicio = Nothing
    Set RetornaServicio = oDoServicio
End Function

