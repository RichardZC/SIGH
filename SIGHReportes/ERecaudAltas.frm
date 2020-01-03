VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ERecaudAltas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recaudación de Altas Hospitalización"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   75
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ERecaudAltas.frx":0000
         DownPicture     =   "ERecaudAltas.frx":0460
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
         Left            =   1410
         Picture         =   "ERecaudAltas.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ERecaudAltas.frx":0D4A
         DownPicture     =   "ERecaudAltas.frx":120E
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
         Left            =   2940
         Picture         =   "ERecaudAltas.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1020
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   5565
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   1980
         TabIndex        =   2
         Top             =   570
         Width           =   1395
         _ExtentX        =   2461
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
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Alta Administ. Final"
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
         Left            =   120
         TabIndex        =   4
         Top             =   615
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Alta Administ. Inic"
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
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   1620
      End
   End
End
Attribute VB_Name = "ERecaudAltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Recaudación de Altas
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Private Sub btnAceptar_Click()
    If Me.txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de alta administrativa inicial", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha de alta administrativa inicial no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    
    If Me.txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese la fecha de alta administrativa final", vbInformation, Me.Caption
        Exit Sub
    Else
        If Not SIGHEntidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha de alta administrativa final no tiene el formato correcto", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Sub
    End If
    
    Me.MousePointer = 11
        Dim oRptRecaudAltas As New RptERecaudAltas
        oRptRecaudAltas.FechaInicio = txtFechaInicio.Text
        oRptRecaudAltas.FechaFin = txtFechaFin.Text
        oRptRecaudAltas.TextoDelFiltro = "Filtros:     F.Altas Administrativas: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
        oRptRecaudAltas.CrearReporte_excel Me.hwnd
    Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Private Sub Form_Load()
    Me.txtFechaInicio.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    Me.txtFechaFin.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
