VERSION 5.00
Begin VB.Form FLETE00_VB6 
   BackColor       =   &H80000016&
   Caption         =   "SISTEMA CONTROL DE PAGO FLETES A TRANSPORTISTAS"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Salir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4560
      TabIndex        =   2
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "  FLETES DE TRANSPORTISTAS  "
      Height          =   5295
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   8055
      Begin VB.Image Image2 
         Height          =   2115
         Left            =   2640
         Picture         =   "FLETE00_VB.frx":0000
         Top             =   1680
         Width           =   2265
      End
   End
   Begin VB.Label Label1 
      Caption         =   "CONTROL DE PAGO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
   Begin VB.Menu Actualizar 
      Caption         =   "Actualizar"
      Begin VB.Menu Paises 
         Caption         =   "Paises"
      End
      Begin VB.Menu Estados 
         Caption         =   "Estados"
      End
      Begin VB.Menu Destino 
         Caption         =   "Destino"
      End
      Begin VB.Menu Procesar_Relacion_Fletes 
         Caption         =   "Procesar Relacion Fletes"
      End
      Begin VB.Menu Actualizar_Relacion_Pago 
         Caption         =   "Actualizar Relacion Pago Fletes"
      End
      Begin VB.Menu Cierre_Fletes_Cancelados 
         Caption         =   "Ejecutar Cierre de Fletes Cancelados"
      End
   End
   Begin VB.Menu Reportes 
      Caption         =   "Reportes"
      Begin VB.Menu Relacion_Pago_Fletes 
         Caption         =   "Relacion de Pago Flegtes"
      End
   End
   Begin VB.Menu Probando 
      Caption         =   "Probando"
   End
   Begin VB.Menu Salir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "FLETE00_VB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  Sistema de Control de Pago Fletes Transportistas.
'  Autor: Henry J. Pulgar B.
'  Creado : 19-10-2006.
'  Actualizado : 15-12-2006.
'  Manejador de Bases de Datos ORACLE Rdbms v 8.0.6.
'  Aplicativo creado en Visual Basic ( Visual Studio v 6.i )
'  Acceso: Al Rdbms a travez de una ODBC.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
Dim CurrentDir As String
Dim CurrentUser As String
'

Private Sub Actualizar_Relacion_Pago_Click()
 Comando = "ifrun60 " & CurrentDir & "FLETE05_FRM " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub



'*****************************************************************************
Private Sub Form_Load()
  CurrentDir = ""
  CurrentUser = "OPS$DESFLE00/OPS$DESFLE00@BD806"
End Sub

'*****************************************************************************
Private Sub Destino_Click()
  Comando = "ifrun60 " & CurrentDir & "FLETE03v3_FRM " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

'*****************************************************************************
Private Sub Estados_Click()
  FLETE02_VB6.Show
End Sub

'******************************************************************************
Private Sub Cmd_Salir_Click()
    Unload Me 'FLETE00_VB6
End Sub

'******************************************************************************
Private Sub Paises_Click()
   FLETE01_VB6.Show
End Sub

'******************************************************************************
Private Sub Probando_Click()
 PRUEBA.Show
End Sub

Private Sub Procesar_Relacion_Fletes_Click()
  Comando = "ifrun60 " & CurrentDir & "FLETE04_FRM " & CurrentUser
  ExeComando = Shell(Comando, vbMaximizedFocus)
End Sub

Private Sub Relacion_Pago_Fletes_Click()
  Comando = "rwrun60 report=" & CurrentDir & "Control_Pago_Fletes.rdf userid=" & CurrentUser
  ExeComando = Shell(Comando, vbNormalFocus)
End Sub

'******************************************************************************
'Private Sub Reportes_Click()
'  PAISES_VB6.Show
'End Sub
'******************************************************************************
Private Sub Salir_Click()
  Unload Me
End Sub
