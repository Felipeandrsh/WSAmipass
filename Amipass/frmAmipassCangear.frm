VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAmipassCangear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cangear Amipesos"
   ClientHeight    =   7905
   ClientLeft      =   7890
   ClientTop       =   3960
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11265
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3075
      Left            =   495
      TabIndex        =   9
      Top             =   4575
      Width           =   10545
      Begin MSComctlLib.ListView LvAmiCange 
         Height          =   2535
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cobrado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Mensaje"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Respuesta Json"
      Height          =   3500
      Left            =   5500
      TabIndex        =   6
      Top             =   1000
      Width           =   5500
      Begin VB.TextBox txtSalida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   330
         Width           =   5190
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cangear Amipesos"
      Height          =   3500
      Left            =   500
      TabIndex        =   0
      Top             =   1000
      Width           =   5000
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3400
         TabIndex        =   8
         Top             =   2800
         Width           =   1300
      End
      Begin VB.TextBox txtRutCustomer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         TabIndex        =   3
         Top             =   400
         Width           =   2700
      End
      Begin VB.TextBox txtCodeLocal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   950
         Width           =   2700
      End
      Begin VB.CommandButton cmdCangearAmipesos 
         Caption         =   "Cangear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         TabIndex        =   1
         Top             =   2800
         Width           =   1300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   5
         Top             =   530
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmAmipassCangear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables Venta
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String, sSalida As String, sInputJson As String

'Variables reporte
Dim sCodLocalReporte As String, sFecha As String, sRutCliente As String

Dim Jrespuesta As Object, itmx As ListItem



Dim sRespuesta As String, sNumTransaccion As String

Private Sub cmdCangearAmipesos_Click()
    
    Dim SpMensajeError As String
    sRutCliente = txtRutCustomer.Text
    sCodLocal = txtCodeLocal.Text
    sRespuesta = getCustomerCange(sRutCliente, sCodLocal)
    
    
    'Transforma a Json
    Set Jrespuesta = JSON.parse(sRespuesta)
    
    'Verifica respuesta
    If Jrespuesta.Item("status") <> 200 Then
        SpMensajeError = "Estado: " & Jrespuesta.Item("status") & " " & vbCrLf & _
                         "Respuesta: No se pudo completar transaccion"
        MsgBox SpMensajeError
        Exit Sub
    End If
    
    frmAmipassCangear.LvAmiCange.ListItems.Clear
    
    Dim Sp_cobrado As String, Bp_cobrado As Boolean
    Sp_cobrado = "NO"
    Bp_cobrado = Replace(Jrespuesta.Item("response").Item("bCobrado"), " ", "") 'Transformo a Boolean
    If Bp_cobrado Then
        Sp_cobrado = "SI"
    End If
    
    'LLENA GRILLA
    Set itmx = frmAmipassCangear.LvAmiCange.ListItems.Add(, , Sp_cobrado)
    itmx.SubItems(1) = Jrespuesta.Item("response").Item("sMensaje")
    
    txtSalida = sRespuesta

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    sCodLocal = "76449"
    
    txtRutCustomer = "NN33028"
    txtCodeLocal = sCodLocal
End Sub

