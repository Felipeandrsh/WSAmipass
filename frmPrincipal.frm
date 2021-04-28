VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Principal"
   ClientHeight    =   6135
   ClientLeft      =   6165
   ClientTop       =   1035
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   675
      Left            =   1215
      TabIndex        =   5
      Top             =   5130
      Width           =   1710
   End
   Begin VB.CommandButton cmdCanjear 
      Caption         =   "Canjear"
      Height          =   675
      Left            =   1215
      TabIndex        =   4
      Top             =   4365
      Width           =   1710
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      Height          =   675
      Left            =   1215
      TabIndex        =   3
      Top             =   3615
      Width           =   1710
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular"
      Height          =   675
      Left            =   1215
      TabIndex        =   2
      Top             =   2835
      Width           =   1710
   End
   Begin VB.CommandButton cmdVerificar 
      Caption         =   "Verificar"
      Height          =   675
      Left            =   1215
      TabIndex        =   1
      Top             =   2070
      Width           =   1710
   End
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar"
      Height          =   675
      Left            =   1215
      TabIndex        =   0
      Top             =   1305
      Width           =   1710
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnular_Click()
frmAmipassAnulacion.Show vbModal
End Sub

Private Sub cmdCanjear_Click()
frmAmipassCangear.Show vbModal
End Sub

Private Sub cmdPagar_Click()
frmAmipassPago.Show vbModal
End Sub

Private Sub cmdReporte_Click()
frmAmipassReporte.Show vbModal
End Sub

Private Sub cmdTest_Click()
formTest.Show vbModal
End Sub

Private Sub cmdVerificar_Click()
frmAmipassVerificacion.Show vbModal
End Sub
