VERSION 5.00
Begin VB.Form formTest 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2085
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   2805
      Width           =   1830
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2070
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   2100
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2055
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1500
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2055
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   870
      Width           =   1785
   End
   Begin VB.CommandButton cmdRespuesta 
      Caption         =   "Respuesta Test"
      Height          =   510
      Left            =   2040
      TabIndex        =   2
      Top             =   4485
      Width           =   1785
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Conexion Test"
      Height          =   510
      Left            =   2025
      TabIndex        =   1
      Top             =   3660
      Width           =   1665
   End
   Begin VB.TextBox txtSalida 
      Height          =   5595
      Left            =   4530
      TabIndex        =   0
      Top             =   135
      Width           =   6210
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   3
      Left            =   900
      TabIndex        =   6
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Promo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1125
      TabIndex        =   5
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo QR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   765
      TabIndex        =   3
      Top             =   930
      Width           =   1215
   End
End
Attribute VB_Name = "formTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String
'Dim a As String
Dim sInputJson As String
Dim jRespueta As Object
Dim sSalida As String

'Dim sbJson As New ChilkatStringBuilder
'Dim success As Long
'Dim JSON As New ChilkatJsonObject

Private Sub Form_Load()
   
    sCodigoQR = "78411803"
    sMonto = "1"
    sCodLocal = "76449"
    sPromo = "0"
    
    Text1 = sCodigoQR
    Text2 = sMonto
    Text3 = sCodLocal
    Text4 = sPromo
    
End Sub

Private Sub cmdRespuesta_Click()
    
    sInputJson = "{CodRespuesta:'1',DesRespuesta: 'APROBADO',CodAutorizacionz: '5270496',Fecha: '2016-09-06 17:05:04.210',Monto: '1000',TokenAN: '465464'}"
    'sInputJson = "{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }"
   
    'Convertimos cadena a Json
    Set jRespueta = json.parse(sInputJson)
    
    'Mostramos json en String
    txtSalida = json.toString(jRespueta)
    'MsgBox JSON.toString(jRespueta)
    
    'MsgBox "Respuesta: " & jRespueta.Item("DesRespuesta")
    
    'Accedemos al contenido
    'jRespueta .Item("items").Item(1).Item ("url")

    'Podemos agregar al Json
    'jRespueta.Item("items").Item(1).Add "ExtraItem", "Extra Data Value"
    
    MsgBox "Contenido Json: " & json.toString(jRespueta)
    
End Sub

Private Sub cmdTest_Click()
    
    sCodigoQR = Text1.Text
    sMonto = Text2.Text
    sCodLocal = Text3.Text
    sPromo = Text4.Text
    
    sSalida = callAmipassPay(sCodigoQR, sMonto, sCodLocal, sPromo)
    txtSalida = json.toString(sSalida)
    MsgBox sSalida
    
End Sub

