VERSION 5.00
Begin VB.Form formTest 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2085
      TabIndex        =   10
      Top             =   2805
      Width           =   1830
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2070
      TabIndex        =   9
      Top             =   2100
      Width           =   1800
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2055
      TabIndex        =   8
      Top             =   1500
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2055
      TabIndex        =   7
      Top             =   870
      Width           =   1785
   End
   Begin VB.CommandButton cmdRespuesta 
      Caption         =   "Respuesta Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1980
      TabIndex        =   2
      Top             =   4440
      Width           =   1785
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2025
      TabIndex        =   1
      Top             =   3660
      Width           =   1665
   End
   Begin VB.TextBox txtSalida 
      Height          =   5595
      Left            =   4605
      TabIndex        =   0
      Top             =   45
      Width           =   9015
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
Dim a As String
Dim sInputJson As String
Dim jRespueta As Object
Dim sSalida As String

'Dim sbJson As New ChilkatStringBuilder
'Dim success As Long
'Dim JSON As New ChilkatJsonObject

Private Sub Form_Load()
   
    sCodigoQR = ""
    sMonto = "0"
    sCodLocal = "76449"
    sPromo = "0"
    
    Text1 = sCodigoQR
    Text2 = sMonto
    Text3 = sCodLocal
    Text4 = sPromo
    
End Sub

Private Sub cmdRespuesta_Click()
    
    sInputJson = "{CodRespuesta:'1',DesRespuesta: 'APROBADO',CodAutorizacionz: '5270496',Fecha: '2016-09-06 17:05:04.210',Monto: '1000',TokenAN: '465464'}"
    sInputJson = "{CodRespuesta:'1',DesRespuesta: 'APROBADO'"
    'sInputJson = "{"CodRespuesta ": "53 ", "DesRespuesta ": "Codigo Tx Invalido ", "CodAutorizacion ": "15752 ", "Fecha ": "2021-03-26 15:45:48.913 ", "Monto ": "0 ", "Saldo ": "0 "}"

    
    'sInputJson = "{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }"
    a = Replace(sInputJson, Chr(34), Chr(39))
   
    MsgBox a
    'Convertimos cadena a Json
    Set jRespueta = JSON.parse(sInputJson)
    
    'Mostramos json en String
    txtSalida = JSON.toString(jRespueta)
    'MsgBox JSON.toString(jRespueta)
    
    'MsgBox "Respuesta: " & jRespueta.Item("DesRespuesta")
    
    'Accedemos al contenido
    'jRespueta .Item("items").Item(1).Item ("url")

    'Podemos agregar al Json
    'jRespueta.Item("items").Item(1).Add "ExtraItem", "Extra Data Value"
    
    MsgBox "Contenido Json: " & JSON.toString(jRespueta)
    
End Sub

Private Sub cmdTest_Click()
    
    sCodigoQR = Text1.Text
    sMonto = Text2.Text
    sCodLocal = Text3.Text
    sPromo = Text4.Text
    
    sSalida = callAmipassPay(sCodigoQR, sMonto, sCodLocal, sPromo)
    
    'a = Replace(sSalida, Chr(34), Chr(39))
    a = Replace(sSalida, Chr(92), Chr(32))
    'String to Json
    Set jRespueta = JSON.parse(sSalida)
    
    txtSalida = JSON.toString(jRespueta)
    MsgBox sSalida
    
End Sub

