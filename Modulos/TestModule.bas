Attribute VB_Name = "TestModule"
Dim sTokenAutenticacion As String, sRespuesta As String, sParametros As String, sUrl As String
Dim respuestaJson As String, sToken As String
Dim objRequest As New WinHttpRequest
Dim r As WinHttpRequest



Public Sub AddParamJSON(sBuff As String, sNombreCampo As String, sValorCampo As String, Optional Cerrar As Boolean)
If Len(sBuff) = 0 Then sBuff = "{"
sBuff = sBuff & Chr$(34) & sNombreCampo & Chr$(34) & ":" & Chr$(34) & sValorCampo & Chr$(34)
If Cerrar Then sBuff = sBuff & "}" Else sBuff = sBuff & ","
End Sub


Public Sub callAmipassPay2(sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String)

    sTokenAutenticacion = "1348901"
    sRespuesta = ""
    
    Set httpURL = New WinHttp.WinHttpRequest

    'sParametros = Format("?NumeroTransaccion={0}&Monto=CodLocal={2}&CodPromocion={3}", sCodigoQR, sMonto, sCodLocal, sPromo)
    'sUrl = Convert.toString("https://pay.amipass.com/wspayTest/PayPAP") & sParametros
    'string sParametros = string.Format("?NumeroTransaccion={0}&Monto={1}&CodLocal={2}&CodPromocion={3}", sCodigoQR, sMonto, sCodLocal, sPromo);
    'string url = "https://pay.amipass.com/wspayTest/PayPAP" + sParametros;
    
    sParametros = "?NumeroTransaccion={" & sCodigoQR & "}&Monto={" & sMonto & "}&CodLocal={" & sCodLocal & "}&CodPromocion={" & sPromo & "}"
    sUrl = "https://pay.amipass.com/wspayTest/PayPAP" & sParametros
    
    httpURL.open "GET", sUrl
    httpURL.send
    sRespuesta = httpURL.responseText
    If sRespuesta = "[]" Then
       MsgBox ("No se obtuvo resultados")
       Exit Sub
    End If

    respuestaJson = "{items:" & sRespuesta & "}"
    'Set respuestaJson = JSON.parse(sInputJson)
    
    
MsgBox ("Paso")
    
    'request = DirectCast(HttpWebRequest.Create(New Uri(url)),HttpWebRequest)
    'request.contentType = "application/json"
    'request.Headers("Authorization") = String.Format("Basic{0}", TokenAutenticacion)
    
    'Using stream As New StreamReader(response.GetResponseStream())
   '     Dim content = stream.ReadToEnd()
   '     respuesta = JsonConvert.DeserializeObject(Of String)(content)
    'End Using
    'Dim resp As RespuestaTransaccionPa = JsonConvert.DeserializeObject(OfRespuestaTransaccionPA)(respuesta)
   ' If resp.CodRespuesta = "1" Them
        'Venta Aprobada'
    'Else
        'Venta Rechazada'
    'End If

    
End Sub

Public Sub callAmipassPayTEST(sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String)

    sToken = "1348901"
    sParametros = "?NumeroTransaccion={" & sCodigoQR & "}&Monto={" & sMonto & "}&CodLocal={" & sCodLocal & "}&CodPromocion={" & sPromo & "}"
    sUrl = "https://pay.amipass.com/wspayTest/PayPAP" & sParametros

    ' Build and send request
    With objRequest

        sTokenAutenticacion = "Basic {" & sToken & "}"
        .open "GET", sUrl, False
        .setRequestHeader "Authorization", sTokenAutenticacion
        .send
         sRespuesta = .responseText
    End With
    
    r = objRequest
    
    If sRespuesta = "[]" Then
        'No se obtuvo respuesta'
        MsgBox ("No se obtuvo resultados")
    End If
    
    MsgBox ("Finish!!")
    respuestaJson = JSON.parse("{items:" & sRespuesta & "}")


End Sub


Private Sub TEEEEST()

    Dim p             As Object
    Dim Texto         As String
    Dim sInputJson    As String
    Dim cab           As Integer

    Set httpURL = New WinHttp.WinHttpRequest

    usua = Trim(txtUsuario)
    Pass = Trim(txtPassword)
    
    'Aloja tu archivo php en tu hosting y cambia esta direccion
    cadena = "http://tupagina.com/prueba/login.php?USUARIO=" & usua & "&PASSWORD=" & Pass
    
    
    httpURL.open "GET", cadena
    httpURL.send
    Texto = httpURL.responseText
    If Texto = "[]" Then
       MsgBox ("No se obtuvo resultados")
       Exit Sub
    End If

    sInputJson = "{items:" & Texto & "}"

    Set p = JSON.parse(sInputJson)

    NOMBRE = p.Item("items").Item(1).Item("NOMBRE")

    MsgBox ("Bienvenido " & NOMBRE)


End Sub

Private Function formaterToStringJson(ByRef str As String, ByRef index As Long)

        'Quita \ y cambia de comillas dobles a simples
        'sRespuesta = Replace(sRespuesta, Chr(34), Chr(39))
        sRespuesta = Replace(sRespuesta, Chr(92), Chr(32))
        
        
        'sRespuesta = Replace(sRespuesta, Chr(123), Chr(91))
        'sRespuesta = Replace(sRespuesta, Chr(125), Chr(93))
        
        Do While index > 0 And index <= Len(str)
          'Select Case Mid(str, index, 1)
          MsgBox str
          
          Select Case Mid(str, index, 1)
          Case "CodRespuesta"
            MsgBox "ESTAS EN EL CASE 1"
          Case "DesRespuesta"
            MsgBox "ESTAS EN EL CASE 2"
          Case Else
            MsgBox "No se encontro"
          End Select
          
          index = index + 1
       Loop
    
            

End Function

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

