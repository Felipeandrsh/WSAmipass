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
    respuestaJson = json.parse("{items:" & sRespuesta & "}")


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

    Set p = json.parse(sInputJson)

    NOMBRE = p.Item("items").Item(1).Item("NOMBRE")

    MsgBox ("Bienvenido " & NOMBRE)


End Sub


