Public oWizard, oConn, ServerName, DatabaseName, UserName, Password
Public oFSH, rd, sEsMain, sEsMulti, rdArchExt
Public oFile, Fso2, Ferr, sNombreAuxArchivo, sArchivo, sCodigDeTransaccion, MailVendedor, MailCliente, sTexto, sParrafo, nArchivo

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Const ObjectType = 5                    'Constante que indica el tipo de objeto. Este vale 5 para los reportes
'************ Datos empresa **********
Const cNombre_Empresa = "MPE"

'************ para usar usuario encriptado **********
Dim sUser, sPassWord, sServer, sDriver, dElementHtml
Dim sDatabase 'Siempre asignar el nombre de la empresa antes de cualquier otra cosa
Dim sCodigoParametrizacion,sCodCom,sChequearContacto,dKeyValue
Dim sSql, sCodemp, sModfor, sCodfor, lNrofor, sNrocta, sNombre, sMail, Attach, sTextoTitulo, sTxtDetalle, sCodapl, lNroapl, lImport, sTxtOP
Dim sQuery, sPriPar, sSegPar, sTerPar, arrArchivos,contador
Dim sMailFrom, sSMTPServerCorreo, sPasswordCorreo, sUsuarioCOrreo, sPuertoCorreo, sTituloMail, sPdfPath
Dim sNombreEmpresa, sDirEmpresa, sCodPosEmpresa, sCodPaiEmpresa, sLocalidadEmpresa, sFirmante, sNoTieneContacto
'****************************************************************************
'*******************Codigo de Parametrizacion a utilizar*********************
'****************************CARGAR SIEMPRE!!!!******************************
'****************************************************************************

                sCodigoParametrizacion = "FCRMVH"

'****************************************************************************


'************ Variables a utilizar ****************************
'@@Cliente          : Recupera Cliente
'@@RazonSocial      : Recupera Razon Social Cliente
'@@Comprobante      : Formulario a enviar a pdf y mail
'@@Numero           : Numero del Formulario a enviar a pdf y mail
'@@NombreEmpresa    : Razon Social Empresa
'@@DireccionEmpresa : Direccion Empresa
'@@PaisEmpresa      : Pais Empresa
'@@CodPosEmpresa    : Codigo Postal Empresa
'@@LocalidadEmpresa : Localidad Empresa
'@@Firmante         : Firmante Mail
'Utilizable en Query SQL
'@@Codemp           : Codigo Empresa =cNombre_Empresa
'*************************************************************
' Main(){
sChequearContacto = ""
openConnection
CreateDictionaryElementHtml
sChequearContacto = chequearContactoSinCheck
unMailCheck = 1
'Cargo parametros Globales es multi es main
esMultiEmpresa

EjecutoEnvioMail sCodigoParametrizacion, sEsMulti, sEsMain

closeConnection

'}
'**********************************************************
'   Funciones
'**************************************************
Function EjecutoEnvioMail(CodigoDeTransaccion, EsMulti, EsMain)
    sNoTieneContacto = "N"
    'Recupero datos del abm
    sSql1 = "Select Convert(Varchar(Max), SAR_ENVMAH_RECUPE) QUERY , "
    sSql1 = sSql1 & " SAR_ENVMAH_MAILOR MAILFROM,"
    sSql1 = sSql1 & " SAR_ENVMAH_SNDUSR USUARIO,"
    sSql1 = sSql1 & " SAR_ENVMAH_PSWUSR PASSWORD,"
    sSql1 = sSql1 & " SAR_ENVMAH_SMTPSR SMTPSERVER,"
    sSql1 = sSql1 & " SAR_ENVMAH_PUERTO PUERTO ,"
    sSql1 = sSql1 & " SAR_ENVMAH_TITULO TITULO, "
    sSql1 = sSql1 & " SAR_ENVMAH_PDFPAT PDFPATH "
    sSql1 = sSql1 & " From SAR_ENVMAH WHERE SAR_ENVMAH_CODIGO='" & CodigoDeTransaccion & "'"
    Set rd1 = oConn.Execute(CStr(sSql1))

    If Not rd1.EOF Then
        sQuery = rd1("QUERY").value
        sMailFrom = rd1("MAILFROM").value
        sUsuarioCOrreo = rd1("USUARIO").value
        sPasswordCorreo = rd1("PASSWORD").value
        sSMTPServerCorreo = rd1("SMTPSERVER").value
        sPuertoCorreo = rd1("PUERTO").value
        sTituloMail = rd1("TITULO").value
        sPdfPath = rd1("PDFPATH").value
    End If
    rd1.Close
    Set rd1 = Nothing


    '---------------------------------------
    'Definir el firmante del mail-----------
    sFirmante = "" '----------
    '---------------------------------------
  sNombreEmpresa = normalValue("NOMEMP","NOMEMP")
  sDirEmpresa = normalValue("DIREMP","DIREMP")
  sCodPaiEmpresa = normalValue("CODPAI","PAIS_CEDE")
  sCodPosEmpresa = normalValue("CODPOS","CODPOS_SEDE")
  sLocalidadEmpresa = normalValue("LOCALI","LOCEMP")

    'Asigno Query recupera Comprobantes
    sQuery = replace(sQuery, "@@Codemp", cNombre_Empresa)
    sSql = sQuery

    Set rd = oConn.Execute(CStr(sSql))

    arrArchivos = Array()
    contador = 0
	 contador2 = 0
    Do While Not rd.EOF
        If EsMulti = "S" Then
            sCodemp = rd("CODEMP").value
        Else
            sCodemp = ""
        End If
        sModfor = rd("MODFOR").value
        sCodfor = rd("CODFOR").value
        sCodCom = rd("CODCOM").value
        lNrofor = rd("NROFOR").value
        sNrocta = rd("NROCTA").value
        sNombre = rd("NOMBRE").value
        sNroCtaConEspacios = rd("NROCTAM").value

        CreateDictionaryKeyValue
        sTexto = ""
        sTituloCache = sTituloMail
        sTituloCache = replace(sTituloCache, "@@Comprobante", sCodfor)
        sTituloCache = replace(sTituloCache, "@@Numero", lNrofor)

        'Busqueda de los 3 parrafos del reporte
        sSql1 = "SELECT "
        sSql1 = sSql1 & " replace(convert(varchar(Max),SAR_ENVMAH_PRIPAR),char(13),'<br>') PRIPAR,"
        sSql1 = sSql1 & " replace(convert(varchar(Max),SAR_ENVMAH_SEGPAR),char(13),'<br>') SEGPAR,"
        sSql1 = sSql1 & " replace(convert(varchar(Max),SAR_ENVMAH_TERPAR),char(13),'<br>') TERPAR"
        sSql1 = sSql1 & " FROM SAR_ENVMAH WHERE SAR_ENVMAH_CODIGO='" & CodigoDeTransaccion & "'"
        Set rd1 = oConn.Execute(CStr(sSql1))
        If Not rd1.EOF Then
            sPriPar = rd1("PRIPAR").value
            sSegPar = rd1("SEGPAR").value
            sTerPar = rd1("TERPAR").value
        End If
        rd1.Close
        Set rd1 = Nothing
        CreateBody sTexto, CodigoDeTransaccion
        iCount = TieneAnexo(sCodemp, sModfor, sCodfor, lNrofor)

        GeneroPDF sPdfPath, sNrocta, sCodemp, sModfor, sCodfor, lNrofor, sEsMain, EsMulti, sCodCom
        if (instr(sCodCom,"FCA")<>0 or instr(sCodCom,"FCP")<>0 or instr(sCodCom,"FCU")<>0 or instr(sCodCom,"FCX")>0)and iCount>0 then
        	 GeneroPDFCV sPdfPath, sNrocta, sCodemp, sModfor, sCodfor, lNrofor, sEsMain, EsMulti,sCodCom
        	End if
        GeneroPDFVT sPdfPath, sNrocta, sCodemp, sModfor, sCodfor, lNrofor, sEsMain, EsMulti

        sQueryArchExt = "SELECT FCRMVH_OLEOLE RUTA FROM FCRMVH WHERE  FCRMVH_MODFOR = '"& sModfor &"' AND "&_
            "FCRMVH_CODFOR = '"& sCodfor &"' AND FCRMVH_NROFOR = "&lNrofor&" AND "&_
            "FCRMVH_CODEMP = '"&sCodemp&"' AND FCRMVH_OLEOLE IS NOT NULL "
        grabarLog_Archivo(Cstr(sQueryArchExt))

        Set rdArchExt = oConn.Execute(CStr(sQueryArchExt))
        grabarLog_Archivo(Cstr(contador))

        Do While Not rdArchExt.EOF
          sRuta = rdArchExt("RUTA").value
          if sRuta <> "NULL" or sRuta <> "" Then
            Redim Preserve arrArchivos(contador)
            arrArchivos(contador) = Cstr(sRuta)
            grabarLog_Archivo("ARCHIVO EXTERNO "+ Cstr(sRuta))
            contador = contador +1
          End if
          rdArchExt.MoveNext
        Loop




        EnviaMail sTexto, sNroCtaConEspacios , sTituloCache, sUsuarioCOrreo, sPasswordCorreo, sSMTPServerCorreo, sPuertoCorreo, nArchivo, sMailFrom
        contador2 = contador2 +1
        grabarLog_Archivo("mail "+ Cstr(contador2))


  	  If sNoTieneContacto = "N" Then
        MarcoComprobantes sCodemp, sModfor, sCodfor, lNrofor, sEsMulti
      End If

     	 arrArchivos = Array()
     	 contador = 0


        rd.MoveNext
    Loop
    rd.Close
    Set rd = Nothing
End Function
Function CreateBody(sBody,CodigoDeTransaccion)
  Dim sQueryTemplate,  Resultado
  Dim rd, ObjRegEx

  sQueryTemplate = "Select USR_TEMPLATE_HTML HTML from USR_TEMPLATE where USR_TEMPLATE_CODIGO = 'LOGO CENTRADO'"
  Set rd = oConn.Execute(CStr(sQueryTemplate))
  sBody = rd("HTML").Value

  'Diccionario de Claves @@ y Valor
  CreateDictionaryKeyValue

  'Recorro las claves del Diccionario y Remplazo en los parrafos
  ReplaceAllKey "Parrafos", dKeyValue
  sSegPar = Replace(sSegPar,"#*","&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<font face='Symbol' size='2'><span style='font-size:10pt;'>·&nbsp&nbsp</span></font>")


  sBody = Replace(sBody,"[#PARRAFO1#]",sPriPar)
  sBody = Replace(sBody,"[#PARRAFO2#]",sSegPar)
  sBody = Replace(sBody,"[#PARRAFO3#]",sTerPar)

  IsExists sBody,"[#BOTON#]","","[#BotonEditable#]"
  IsExists sBody,"[#LINK#]","","[#BotonEditable#]"
  IsExists sBody,"[#DIRECCION#]",Trim(sDirEmpresa) ,"[#DireccionEditable#]"
  IsExists sBody,"[#EMAIL#]",Trim(sEmail),"[#EmailPhoneEditable#]"
  IsExists sBody,"[#TELEFONO#]",Trim(sTelefono),"[#EmailPhoneEditable#]"
  IsExists sBody,"[#LOGO#]","http://i65.tinypic.com/zodfl4.png","[#LogoEditable#]"
  IsExists sBody,"[#TITULO#]","Envio de Facturas","[#TituloEditable#]"
  IsExists sBody,"[#INSTAGRAM#]","","[#InstagramEditable#]"
  IsExists sBody,"[#FACEBOOK#]","https://www.facebook.com/IRTMedicinaLaboral","[#FacebookEditable#]"
  IsExists sBody,"[#TWITTER#]","","[#TwitterEditable#]"
  IsExists sBody,"[#LIKEDIN#]","" ,"[#LinkedinEditable#]"

End Function
Function CreateDictionaryKeyValue()
  Set dKeyValue = CreateObject("Scripting.Dictionary")
  dKeyValue.Add "@@Cliente" , sNrocta
  dKeyValue.Add "@@RazonSocial" , sNombre
  dKeyValue.Add "@@Comprobante", sCodfor
  dKeyValue.Add "@@Numero", lNrofor
  dKeyValue.Add "@@NombreEmpresa", sNombreEmpresa
  dKeyValue.Add "@@DireccionEmpresa", sDirEmpresa
  dKeyValue.Add "@@PaisEmpresa", sCodPaiEmpresa
  dKeyValue.Add "@@CodPosEmpresa", sCodPosEmpresa
  dKeyValue.Add "@@LocalidadEmpresa", sLocalidadEmpresa
  dKeyValue.Add "@@Firmante", sFirmante
End Function
Function ReplaceAllKey(Resultado , diccionario)
  For Each skey in diccionario.Keys
     RegularExpression skey, ObjRegEx
     if Resultado = "Parrafos" then
        sPriPar = ObjRegEx.Replace(sPriPar, CStr(diccionario.Item(skey)))
        sSegPar = ObjRegEx.Replace(sSegPar, CStr(diccionario.Item(skey)))
        sTerPar = ObjRegEx.Replace(sTerPar, CStr(diccionario.Item(skey)))
     else
        Resultado = ObjRegEx.Replace(Resultado, CStr(diccionario.Item(skey)))
     End if
  Next
End Function
Function IsExists(sBody,clave,valor,tipo)
  if valor = ""  then
    sBody = Replace(sBody,clave,"")
    sBody = Replace(sBody,tipo,"none")
  else
    valida = "inline"
    sBody = Replace(sBody,clave,valor)
    sBody = Replace(sBody,tipo,"inline")
  End If
End Function
Function RegularExpression(pattern,ObjRegEx)
  Set ObjRegEx = CreateObject("VBScript.RegExp")
  ObjRegEx.Global = True  'Esto hace que no solo busque en la primera que encuentre
  ObjRegEx.IgnoreCase = true    'no sensible a las mayúsculas
  ObjRegEx.Pattern = CStr(pattern)
End Function
Function EnviaMail(sTexto, sCliente, sTituloMail, sendusername, sendpassword, smptserver, smtpserverport, sAttachment, MailDesde)

   Dim iMsg
   Dim iConf
   Dim Flds
   Dim SendTo
   Dim sFrom
   Dim sSqlMails, rdMail, L, sMails

   sMails = ""
   L = 0
  ' sMails = "harena@softland.com.ar"
    On Error Resume Next
     sSqlMails = "Select VTMCLC_DIREML DIREML From VTMCLC Where VTMCLC_NROCTA='" & sCliente & "' And IsNull(USR_VTMCLC_ENVMAI,'N')='S' "
     Set rdMail = oConn.Execute(CStr(sSqlMails))

     Do While Not rdMail.EOF
        L = L + 1
         If L = 1 Then
             sMails = rdMail("DIREML").value
         Else
             sMails = sMails & ";" & rdMail("DIREML").value
         End If
         rdMail.MoveNext
     Loop
     rdMail.Close
     Set rdMail = Nothing

   grabarLog_Archivo("Mail Prueba: "+sMails)

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    SendTo = ""
    sendtoCC = ""
    sendtoBCC = ""
    iConf.Load -1 ' CDO Source Defaults
    Set Flds = iConf.Fields

    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = sendusername
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = sendpassword
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smptserver
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpserverport
        .Update
    End With


     if sChequearContacto <> "" and unMailCheck = 1 then
    	 		grabarLog_Archivo("Contactos sin check para enviar Facturas... ")
    	 		grabarLog_Archivo(sChequearContacto)

    	        sNoTieneContacto = "S"
    	        sFrom = MailDesde
    	        sTitulo = "Cliente sin check para enviar Facturas por mail"
    	        With iMsg
    	            Set .Configuration = iConf
    	            .To = "facturaelectronica@irt-sa.com.ar"
    	            .From = sFrom
    	            .Subject = sTitulo
    	            .HTMLBody =  Replace(dElementHtml.Item("sHtml"),"#TABLE",sChequearContacto)
    	            .Send
    	        End With
          unMailCheck = 0
    	End if
    If Trim(sMails) <> "" Then ' el cliente tiene contacto marcado para envio de mail

        sFrom = MailDesde
        sNoTieneContacto = "N"
        sTitulo = sTituloMail

        With iMsg
            Set .Configuration = iConf
             .To = CStr(sMails)
    	   'Descomentar cuando se defina la cuenta que quieren usar como copia oculta para tener control de los mails
             '.Bcc = "softlanddar@gmail.com"
            .From = sFrom
    ' Utilizar en caso de querer que el mail se responda a una determinada direccion
            '.ReplyTo = "MAIL@DOMINIO.com.ar"
            .Subject = sTitulo
            .HTMLBody = sTexto
            FOR EACH sAttachment IN arrArchivos
            grabarLog_Archivo("sAttachment: "&sAttachment)
            .AddAttachment sAttachment
            NEXT
            .Send
        End With
    Else
    	grabarLog_Archivo("Mail rechazado: "+sMails)
        sNoTieneContacto = "S"
        sFrom = MailDesde
        sTitulo = "Cliente sin contacto para envio mail"
        With iMsg
            Set .Configuration = iConf
            .To = "facturaelectronica@irt-sa.com.ar"
            .From = sFrom
    ' Utilizar en caso de querer que el mail se responda a una determinada direccion
            .Subject = sTitulo
            .HTMLBody = "El cliente " & sCliente & " no tiene definido ningun contacto para envio de mail"
            .Send
        End With
    End If
  If Err.Number > 0 Then
       sNoTieneContacto = "S"
        sFrom = MailDesde
        sTitulo = "Problemas en el mail"
        With iMsg
            Set .Configuration = iConf
            .To = "facturaelectronica@irt-sa.com.ar"
            .From = sFrom
    ' Utilizar en caso de querer que el mail se responda a una determinada direccion
            .Subject = sTitulo
            .HTMLBody = "El cliente " & sCliente & " no tiene el mail: "& sMails
            .Send
        End With
  End If
End Function
Function GeneroPDFVT(Path, Cliente, Empresa, Modulo, Formulario, Numero, EmpresaMain, EmpresaMulti)
  Dim Reporte, rdPdf,oFso
  nArchivo = Path & "Saldos por Aplicación -" & CStr(Cliente) & ".pdf"

  Redim Preserve arrArchivos(contador)
  arrArchivos(contador) = nArchivo
  contador = contador +1

  Set oFso = CreateObject("scripting.filesystemObject")

  If oFso.FileExists(nArchivo) Then
      oFso.DeleteFile nArchivo
  End If



  Set oReport = oApplication.Companies(Empresa).GetObject(Trim(CStr("USR_VTR_MVC_SAA")), 5, "")
  oReport.Parameters("VTRMVC_NROCTA").ValueFrom = Cliente
  Set oRenderer = oReport.GetRenderer(1)
  oRenderer.PrinterName = "PDFCreator"
  oRenderer.PrintToFile = True
  oRenderer.OutputPath = nArchivo
  oRenderer.Render

End Function
Function GeneroPDFCV(Path, Cliente, Empresa, Modulo, Formulario, Numero, EmpresaMain, EmpresaMulti,sCodCom)
  Dim Reporte, rdPdf, oFso
  nArchivo = Path & "Detalle-" & Cliente & "-" & Formulario & "-" & CStr(Numero) & ".pdf"
  Redim Preserve arrArchivos(contador)
  arrArchivos(contador) = nArchivo
   contador = contador +1

  Set oFso = CreateObject("scripting.filesystemObject")

  If oFso.FileExists(nArchivo) Then
     oFso.DeleteFile nArchivo
  End If

  If instr(sCodCom,"FCA")<>0 or instr(sCodCom,"FCU")<>0 Then
    Reporte = "USR_CVRNVA"
   Else
   	If instr(sCodCom,"FCX")<>0 Then
    	Reporte = "USR_W_ANEXO210"
    Else
    	If instr(sCodCom,"FCP")<>0 Then
    		Reporte = "USR_ANEX_230"
    	End If
    End If
  End If

  Set oReport = oApplication.Companies(Empresa).GetObject(Trim(CStr(Reporte)), 5, "")
  oReport.Parameters("PARAM_CODEMP").ValueFrom = Empresa
  grabarLog_Archivo(Empresa)
  grabarLog_Archivo(Formulario)
  grabarLog_Archivo(Modulo)
  grabarLog_Archivo(Cstr(Numero))


  oReport.Parameters("PARAM_MODFOR").ValueFrom = Modulo
  oReport.Parameters("PARAM_CODFOR").ValueFrom = Formulario
  oReport.Parameters("PARAM_NROFOR").ValueFrom = Numero
  oReport.Parameters("PARAM_NROFOR").ValueTo = Numero
  Set oRenderer = oReport.GetRenderer(1)
  oRenderer.PrinterName = "PDFCreator"
  oRenderer.PrintToFile = True
  oRenderer.OutputPath = nArchivo
  oRenderer.Render

End Function
Function GeneroPDF(Path, Cliente, Empresa, Modulo, Formulario, Numero, EmpresaMain, EmpresaMulti, sCodCom)
    Dim Reporte, rdPdf, Rango, oFso

    Rango = 0
    nArchivo = Path & Cliente & "-" & Formulario & "-" & CStr(Numero) & ".pdf"

    Redim Preserve arrArchivos(contador)
    arrArchivos(contador) = nArchivo
    contador = contador + 1

    Set oFso = CreateObject("scripting.filesystemObject")

    If oFso.FileExists(nArchivo) Then
       oFso.DeleteFile nArchivo
    End If



    sSql = "Select GRCCBF_RPTNAM REPORTE From GRCCBF "
    sSql = sSql & " Where "
    If EmpresaMulti = "S" Then
        sSql = sSql & " GRCCBF_CODEMP='" & Empresa & "' And "
    End If
    sSql = sSql & " GRCCBF_MODFOR='" & Modulo & "' And "
    sSql = sSql & " GRCCBF_CODCOM='" & sCodCom & "' And "
    sSql = sSql & " GRCCBF_CODFOR='" & Formulario & "' "
    Set rdPdf = oConn.Execute(CStr(sSql))
    If Not rdPdf.EOF Then
        Reporte = rdPdf("REPORTE").value
    Else
        Select Case Modulo
            Case "VT"
                Reporte = "VTF_MVH_VLA"
            Case "PV"
                Reporte = "PVF_MVH_FOR"
            Case "CO"
                Reporte = "COR_MVH_FOC"
            Case "FC"
                Reporte = "FCF_MVH_IMP"
            Case "ST"
                Reporte = "STF_MVH_FOR"
            Case "SP"
                Reporte = "GPF_MVH_FOR"
        End Select
    End If
    rdPdf.Close
    Set rdPdf = Nothing

    Set oReport = oApplication.Companies(cNombre_Empresa).GetObject(Trim(CStr(Reporte)), 5, "")

    If EmpresaMulti = "S" And EmpresaMain = "S" Then
        sSqlCtrlPar = "Select Count(*) Existe FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_CODEMP'"
        Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
        If Not rdCtrlPar.EOF Then
            Existe = rdCtrlPar("Existe").value
        Else
            Existe = 0
        End If
        rdCtrlPar.Close
        Set rdCtrlPar = Nothing
        If Existe > 0 Then
            oReport.Parameters("" & Modulo & "RMVH_CODEMP").ValueFrom = Empresa
        End If
    End If
    sSqlCtrlPar = "Select Count(*) Existe FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_MODFOR'"
    Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
    If Not rdCtrlPar.EOF Then
        Existe = rdCtrlPar("Existe").value
    Else
        Existe = 0
    End If
    rdCtrlPar.Close
    Set rdCtrlPar = Nothing
    If Existe > 0 Then
        oReport.Parameters("" & Modulo & "RMVH_MODFOR").ValueFrom = Modulo
    End If
    sSqlCtrlPar = "Select Count(*) Existe FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_CODFOR'"
    Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
    If Not rdCtrlPar.EOF Then
        Existe = rdCtrlPar("Existe").value
    Else
        Existe = 0
    End If
    rdCtrlPar.Close
    Set rdCtrlPar = Nothing
    If Existe > 0 Then
        sSqlCtrlPar = "Select Convert(numeric(1),ParameterRange) Rango FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_CODFOR'"
        Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
        If Not rdCtrlPar.EOF Then
            Rango = rdCtrlPar("Rango").value
        Else
            Rango = 0
        End If
        rdCtrlPar.Close
        Set rdCtrlPar = Nothing
        oReport.Parameters("" & Modulo & "RMVH_CODFOR").ValueFrom = Formulario
        If CInt(Rango) = 1 Then
            oReport.Parameters("" & Modulo & "RMVH_CODFOR").ValueTo = Formulario
        End If
    End If

    sSqlCtrlPar = "Select Count(*) Existe FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_NROFOR'"
    Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
    If Not rdCtrlPar.EOF Then
        Existe = rdCtrlPar("Existe").value
    Else
        Existe = 0
    End If
    rdCtrlPar.Close
    Set rdCtrlPar = Nothing
    If Existe > 0 Then
        sSqlCtrlPar = "Select Convert(numeric(1),ParameterRange) Rango FROM CWRMFIELDS Where REPORTNAME='" & Reporte & "' AND ISPARAMETER = 1 And FieldID ='" & Modulo & "RMVH_NROFOR'"
        Set rdCtrlPar = oConn.Execute(CStr(sSqlCtrlPar))
        If Not rdCtrlPar.EOF Then
            Rango = rdCtrlPar("Rango").value
        Else
            Rango = 0
        End If
        rdCtrlPar.Close
        Set rdCtrlPar = Nothing
        oReport.Parameters("" & Modulo & "RMVH_NROFOR").ValueFrom = Numero
        If CInt(Rango) = 1 Then
            oReport.Parameters("" & Modulo & "RMVH_NROFOR").ValueTo = Numero
        End If

    End If
    Set oRenderer = oReport.GetRenderer(1)
    oRenderer.PrinterName = "PDFCreator"
    oRenderer.PrintToFile = True
    oRenderer.OutputPath = nArchivo
    oRenderer.Render
End Function
Function MarcoComprobantes(Empresa, Modulo, Formulario, Numero, EmpresaMulti)
    Dim sSqlMarco, rdMarco
    sSqlMarco = "Update "

    sSqlMarco = sSqlMarco & Modulo & "RMVH "
    sSqlMarco = sSqlMarco & " Set USR_" & Modulo & "RMVH_ENVMAI ='S' "
    sSqlMarco = sSqlMarco & " Where "
    If EmpresaMulti = "S" Then
        sSqlMarco = sSqlMarco & Modulo & "RMVH_CODEMP='" & Empresa & "' And "
    End If
    sSqlMarco = sSqlMarco & Modulo & "RMVH_MODFOR='" & Modulo & "' And "
    sSqlMarco = sSqlMarco & Modulo & "RMVH_CODFOR='" & Formulario & "' And "
    sSqlMarco = sSqlMarco & Modulo & "RMVH_NROFOR=" & Numero
    oConn.Execute (CStr(sSqlMarco))
End Function
Function normalValue(codigo1,codigo2)

  sSql1 = "SELECT NORMALVALUE "& codigo1 &" FROM CWPARAMETERS WHERE PARAMNAME='GR_"&codigo2&"'"
  If EsMulti = "S" Then
      sSql1 = sSql1 & " AND COMPANYNAME='" & cNombre_Empresa & "' "
  End If
  Set rd1 = oConn.Execute(CStr(sSql1))
  If Not rd1.EOF Then
    sValor =  rd1(codigo1).value
  End If
  rd1.Close
  Set rd1 = Nothing
   normalValue = sValor
End Function
Function openConnection()
  Set oConn = CreateObject("ADODB.Connection")
  DBProperties.CompanyName = cNombre_Empresa
  sUser = DBProperties.User               'Devuelve el usuario de la base de datos
  sPassWord = DBProperties.Password       'Devuelve la password de la base de datos
  sDatabase = DBProperties.Database       'Devuelve el nombre de la base de datos
  sServer = DBProperties.Server           'Devuelve el nombre del servidor de base de datos
  sDriver = DBProperties.Driver           'Devuelve el nombre del driver del servidor de base de datos
  oConn.Provider = "sqloledb"
  oConn.Properties("Data Source").value = sServer
  oConn.Properties("Initial Catalog").value = sDatabase
  oConn.Properties("User ID").value = sUser
  oConn.Properties("Password").value = sPassWord
  oConn.Open
End Function
Function closeConnection()
  oConn.Close
  Set oConn = Nothing
End Function
Function esMultiEmpresa()
  sSql = "SELECT ISNULL(ISMULTI,'N') MULTI, ISNULL(ISMAIN,'N') MAIN From CWSGCORE.DBO.CWOMCOMPANIES Where NAME='" & cNombre_Empresa & "'"
  Set rd = oConn.Execute(CStr(sSql))
  If Not rd.EOF Then
      sEsMulti = rd("MULTI").value
      sEsMain = rd("MAIN").value
  End If
  rd.Close
  Set rd = Nothing
End Function
Sub grabarLog_Archivo(pDato)
    Dim strArchivo, archivo, fso
    Dim ParaEscritura, ParaAnexar

    Set fso = CreateObject("Scripting.FileSystemObject")
    strArchivo = "C:\log\File_" + Replace(Replace(CStr(Date), "/", "-"), ":", ".") + ".log"
    If FileExists(strArchivo) Then
        ParaAnexar = 8
        Set archivo = fso.OpenTextFile(strArchivo, ParaAnexar, False)
    Else
        ParaEscritura = 2
        Set archivo = fso.CreateTextFile(strArchivo)
    End If

    archivo.Write (CStr(Now) + " - " + pDato + vbCRLF)
    archivo.Close
End Sub
Function FileExists(fileName)
    ' aqui On Error Resume Next
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(fileName) Then
        FileExists = True
    Else
        FileExists = False
    End If
    If Err > 0 Then Err.Clear
    Set objFso = Nothing
End Function
Function TieneAnexo(Empresa,Modulo,Formulario,Numero)
 sSql="SELECT count(USR_CVRNVA.USR_CVRNVA_IDENTI) as CANTIDAD "&_
			"FROM USR_CVRNVA "&_
			"WHERE ((USR_CVRNVA_IDENTI = (Select USR_VTRMVH_IDENTI from VTRMVH "&_
					 "Where VTRMVH_CODEMP = '"& Empresa &"' And "&_
					 "VTRMVH_MODFOR = '"&Modulo&"' And "&_
					 "VTRMVH_CODFOR = '"&Formulario&"' And VTRMVH_NROFOR ="&Numero&")) "&_
			   "or (USR_CVRNVA_IDENTI = (Select top 1 FCRMVI_NROORI from FCRMVI "&_
				 "Where FCRMVI_CODEMP = '"&Empresa&"' And "&_
				 "FCRMVI_MODFOR = '"&Modulo&"' And "&_
				 "FCRMVI_CODFOR = '"&Formulario&"' And FCRMVI_NROFOR = "&Numero&")))"


	Set rd3  = oConn.Execute(CStr(sSql))
	If not rd3.EOF then
		 sCount = rd3("CANTIDAD").value
	End if
	TieneAnexo = sCount

End Function
Function CreateDictionaryElementHtml()
      Set dElementHtml = CreateObject("Scripting.Dictionary")
      dElementHtml.Add "sTempTable","<table id = ""Table"" cellpadding='0' cellspacing='0'"&_
                     "align=""center"" width='100%'> #TABLE# </table>"
      dElementHtml.Add "sThead", "<th style =""background-color:  #094293;	color: #fff;font-family: Poppins-Regular;font-weight: unset !important; font-size: 18px; border-bottom: 1px solid #094293;text-align: left; padding: 8px;"">#HCOLNAME</th>"
      dElementHtml.Add "sTRow", "<td style =""font-family: Poppins-Regular;font-weight: unset !important;color: #666666; font-size: 18px; border-bottom: 1px solid #4040ff;text-align: left; padding: 8px;"">#RCOLNAME</td>"
      dElementHtml.Add "sHtml","<!doctype html> <html lang=""en""> <head> <title> Mail de Errores</title>  <meta charset=""utf-8""> <meta name=""viewport"" content=""width=device-width, initial-scale=1, shrink-to-fit=no"">  </head>  <body><h2 align=""center""><b><u> Empresas sin check de envio de Facturacion Electronica</b></u></h2><br>  #TABLE  </body> </html> "
End Function
Function chequearContactoSinCheck()
  Dim rdMail2
   grabarLog_Archivo("Chequeando mails...")


  sSqlMails2 = "SELECT  Distinct "&_
              "VTRMVH_NROCTA NROCTAM, "&_
              "VTRMVH_NROSUB NROCTA, "&_
              "VTMCLH_NOMBRE NOMBRE, "&_
              "VTRMVH_CODEMP CODEMP "&_
              "FROM VTRMVH "&_
              "INNER JOIN VTMCLH ON VTRMVH_NROCTA=VTMCLH_NROCTA "&_
              "INNER JOIN VTMCLC ON VTMCLC_NROCTA=VTMCLH_NROCTA AND ISNULL(USR_VTMCLC_ENVMAI,'N')='N' "&_
              "INNER JOIN GRCFOR ON GRCFOR_MODFOR=VTRMVH_MODFOR and	"&_
              "GRCFOR_CODFOR=VTRMVH_CODFOR "&_
              "WHERE	ISNULL(USR_VTRMVH_ENVMAI,'N')='N' And "&_
              "ISNULL(USR_VTMCLH_ENVMAI,'N')='S' And "&_
              "ISNULL(USR_GRCFOR_ENVMAI,'N')='S' "&_
              "And VTRMVH_NROCAE is not null "&_
              "and VTRMVH_NROSUB not in (SELECT  Distinct VTRMVH_NROSUB NROCTA "&_
						            "FROM VTRMVH  "&_
						            "INNER JOIN VTMCLH ON VTRMVH_NROCTA=VTMCLH_NROCTA "&_
						            "INNER JOIN VTMCLC ON VTMCLC_NROCTA=VTMCLH_NROCTA AND ISNULL(USR_VTMCLC_ENVMAI,'N')='S'  "&_
						            "INNER JOIN GRCFOR ON GRCFOR_MODFOR=VTRMVH_MODFOR and	"&_
						            "GRCFOR_CODFOR=VTRMVH_CODFOR  "&_
						            "WHERE	ISNULL(USR_VTRMVH_ENVMAI,'N')='N' And  "&_
						            "ISNULL(USR_VTMCLH_ENVMAI,'N')='S' And  "&_
						            "ISNULL(USR_GRCFOR_ENVMAI,'N')='S'  "&_
						            "And VTRMVH_NROCAE is not null) "&_
	          "ORDER BY NROCTAM, NROCTA "
      Set rdMail2 = oConn.Execute(CStr(sSqlMails2))

      sHeader ="<tr>"
      sHeader = sHeader & Replace(dElementHtml.Item("sThead"),"#HCOLNAME","Nombre")
      sHeader = sHeader & Replace(dElementHtml.Item("sThead"),"#HCOLNAME","NroCta")
      sHeader = sHeader & Replace(dElementHtml.Item("sThead"),"#HCOLNAME","SubCuenta")
      sHeader = sHeader & Replace(dElementHtml.Item("sThead"),"#HCOLNAME","CodEmp")
      sHeader = sHeader &"</tr>"
          aRow =""

      Do While Not rdMail2.EOF
          sRow = "<tr>"
          sRow = sRow & Replace(dElementHtml.Item("sTRow"),"#RCOLNAME",rdMail2("NOMBRE").value)
          sRow = sRow & Replace(dElementHtml.Item("sTRow"),"#RCOLNAME",rdMail2("NROCTA").value)
          sRow = sRow & Replace(dElementHtml.Item("sTRow"),"#RCOLNAME",rdMail2("NROCTAM").value)
          sRow = sRow & Replace(dElementHtml.Item("sTRow"),"#RCOLNAME",rdMail2("CODEMP").value)
          sRow = sRow & "</tr>"
          aRow = aRow & sRow
          sRow = "<tr>"
          rdMail2.MoveNext
      Loop
      rdMail2.Close
      Set rdMail2 = Nothing
      sTableContent = sHeader & aRow
      stable =  Replace(dElementHtml.Item("sTempTable"),"#TABLE#",sTableContent)

      chequearContactoSinCheck = stable

End Function
