Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.X509Certificates
Imports System.Xml
Imports Ionic.Zip
Imports System.Runtime.InteropServices
Imports System.Xml.Schema
Imports System.Windows.Forms

Namespace firmadoCE
    <ClassInterface(ClassInterfaceType.AutoDual)>
    <ProgId("firmadoCE.firmado")>
    Public Class firmado
        Public Function firmar(ByVal cRutaArchivo As String, ByVal cArchivo As String, ByVal cCertificado As String, ByVal cClave As String) As String
            Dim cRpta As String

            Try
                Dim local_typoDocumento As String = cArchivo.Substring(12, 2) ' retorna 01 o 03 0 ...
                Dim l_xpath As String = ""
                Dim f_certificat As String = cCertificado
                Dim f_pwd As String = cClave
                Dim xmlFile As String = cRutaArchivo & cArchivo
                Dim MonCertificat As X509Certificate2 = New X509Certificate2(f_certificat, f_pwd)
                Dim xmlDoc As XmlDocument = New XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(xmlFile)
                Dim signedXml As SignedXml = New SignedXml(xmlDoc)
                signedXml.SigningKey = MonCertificat.PrivateKey
                Dim KeyInfo As KeyInfo = New KeyInfo()
                Dim Reference As Reference = New Reference()
                Reference.Uri = ""
                Reference.AddTransform(New XmlDsigEnvelopedSignatureTransform())
                signedXml.AddReference(Reference)
                Dim X509Chain As X509Chain = New X509Chain()
                X509Chain.Build(MonCertificat)
                Dim local_element As X509ChainElement = X509Chain.ChainElements(0)
                Dim x509Data As KeyInfoX509Data = New KeyInfoX509Data(local_element.Certificate)
                Dim subjectName As String = local_element.Certificate.Subject
                x509Data.AddSubjectName(subjectName)
                KeyInfo.AddClause(x509Data)
                signedXml.KeyInfo = KeyInfo
                signedXml.ComputeSignature()
                Dim signature As XmlElement = signedXml.GetXml()
                signature.Prefix = "ds"
                signedXml.ComputeSignature()
                For Each node As XmlNode In signature.SelectNodes("descendant-or-self::*[namespace-uri()='http://www.w3.org/2000/09/xmldsig#']")
                    'node.Prefix = "ds"
                    If node.LocalName = "Signature" Then
                        Dim newAttribute As XmlAttribute = xmlDoc.CreateAttribute("Id")
                        newAttribute.Value = "SignatureSP"
                        node.Attributes.Append(newAttribute)
                    End If
                Next node
                Dim nsMgr As XmlNamespaceManager
                nsMgr = New XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("sac", "urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1")
                nsMgr.AddNamespace("ccts", "urn:un:unece:uncefact:documentation:2")
                nsMgr.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")

                Select Case local_typoDocumento
                    Case "01", "03" 'FACTURA Y BOLETA
                        nsMgr.AddNamespace("tns", "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2")
                        l_xpath = "/tns:Invoice/ext:UBLExtensions/ext:UBLExtension[1]/ext:ExtensionContent"
                    Case "07" 'NOTA DE CREDITO
                        nsMgr.AddNamespace("tns", "urn:oasis:names:specification:ubl:schema:xsd:CreditNote-2")
                        l_xpath = "/tns:CreditNote/ext:UBLExtensions/ext:UBLExtension[1]/ext:ExtensionContent"
                    Case "08" 'NOTA DE DEBITO
                        nsMgr.AddNamespace("tns", "urn:oasis:names:specification:ubl:schema:xsd:DebitNote-2")
                        l_xpath = "/tns:DebitNote/ext:UBLExtensions/ext:UBLExtension[1]/ext:ExtensionContent"
                    Case "09" 'GUIA DE REMISION REMITENTE
                        nsMgr.AddNamespace("tns", "urn:oasis:names:specification:ubl:schema:xsd:DespatchAdvice-2")
                        l_xpath = "/tns:DespatchAdvice/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                    Case "20" 'RETENCION
                        nsMgr.AddNamespace("tns", "urn:sunat:names:specification:ubl:peru:schema:xsd:Retention-1")
                        l_xpath = "/tns:Retention/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                    Case "31" 'GUIA DE REMISION REMITENTE TRANSPORTISTA
                        nsMgr.AddNamespace("tns", "urn:oasis:names:specification:ubl:schema:xsd:DespatchAdvice-2")
                        l_xpath = "/tns:DespatchAdvice/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                    Case "40" 'PERCEPCION
                        nsMgr.AddNamespace("tns", "urn:sunat:names:specification:ubl:peru:schema:xsd:Perception-1")
                        l_xpath = "/tns:Perception/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                    Case "RA" 'COMUNICACION DE BAJA
                        nsMgr.AddNamespace("tns", "urn:sunat:names:specification:ubl:peru:schema:xsd:VoidedDocuments-1")
                        l_xpath = "/tns:VoidedDocuments/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                    Case "RC" 'RESUMEN DIARIO
                        nsMgr.AddNamespace("tns", "urn:sunat:names:specification:ubl:peru:schema:xsd:SummaryDocuments-1")
                        l_xpath = "/tns:SummaryDocuments/ext:UBLExtensions/ext:UBLExtension/ext:ExtensionContent"
                End Select
                nsMgr.AddNamespace("cac", "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2")
                nsMgr.AddNamespace("udt", "urn:un:unece:uncefact:data:specification:UnqualifiedDataTypesSchemaModule:2")
                nsMgr.AddNamespace("ext", "urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2")
                nsMgr.AddNamespace("qdt", "urn:oasis:names:specification:ubl:schema:xsd:QualifiedDatatypes-2")
                nsMgr.AddNamespace("cbc", "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2")
                nsMgr.AddNamespace("ds", "http://www.w3.org/2000/09/xmldsig#")
                xmlDoc.SelectSingleNode(l_xpath, nsMgr).AppendChild(xmlDoc.ImportNode(signature, True))
                xmlDoc.Save(xmlFile)
                Dim nodeList As XmlNodeList = xmlDoc.GetElementsByTagName("ds:Signature")
                If (nodeList.Count <> 1) Then
                    cRpta = "SE PRODUJO ERROR EN LA FIRMA"
                End If
                signedXml.LoadXml(CType(nodeList(0), XmlElement))
                If (signedXml.CheckSignature()) = False Then            ' verificacion de la firma generada
                    cRpta = "SE PRODUJO UN ERROR EN LA FIRMA  DE DOCUMENTO"
                Else
                    cRpta = "OK"
                End If
            Catch ex As Exception
                cRpta = ex.Message
            End Try

            Return cRpta
        End Function

        Public Function comprimir(ByVal cnombrearchivoOrigen As String, ByVal cnombreArchivoDestino As String) As String
            Dim zip As Ionic.Zip.ZipFile = New Ionic.Zip.ZipFile()
            zip.AddFile(cnombrearchivoOrigen, "") ' se puede seguir agregando mas con a misma funcion
            zip.Save(cnombreArchivoDestino)
            Dim rpta As String = "OK"
            Return rpta
        End Function
        Public Function extraer(ByVal cnombrearchivoOrigen As String, ByVal ruta As String) As String
            Dim rpta As String
            rpta = "OK"
            Try
                Dim zip = ZipFile.Read(cnombrearchivoOrigen)
                zip.ExtractAll(ruta, True)
                zip.Dispose()
            Catch ex As Exception
                rpta = ex.Message
            End Try
            Return rpta
        End Function
        Public Function pdf417(ByVal ctextoCodificar As String, ByVal cruta As String, ByVal cnombreArchivoPNG As String, ByVal cbarcode As String, ByVal nAncho As Integer, ByVal nAlto As Integer, ByVal nMargen As Integer) As String
            Dim options As New ZXing.QrCode.QrCodeEncodingOptions
            options.DisableECI = True
            options.CharacterSet = "UTF-8"
            options.Width = nAncho
            options.Height = nAlto
            options.Margin = nMargen
            Dim qr As New ZXing.BarcodeWriter()
            qr.Options = options
            If cbarcode = "1" Then
                qr.Format = ZXing.BarcodeFormat.PDF_417
            End If
            If cbarcode = "2" Then
                qr.Format = ZXing.BarcodeFormat.QR_CODE
            End If
            If cbarcode = "3" Then
                qr.Format = ZXing.BarcodeFormat.DATA_MATRIX
            End If
            Dim result As New System.Drawing.Bitmap(qr.Write(Trim(ctextoCodificar)))
            result.Save(cruta & cnombreArchivoPNG)
            Dim rpta As String
            rpta = "ok"
            Return rpta
        End Function
        Public Resultado As Boolean
        Public txtresultado As String
        Private Sub AdminEventoValidacion(sender As Object, args As ValidationEventArgs)
            Resultado = False
        End Sub
        Public Function validarXML(ByVal txtxml As String, ByVal txtxsd As String) As String
            Resultado = True
            txtresultado = "OK"
            'DESDE EL LLAMADO SE DEBERA VALIDAR LA EXISTENCIA DE LOS ARCHIVOS ENVIADOS COMO PARAMETROS
            Dim xmlR As New XmlTextReader(txtxml)
            Dim xsdR As New XmlValidatingReader(xmlR)
            Try
                xsdR.Schemas.Add(Nothing, txtxsd)
                xsdR.ValidationType = ValidationType.Schema
                AddHandler xsdR.ValidationEventHandler, New ValidationEventHandler(AddressOf AdminEventoValidacion)
                While xsdR.Read()
                    Application.DoEvents()
                End While
                xsdR.Close()
            Catch ex As Exception
                'txtresultado = If(Resultado, "Archivo XML correcto con respecto al esquema XSD", ex.Message)
                txtresultado = ex.Message
            End Try
            ' SI NO HAY NINGUN ERROR REVOLVERA OK
            xsdR.Close()
            Return txtresultado
        End Function
        Public Function FTPenviar_archivoRemoto(ByVal rutaarchivoftp As String, userftp As String, ByVal passftp As String, nombrearchivolocal As String) As String
            Dim miUri As String = rutaarchivoftp
            Dim miRequest As Net.FtpWebRequest = Net.WebRequest.Create(miUri)
            miRequest.Credentials = New Net.NetworkCredential(userftp, passftp)
            miRequest.Method = Net.WebRequestMethods.Ftp.UploadFile
            Try
                Dim bFile() As Byte = System.IO.File.ReadAllBytes(nombrearchivolocal)
                Dim miStream As System.IO.Stream = miRequest.GetRequestStream()
                miStream.Write(bFile, 0, bFile.Length)
                miStream.Close()
                miStream.Dispose()
                Return "OK"
            Catch ex As Exception
                Return "ERROR-" & ex.Message
            End Try
        End Function
        Public Function FTPverificar_archivoRemoto(ByVal rutayarchivo As String, ByVal userftp As String, ByVal passftp As String) As String
            Dim miUri As String = rutayarchivo    '  dir    "ftp://ftp.midominio.com/carpeta/fichero.jpg"
            Dim miRequest As Net.FtpWebRequest = Net.WebRequest.Create(miUri)
            miRequest.Credentials = New Net.NetworkCredential(userftp, passftp)
            miRequest.Method = Net.WebRequestMethods.Ftp.GetFileSize
            Try
                Dim response As Net.FtpWebResponse = miRequest.GetResponse()
                ' THE FILE EXISTS
            Catch ex As Net.WebException
                Dim response As Net.FtpWebResponse = ex.Response
                If Net.FtpStatusCode.ActionNotTakenFileUnavailable = response.StatusCode Then
                    ' THE FILE DOES NOT EXIST
                    Return "ERROR"
                End If
            End Try
            Return "OK"
        End Function
        Public Function EnviarEmail(ByVal razonsocialemisor As String, ByVal AsuntoEmail As String, ByVal mailEnvio As String, ByVal claveemail As String, ByVal servidorEnvio As String, ByVal portEnvioSMTP As String, ByVal emailestino As String, ByVal archivoadjunto As String, archivoadjunto2 As String, archivoadjunto3 As String) As String
            Dim archivo As New System.Net.Mail.Attachment(archivoadjunto)
            Dim archivo2 As New System.Net.Mail.Attachment(archivoadjunto2)
            Dim archivo3 As New System.Net.Mail.Attachment(archivoadjunto3)
            Dim Message As New System.Net.Mail.MailMessage()
            Dim SMTP As New System.Net.Mail.SmtpClient
            'CONFIGURACIÓN DEL STMP
            '---------'("cuenta de correo", "contraseña")
            SMTP.Credentials = New System.Net.NetworkCredential(mailEnvio, claveemail)
            SMTP.Host = servidorEnvio
            SMTP.Port = portEnvioSMTP
            SMTP.EnableSsl = True
            ' CONFIGURACION DEL MENSAJE
            Message.[To].Add(emailestino) ' Acá se escribe la cuenta de correo al que se le quiere enviar el e-mail
            '-------"Quien lo envía","Nombre de quien lo envía"  
            Message.From = New System.Net.Mail.MailAddress(mailEnvio, claveemail, System.Text.Encoding.UTF8) 'Quien envía el e-mail
            Message.Subject = razonsocialemisor & AsuntoEmail
            Message.SubjectEncoding = System.Text.Encoding.UTF8 'Codificacion
            Message.IsBodyHtml = True
            Message.Body = "<font size=10>Se le ha enviado una nueva factura electronica</font> <font color=red><b>a test</b></font>"
            Message.Body = "<TABLE border = 4 cellspacing = 4 cellpadding = 4 width =80%> <TH align = center> " & razonsocialemisor & " <TR> <TD align = LEFT>" & "Se le ha enviado una facura electronica" & "</TABLE> "

            Message.BodyEncoding = System.Text.Encoding.UTF8
            Message.Priority = System.Net.Mail.MailPriority.Normal
            '------- Message.IsBodyHtml = False
            Message.Attachments.Add(archivo)
            Message.Attachments.Add(archivo2)
            Message.Attachments.Add(archivo3)
            'ENVIO
            Try
                SMTP.Send(Message)
                Return "OK"
            Catch ex As System.Net.Mail.SmtpException
                Return "ERROR-" & ex.ToString
            End Try
        End Function
        Public Function isOnline() As String
            Dim Url As New System.Uri("https://www.google.com")
            Dim oWebReq As System.Net.WebRequest
            oWebReq = System.Net.WebRequest.Create(Url)
            Dim oResp As System.Net.WebResponse
            Try
                oResp = oWebReq.GetResponse
                oResp.Close()
                oWebReq = Nothing
                Return "OK"
            Catch ex As Exception
                oResp.Close()
                oWebReq = Nothing
                Return "ERROR- NO HAY CONEXION A INTERNET"
            End Try
        End Function
    End Class
End Namespace