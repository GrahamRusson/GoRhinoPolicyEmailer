Imports System
Imports System.IO
Imports System.Text
Imports System.Reflection
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports TheArtOfDev.HtmlRenderer
Imports System.Net
Imports SendGrid

Public Class Email

    Dim _clientInformation As ClientInformation
    Dim _goRhinoPackage As GoRhinoPackage
    Dim _emailInformation As EmailInformation
    Dim _statusMessage As String = ""

    Public Sub New(clientInformation As ClientInformation, packageInformation As GoRhinoPackage, emailInformation As EmailInformation)
        _clientInformation = clientInformation
        _goRhinoPackage = packageInformation
        _emailInformation = emailInformation
    End Sub

    Property ClientInformation() As ClientInformation
        Get
            Return _clientInformation
        End Get
        Set(ByVal Value As ClientInformation)
            _clientInformation = Value
        End Set
    End Property

    Property GoRhinoPackage() As GoRhinoPackage
        Get
            Return _goRhinoPackage
        End Get
        Set(ByVal Value As GoRhinoPackage)
            _goRhinoPackage = Value
        End Set
    End Property

    Property EmailInformation() As EmailInformation
        Get
            Return _emailInformation
        End Get
        Set(ByVal Value As EmailInformation)
            _emailInformation = Value
        End Set
    End Property

    ReadOnly Property StatusMessage() As String
        Get
            Return _statusMessage
        End Get
    End Property

    Public Function SendMail(RootPath As String) As Boolean

        Try

            ' To create the PdfContentTemplate.html file, if it changes, just "view source" in your browser, and then replace the content of this file, and it becomes the new template for generating the PDF. The placeholders get replaced below, then saved to the new PDF file.

            'Dim rootPath As String = (New System.Uri(Assembly.GetExecutingAssembly().CodeBase)).AbsolutePath
            'RootPath = rootPath.ToLower().Replace("gorhinopolicyemailer.dll", "")

            Dim path As String = ""

            If _goRhinoPackage.SelectedPackage = GoRhinoPackage.YouAssistPackages.Arrive Then
                path = rootPath & "\PdfContentTemplateForPackage_Arrive.html"
            ElseIf _goRhinoPackage.SelectedPackage = GoRhinoPackage.YouAssistPackages.Drive Then
                path = rootPath & "\PdfContentTemplateForPackage_Drive.html"
            ElseIf _goRhinoPackage.SelectedPackage = GoRhinoPackage.YouAssistPackages.ER Then
                path = rootPath & "\PdfContentTemplateForPackage_ER.html"
            ElseIf _goRhinoPackage.SelectedPackage = GoRhinoPackage.YouAssistPackages.Protect Then
                path = rootPath & "\PdfContentTemplateForPackage_Protect.html"
            Else
                Throw New Exception("A valid policy package was not specified.")
            End If

            ' Open the file and read.
            Dim htmlString As String = File.ReadAllText(path)

            ' Add the information...
            htmlString = htmlString.Replace("[<span>client name]", "<span>" & ClientInformation.ClientName) ' This happens because this tag has a span included in it. Fix this in the template to make things easier to deal with!
            htmlString = htmlString.Replace("[client_policy_number]", ClientInformation.ClientPolicyNumber)
            htmlString = htmlString.Replace("[sub_plan_name]", ClientInformation.SubPlanName)
            htmlString = htmlString.Replace("[Price]", ClientInformation.Price)
            htmlString = htmlString.Replace("[client_full_names]", ClientInformation.ClientFullNames)
            htmlString = htmlString.Replace("[client_id]", ClientInformation.ClientId)
            htmlString = htmlString.Replace("[client_phone]", ClientInformation.ClientPhone)
            htmlString = htmlString.Replace("[client_email]", ClientInformation.ClientEmail)
            htmlString = htmlString.Replace("[client_address]", ClientInformation.ClientAddress)

            Dim fileName As String = rootPath & "\YouAssistPolicy_" & ClientInformation.ClientPolicyNumber & ".pdf"

            Dim configuration As TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig = New TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerateConfig()
            configuration.PageSize = PageSize.Letter
            configuration.MarginLeft = 25
            configuration.MarginRight = 25
            Dim doc As PdfSharp.Pdf.PdfDocument = TheArtOfDev.HtmlRenderer.PdfSharp.PdfGenerator.GeneratePdf(htmlString, configuration)
            doc.Save(fileName)
            doc.Close()

            Dim msg = New SendGridMessage()

            msg.From = New Net.Mail.MailAddress(_emailInformation.SentFromEmailAddress, _emailInformation.SentFromEmailAddressTitle)
            msg.AddTo(ClientInformation.ClientEmail)

            msg.Subject = _emailInformation.EmailSubject

            msg.Html = _emailInformation.EmailBody
            'msg.Text = "Hello world" ' I don't think we are sending TEXT emails, but leaving this note here just in case.

            msg.AddAttachment(fileName)

            Dim credentials As NetworkCredential = New NetworkCredential(_emailInformation.SENDGRID_USERNAME, _emailInformation.SENDGRID_PASSWORD)
            Dim client = New SendGrid.Web(credentials)
            client.Deliver(msg) '.DeliverAsync(msg)

            _statusMessage = "Method SendMail() completed execution successfully."
            Return True

        Catch ex As Exception
            _statusMessage = "Exception in SendMail(): " & ex.Message & ", " & ex.StackTrace
            Return False
        End Try

    End Function

End Class

Public Class GoRhinoPackage

    Dim _selectedPackage As YouAssistPackages

    Property SelectedPackage() As Integer
        Get
            Return _selectedPackage
        End Get
        Set(ByVal Value As Integer)
            _selectedPackage = Value
        End Set
    End Property

    Public Enum YouAssistPackages
        ER
        Protect
        Drive
        Arrive
    End Enum

End Class

Public Class ClientInformation

    Dim _clientPolicyNumber As String = "ClientPolicyNumber"
    Dim _subPlanName As String = "SubPlanName"
    Dim _price As String = "Price"
    Dim _clientName As String = "ClientName"
    Dim _clientFullNames As String = "ClientFullNames"
    Dim _clientPhone As String = "ClientPhone"
    Dim _clientId As String = "ClientId"
    Dim _clientEmail As String = "ClientEmail"
    Dim _clientAddress As String = "ClientAddress"

    Property ClientPolicyNumber() As String
        Get
            Return _clientPolicyNumber
        End Get
        Set(ByVal Value As String)
            _clientPolicyNumber = Value
        End Set
    End Property

    Property SubPlanName() As String
        Get
            Return _subPlanName
        End Get
        Set(ByVal Value As String)
            _subPlanName = Value
        End Set
    End Property

    Property Price() As String
        Get
            Return _price
        End Get
        Set(ByVal Value As String)
            _price = Value
        End Set
    End Property

    Property ClientName() As String
        Get
            Return _clientName
        End Get
        Set(ByVal Value As String)
            _clientName = Value
        End Set
    End Property

    Property ClientFullNames() As String
        Get
            Return _clientFullNames
        End Get
        Set(ByVal Value As String)
            _clientFullNames = Value
        End Set
    End Property

    Property ClientPhone() As String
        Get
            Return _clientPhone
        End Get
        Set(ByVal Value As String)
            _clientPhone = Value
        End Set
    End Property

    Property ClientId() As String
        Get
            Return _clientId
        End Get
        Set(ByVal Value As String)
            _clientId = Value
        End Set
    End Property

    Property ClientEmail() As String
        Get
            Return _clientEmail
        End Get
        Set(ByVal Value As String)
            _clientEmail = Value
        End Set
    End Property

    Property ClientAddress() As String
        Get
            Return _clientAddress
        End Get
        Set(ByVal Value As String)
            _clientAddress = Value
        End Set
    End Property

End Class

Public Class EmailInformation

    Dim _SENDGRID_USERNAME As String = ""
    Dim _SENDGRID_PASSWORD As String = ""
    Dim _sentFromEmailAddress As String = ""
    Dim _sentFromEmailAddressTitle As String = ""
    Dim _emailSubject As String = ""
    Dim _emailBody As String = ""

    Property SENDGRID_USERNAME() As String
        Get
            Return _SENDGRID_USERNAME
        End Get
        Set(ByVal Value As String)
            _SENDGRID_USERNAME = Value
        End Set
    End Property

    Property SENDGRID_PASSWORD() As String
        Get
            Return _SENDGRID_PASSWORD
        End Get
        Set(ByVal Value As String)
            _SENDGRID_PASSWORD = Value
        End Set
    End Property

    Property SentFromEmailAddress() As String
        Get
            Return _sentFromEmailAddress
        End Get
        Set(ByVal Value As String)
            _sentFromEmailAddress = Value
        End Set
    End Property

    Property SentFromEmailAddressTitle() As String
        Get
            Return _sentFromEmailAddressTitle
        End Get
        Set(ByVal Value As String)
            _sentFromEmailAddressTitle = Value
        End Set
    End Property

    Property EmailSubject() As String
        Get
            Return _emailSubject
        End Get
        Set(ByVal Value As String)
            _emailSubject = Value
        End Set
    End Property

    Property EmailBody() As String
        Get
            Return _emailBody
        End Get
        Set(ByVal Value As String)
            _emailBody = Value
        End Set
    End Property

End Class