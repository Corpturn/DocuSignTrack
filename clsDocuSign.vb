Imports System.Net.Http
Imports System.Net.Http.Headers
Imports Newtonsoft.Json
Imports System.IO
Imports System.Data.SqlClient
Imports DocuSign.eSign.Client
Imports DocuSign.eSign.Api
Imports DocuSign.eSign.Model

Public Class clsDocuSign
    Public Property envelopeId As String
    Public Property status As String
    Public Property emailSubject As String
    Public Property createdDateTime As DateTime
    Public Property completedDateTime As Nullable(Of DateTime)

    Public Shared DB_SVR As String = "ID003"
    Public Shared strCatalog As String = "CTDB"

    Public Shared CTDOC_Test As String = "\\ID003\D$\CTDocTest\"
    Public Shared CTDOC = "\\ID003\D$\CTDoc\"
    Public Shared SavePath As String
    Public Shared AppPath As String = AppContext.BaseDirectory

    Public Shared strConnStr As String = "Data Source=" & DB_SVR & ";Initial Catalog=" & strCatalog & ";Trusted_Connection=True"
    Public Shared IntegrationKey As String = "2071c185-f980-4850-9a93-dcb5ff118d62"
    Public Shared UserId As String = "635da419-6b5f-44fa-81f6-7f56738ef3db" 'GUID of user
    Public Shared AuthServer As String = "account.docusign.com" 'demo: account-d.docusign.com

    Public Shared privateKeyBytes As Byte()
    Public Shared keypairID As String = "81d49729-b04e-4e7c-970d-850fd91c6ac8"

    Public Shared AccessToken As String

    Private Shared client As New HttpClient()

    Public Shared Sub ProcessEnvelopes()
        Try
            SendProcessReport("Starting DocuSign envelope processing...")
            InitializeVariables()

            'get outstanding envelopes from database
            Console.WriteLine("Getting outstanding envelopes!")
            Dim dsEnvelopes As DataTable = GetDocuSignEnvelopes()
            Dim EnvelopeCount As Integer = dsEnvelopes.Rows.Count

            Console.WriteLine("Initializing DocuSign API references!")

            'client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", AccessToken)

            Dim apiClient As New DocuSignClient()
            apiClient.SetOAuthBasePath("account.docusign.com")

            'Get user info
            Dim userInfo = apiClient.GetUserInfo(AccessToken)

            'Get AccountId
            Dim AccountId As String = userInfo.Accounts(0).AccountId

            ''Get BaseUrl
            Dim BaseUri As String = userInfo.Accounts(0).BaseUri & "/restapi"

            '=====================================================
            Console.WriteLine("Creating API objects And authenticating account!")
            'Create API client
            Dim apiClient2 As New DocuSignClient(BaseUri)
            'Attach token
            apiClient2.Configuration.DefaultHeader("Authorization") = "Bearer " & AccessToken
            Dim mEnvelopesApi As New EnvelopesApi(apiClient2)

            Dim mDateViewed As Nullable(Of DateTime)
            Dim mDateCompleted As Nullable(Of DateTime)
            Dim mDateDeclined As Nullable(Of DateTime)
            Dim mAttachedDocID As Integer


            'then go through each envelope and check status via API
            For Each row As DataRow In dsEnvelopes.Rows
                Dim EnvelopeId As String = row("EnvelopeID").ToString()
                Dim mDatesent As Date = row("DateSent")
                Dim mID As String = row("LeadID").ToString()
                Dim mDocDesc As String = row("DocumentDesc").ToString()

                Console.WriteLine("Processing DocuSign envelopes!")

                Try
                    SavePath = "C:\PayOff\TempFax\" & row("DocumentDesc") & "_" & mDatesent.ToString("yyyy_MM_dd") & ".pdf"

                    Dim mEnvelope = mEnvelopesApi.GetEnvelope(AccountId, EnvelopeId)

                    If mEnvelope.Status.ToLower() = "delivered" Then
                        mDateViewed = DateTime.Parse(mEnvelope.DeliveredDateTime)
                    End If

                    If mEnvelope.Status.ToLower() = "declined" Then
                        mDateDeclined = DateTime.Parse(mEnvelope.DeclinedDateTime)
                    End If

                    If mEnvelope.Status.ToLower() = "completed" Then
                        mDateCompleted = DateTime.Parse(mEnvelope.CompletedDateTime)
                    End If


                    'If mEnvelope.Status.ToLower() = "completed" Then
                    '    mDateCompleted = DateTime.Parse(mEnvelope.CompletedDateTime)
                    'Else
                    '    If mEnvelope.Status.ToLower() = "delivered" Then
                    '        mDateViewed = DateTime.Parse(mEnvelope.StatusChangedDateTime)
                    '    End If
                    'End If

                    If IsNothing(mDateCompleted) = False Or IsNothing(mDateViewed) = False Then

                        If IsDBNull(mDateCompleted) = False Then

                            'Download PDF
                            DownloadEnvelopePdf(AccountId, EnvelopeId, SavePath)

                            'attach the to the Lead/Client in the database 
                            mAttachedDocID = AttachDocument(mID, mDocDesc)

                            'Update the docusing data table to indicate that this envelope has been signed and attached
                            UpdateEnvelopeStatus(EnvelopeId, mDateViewed, mDateCompleted, mDateDeclined, mAttachedDocID)
                        Else
                            'Update the docusing data table to indicate that this envelope has been signed and attached
                            UpdateEnvelopeStatus(EnvelopeId, mDateViewed, mDateCompleted, mDateDeclined, mAttachedDocID)
                        End If
                        Console.WriteLine($"Downloaded PDF for {EnvelopeId}")
                    End If

                Catch ex As Exception
                    Console.WriteLine($"Error with {EnvelopeId}: {ex.Message}")
                    Console.ReadKey()
                End Try
            Next

            SendProcessReport("DocuSign envelope processing completed!")
        Catch ex As Exception
            SendProcessReport(ex.Message)
            Console.WriteLine("ProcessEnvelopes " & ex.Message)
            Console.ReadKey()
        End Try

    End Sub

    Public Shared Sub InitializeVariables()
        Try
            privateKeyBytes = File.ReadAllBytes(AppPath & "Keys\private.key")

            AccessToken = GetDocuSignAccessToken()
            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", AccessToken)

        Catch ex As Exception
            SendProcessReport(ex.Message)
            Console.WriteLine("ProcessEnvelopes " & ex.Message)
            Console.ReadKey()
        End Try
    End Sub
    Public Shared Function AttachDocument(sID As String, mDocName As String) As Integer
        Try
            Dim sDoc As String, mNewDocName As String
            Dim sDocType As String

            Dim sDocClass As String
            Dim sDocDescr As String
            Dim iDocID As Long = 0, bFaxSaved As Boolean = False
            Dim sFaxName As String = ""
            Dim bClientFax As Boolean
            Dim bCloserViewed As Boolean, bAdvoViewed As Boolean, bSettlerViewed As Boolean

            bCloserViewed = True
            bAdvoViewed = True
            bSettlerViewed = True
            bClientFax = True
            sDoc = SavePath
            sDocType = "pdf"
            sDocDescr = mDocName

            If InStr(1, mDocName, "Contract") > 0 Then
                sDocClass = "CNT"
                sDocDescr = "Clients Signed Contract"
            ElseIf InStr(1, mDocName, "Creditor Submission Form") > 0 Then
                sDocClass = "CSF"
                sDocDescr = mDocName
            ElseIf InStr(1, mDocName, "Draft") > 0 Then
                sDocClass = "DRA"
                sDocDescr = mDocName
            Else
                sDocClass = "DOC"
                sDocDescr = mDocName
            End If

            iDocID = DocDetails_Insert(sID, bClientFax, sDocType, sDocClass, sDocDescr, 0, 0, False, 0, 0, "", bCloserViewed, bAdvoViewed, bSettlerViewed)
            'move the file to the fax folder with new name (DocID.DocType) and update DocDetails with new name
            If iDocID > 0 Then
                mNewDocName = iDocID.ToString & "." & sDocType
                bFaxSaved = CopyFax(sDoc, mNewDocName)
            End If

            'If Not bFaxSaved And iDocID > 0 Then
            '    'ExecSQL("UPDATE Docs SET DocDescription = 'Error. ' + DocDescription WHERE DocID = '" + iDocID.ToString + "'")
            'End If

            If iDocID > 0 And bFaxSaved Then
                Return iDocID
            Else
                Return 0
            End If

        Catch ex As Exception
            Console.WriteLine("AttachDocument: " & ex.Message)
            Console.ReadKey()

            Return 0
        End Try
    End Function

    Public Shared Function DocDetails_Insert(sID As String, bClient As Boolean, sDocType As String, sDocClass As String, sDocDesc As String,
                                      Optional bTestimonial As Byte = 0, Optional bLegal As Byte = 0, Optional bMarketing As Boolean = False,
                                      Optional mBDSourceID As Integer = 0, Optional DocuSign As Integer = 0, Optional WhoViewed As String = "",
                                      Optional bCloserViewed As Boolean = False, Optional bAdvoViewed As Boolean = False, Optional bSettlerViewed As Boolean = False) As Long
        Dim conn As New SqlConnection(strConnStr)
        Dim cmd As New SqlCommand("DocDetails_Insert", conn)
        Dim pReturn As New SqlParameter()
        Dim iDocID As Long = 0
        Dim sAttachedBy As String = "System - auto"

        Try
            conn.Open()
            sAttachedBy = "System - auto"
            With cmd
                .CommandType = CommandType.StoredProcedure
                pReturn.Direction = ParameterDirection.ReturnValue
                .Parameters.Add(pReturn)
                .Parameters.Add("@FromClt", SqlDbType.Bit).Value = bClient
                .Parameters.Add("@ID", SqlDbType.VarChar, 25).Value = sID
                .Parameters.Add("@DocClass", SqlDbType.VarChar, 5).Value = sDocClass
                .Parameters.Add("@DocDescription", SqlDbType.VarChar, 255).Value = sDocDesc
                .Parameters.Add("@DocType", SqlDbType.VarChar, 100).Value = sDocType
                .Parameters.Add("@Testimonial", SqlDbType.Bit).Value = bTestimonial
                .Parameters.Add("@Legal", SqlDbType.Bit).Value = bLegal
                .Parameters.Add("@Marketing", SqlDbType.Bit).Value = bMarketing   '----- as per Olga, 6/5/08 ------
                .Parameters.Add("@AttachedBy", SqlDbType.VarChar, 30).Value = sAttachedBy
                .Parameters.Add("@BDSourceID", SqlDbType.Int).Value = mBDSourceID   '------Chris - 08/21/2017
                .Parameters.Add("@WhoViewed", SqlDbType.VarChar, 10).Value = WhoViewed   '---- BK-2/16/2022 (task 3773)

                .Parameters.Add("@frmCloserViewed", SqlDbType.Bit).Value = bCloserViewed    '---- Chris - 08/09/2023 (task 4436)
                .Parameters.Add("@frmAdvoViewed", SqlDbType.Bit).Value = bAdvoViewed    '---- Chris - 08/09/2023 (task 4436)
                .Parameters.Add("@frmSettlerViewed", SqlDbType.Bit).Value = bSettlerViewed    '---- Chris - 08/09/2023 (task 4436)

                .ExecuteNonQuery()
                iDocID = pReturn.Value
                .Dispose()
            End With

            If iDocID > 0 Then
                Return iDocID
            Else
                Return 0
            End If

        Catch ex As Exception
            Console.WriteLine("DocDetails_Insert: " & ex.Message)
            Console.ReadKey()
            conn.Close()
            Return 0
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Function

    Public Shared Function CopyFax(ByVal sOrigFaxName As String, ByVal sNewFaxName As String) As Boolean
        Try
            Dim sNewFaxFullName As String

            'If working on test database, put the document in the test server - Chris - 04/23/2020
            'Because the Doc table uses auto identity to generate doc ID, it conflicts when working with test database and the document is going to the same place
            If strCatalog.ToUpper <> "CTDB" Then
                If (Not System.IO.Directory.Exists(CTDOC_Test & Now.Year.ToString)) Then
                    System.IO.Directory.CreateDirectory(CTDOC_Test & Now.Year.ToString)
                End If
                sNewFaxFullName = CTDOC_Test & Now.Year.ToString & "\" & sNewFaxName
            Else
                If (Not System.IO.Directory.Exists(CTDOC & Now.Year.ToString)) Then
                    System.IO.Directory.CreateDirectory(CTDOC & Now.Year.ToString)
                End If
                sNewFaxFullName = CTDOC & Now.Year.ToString & "\" & sNewFaxName
            End If

            If Dir(sOrigFaxName) <> "" Then
                IO.File.Copy(sOrigFaxName, sNewFaxFullName)
            End If

            If Dir(sNewFaxFullName) <> "" Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Console.WriteLine("CopyFax: " & ex.Message)
            Console.ReadKey()
            Return False
        End Try
    End Function


    Public Shared Function GetDocuSignEnvelopes() As DataTable
        Try
            Dim cnn As New SqlConnection(strConnStr)

            Dim cmd As SqlCommand = New SqlCommand("GetDocuSignEnvelopes", cnn)
            Dim sda As New SqlDataAdapter(cmd)
            Dim dsData As New DataTable, mReturn As Integer = 0
            cmd.CommandType = CommandType.StoredProcedure

            sda.Fill(dsData)

            Return dsData
        Catch ex As Exception
            Console.WriteLine("GetDocuSignEnvelopes: " & ex.Message)
            Console.ReadKey()

            Return Nothing
        End Try
    End Function


    Public Shared Function GetEnvelopeData(AccountId As String, EnvelopeId As String) As EnvelopeInfo
        Try

            Dim Url As String =
                $"https://na3.docusign.net/restapi/v2.1/accounts/{AccountId}/envelopes/{EnvelopeId}/documents/combined"

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", AccessToken)

            Dim response = client.GetAsync(Url).GetAwaiter().GetResult()
            response.EnsureSuccessStatusCode()

            Dim json As String = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()

            Console.WriteLine(json)

            Dim env As EnvelopeInfo = JsonConvert.DeserializeObject(Of EnvelopeInfo)(json)

            Return env
        Catch ex As Exception
            Console.WriteLine("GetEnvelopeData: " & ex.Message)
            Console.ReadKey()
            Return Nothing

        End Try
    End Function

    Public Shared Sub DownloadEnvelopePdf(AccountId As String, EnvelopeId As String, SavePath As String)
        Try
            'combined downloads all documents as one PDF
            Dim Url As String =
                $"https://na3.docusign.net/restapi/v2.1/accounts/{AccountId}/envelopes/{EnvelopeId}/documents/combined"

            Dim response = client.GetAsync(Url).GetAwaiter().GetResult()

            response.EnsureSuccessStatusCode()

            Dim pdfBytes() As Byte = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()

            File.WriteAllBytes(SavePath, pdfBytes)

        Catch ex As Exception
            Console.WriteLine("GetEnvelopeData: " & ex.Message)
            Console.ReadKey()
        End Try

    End Sub

    Public Shared Sub UpdateEnvelopeStatus(EnvelopeId As String, DateOpened As Nullable(Of DateTime), DateCompleted As Nullable(Of DateTime), DateDeclined As Nullable(Of DateTime), AttachedDocID As Integer)
        Try

            Dim cnn As New SqlConnection(strConnStr), iAffected As Integer = 0
            Dim Sqlqry As String = "", DateViewedStr As String = "Null", DateCompletedStr As String = "Null", DateDeclinedStr As String = "Null"
            If IsDate(DateOpened) Then DateViewedStr = "'" & DateOpened & "'"
            If IsDate(DateCompleted) Then DateCompletedStr = "'" & DateCompleted & "'"
            If IsDate(DateDeclined) Then DateDeclinedStr = "'" & DateDeclined & "'"

            Sqlqry = "Update DocuSignEnvelopes Set DateViewed=" & DateViewedStr & ",DateCompleted=" & DateCompletedStr & ",DateDeclined=" & DateDeclinedStr & ",AttachedDocID=" & AttachedDocID & " Where EnvelopeID='" & EnvelopeId & "'"

            Dim cmd As SqlCommand = New SqlCommand(Sqlqry, cnn)
            cmd.CommandType = CommandType.Text
            cnn.Open()
            iAffected = cmd.ExecuteNonQuery()
            If cnn.State = ConnectionState.Open Then
                cnn.Close()
            End If

        Catch ex As Exception
            Console.WriteLine("UpdateEnvelopeStatus: " & ex.Message)
            Console.ReadKey()
        End Try
    End Sub


    Public Shared Function GetDocuSignAccessToken() As String
        Try
            'Create API client
            Dim apiClient As New DocuSignClient()

            'Request JWT token
            Dim oauthToken =
            apiClient.RequestJWTUserToken(
                IntegrationKey,
                UserId,
                AuthServer,
                privateKeyBytes,
                1)

            Return oauthToken.access_token

        Catch ex As Exception
            Console.WriteLine("GetDocuSignAccessToken: " & ex.Message)
            Console.ReadKey()
            Return ""

        End Try

    End Function


    '============================to get all envelopes for a particular date

    Public Shared Sub GetEnvelopesByDate(SearchDate1 As DateTime, SearchDate2 As DateTime)
        Dim AccessToken As String = GetDocuSignAccessToken()

        'Get account information
        Dim authClient As New ApiClient()

        authClient.SetOAuthBasePath("account.docusign.com")

        Dim userInfo = authClient.GetUserInfo(AccessToken)

        Dim AccountId As String =
            userInfo.Accounts(0).AccountId

        Dim BaseUri As String =
            userInfo.Accounts(0).BaseUri & "/restapi"

        'Create REST client
        Dim apiClient As New ApiClient(BaseUri)

        apiClient.Configuration.DefaultHeader("Authorization") =
            "Bearer " & AccessToken

        Dim envelopesApi As New EnvelopesApi(apiClient)

        'Date you want to search
        'Dim SearchDate As Date = #6/1/2026#

        'Dim FromDate As String =
        '    SearchDate1.ToString("yyyy-MM-ddT00:00:00.0000000Z")

        Dim FromDate As String =
    SearchDate1.ToUniversalTime().
    ToString("yyyy-MM-ddTHH:mm:ssZ")

        'Dim ToDate As String =
        '    SearchDate.AddDays(1).
        '               AddSeconds(-1).
        '               ToString("yyyy-MM-ddT23:59:59.9999999Z")

        'Dim ToDate As String =
        '    SearchDate2.ToString("yyyy-MM-ddT00:00:00.0000000Z")

        Dim ToDate As String =
    SearchDate2.ToUniversalTime().
    ToString("yyyy-MM-ddTHH:mm:ssZ")

        Dim options As New EnvelopesApi.ListStatusChangesOptions()

        options.fromDate = "2025-01-01T00:00:00Z"
        'options.toDate = ToDate
        options.status = "any"

        'Dim results As EnvelopesInformation =
        '    envelopesApi.ListStatusChanges(
        '        AccountId,
        '        options)

        Dim results As EnvelopesInformation =
    envelopesApi.ListStatusChanges(
        AccountId,
        options)

        If results Is Nothing Then
            Console.WriteLine("Results is Nothing")
        Else

            Console.WriteLine(
        "ResultSetSize = " & results.ResultSetSize &
        vbCrLf &
        "TotalSetSize = " & results.TotalSetSize)

            If results.Envelopes Is Nothing Then
                Console.WriteLine("Envelopes collection is Nothing")
            Else
                Console.WriteLine("Envelope count = " &
                        results.Envelopes.Count)
            End If

        End If


        '=============================================
        If results.Envelopes Is Nothing Then Exit Sub

        For Each env As Envelope In results.Envelopes

            Debug.WriteLine("Envelope ID: " &
                            env.EnvelopeId)

            Debug.WriteLine("Subject: " &
                            env.EmailSubject)

            Debug.WriteLine("Status: " &
                            env.Status)

            Debug.WriteLine("Created: " &
                            env.CreatedDateTime)

            Debug.WriteLine("Sent: " &
                            env.StatusChangedDateTime)

            Debug.WriteLine("Completed: " &
                            env.CompletedDateTime)

            If env.Sender IsNot Nothing Then

                Debug.WriteLine("Sender: " &
                                env.Sender.UserName)

                Debug.WriteLine("Email: " &
                                env.Sender.Email)

            End If

            Debug.WriteLine("--------------------")

        Next

    End Sub




End Class
