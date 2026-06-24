Imports Microsoft.Office.Interop

Module Module1

    Sub Main()
        Console.WriteLine("Starting process...!")
        clsDocuSign.ProcessEnvelopes()
    End Sub


    Public Sub SendProcessReport(mMsg As String, Optional mMsg2 As String = "")
        Try
            Dim objOutlook As Outlook.Application, mSubject As String, mBody As String
            objOutlook = New Outlook.Application
            objOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            Dim NewMsg As Outlook.MailItem = objOutlook.CreateItem(Outlook.OlItemType.olMailItem)

            mSubject = "DocuSign Tracker"

            mBody = mMsg & vbCrLf & vbCrLf & mMsg2

            With NewMsg
                .To = "cmamah@corpturn.com"
                .Subject = mSubject
                .Body = mBody
                .Send()
            End With

        Catch ex As Exception
            'MessageBox.Show(ex.Message, "Sending accounting process Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'Log_Error(LOG, "SendAccountingProcessReport - " & ex.Message)

        End Try
    End Sub

End Module
