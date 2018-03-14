Public Class Class1
    ' Changed test
    <Test(), Description("MemberFieldChangedTestSQL"), TestCase(11614688)>
    Public Sub MemberChangeNotification_TestChanged(Optional ByVal member_Id As Integer = 0)
        ' Get any 'Changed' fields from Member_Change_Notifications            
        Dim status As String = "Changed"
        Dim changedList As New DatabaseObjects.Member_Change_Notifications(pChangeStatus:=status, pOrderBy:="Member_ID")
        ' Loop thru them and create a list of Member_Changes which summarizes the list of multiple fields into a dictionary            
        If changedList.Count > 0 Then
            Dim curMemberChanges As List(Of Member_Field_Changes) = New List(Of Member_Field_Changes)

            For idx As Integer = 0 To changedList.Count - 1
                Dim current = DirectCast(changedList.m_Table.Rows.Item(idx), TVCManaged.TVCDataset.Member_Change_NotificationsRow)
                AddToList(curMemberChanges, current.Member_ID, current.Field_Name, current.Requested_Value, status)
            Next
            ' Loop through new formatted list and send                
            If curMemberChanges.Count > 0 Then
                For Each mchange As Member_Field_Changes In curMemberChanges
                    Dim tm = CreateTemplatedMessageSender()(mchange, status, "Fields Changed")
                    '   Dim tm As TemplatedMessageSender = New TemplatedMessageSender("Account Notifications", "Fields Changed")
                    '   m_TestMember1 = DatabaseObjects.Member.Factory(mchange.Member_ID)
                    '   tm.Recipient = New TemplatedMessageRecipient() With {.MemberID = mchange.Member_ID}
                    '   tm.Placeholder = TemplatePlaceholders.GetConfiguredPlaceholders("AccountNotifications.FieldsChanged", dboMember:=m_TestMember1)
                    '   tm.Placeholder.Item("FieldChanges").ReplacementValue = mchange.ToString()
                    '   tm.SenderEmailAddress = "newmcaassociate@tvcmarketing.com"
                    '   tm.SenderFullName = "Motor Club of America"
                    '   tm.OutgoingMessagesDisabled = False
                Next
            End If
        End If
    End Sub

    ' Pending test
    <Test(), Description("MemberFieldPendningTestSQL"),
     TestCase()>
    Public Sub MemberChangeNotification_TestPending()
        ' Get any 'Pending' fields from Member_Change_Notifications            
        Dim status As String = "Pending"
        Dim changedList As New DatabaseObjects.Member_Change_Notifications(pChangeStatus:=status, pOrderBy:="Member_ID")
        ' Loop thru them and create a list of Member_Changes which summarizes the list of multiple fields into a dictionary            
        If changedList.Count > 0 Then
            Dim curMemberChanges As List(Of Member_Field_Changes) = New List(Of Member_Field_Changes)

            For idx As Integer = 0 To changedList.Count - 1
                Dim current = DirectCast(changedList.m_Table.Rows.Item(idx), TVCManaged.TVCDataset.Member_Change_NotificationsRow)
                AddToList(curMemberChanges, current.Member_ID, current.Field_Name, current.Requested_Value, status)
            Next
            ' Loop through new formatted list and send                
            If curMemberChanges.Count > 0 Then
                For Each mchange As Member_Field_Changes In curMemberChanges
                    Dim tm = CreateTemplatedMessageSender()(mchange, status, "Fields Pending")
                Next
            End If
        End If
    End Sub

    ' Approved Test
    <Test(), Description("MemberFieldsApprovedTestSQL"),
     TestCase()>
    Public Sub MemberChangeNotification_TestApproved()
        ' Get any 'Approved' fields from Member_Change_Notifications            
        Dim status As String = "Approved"
        Dim changedList As New DatabaseObjects.Member_Change_Notifications(pChangeStatus:=status, pOrderBy:="Member_ID")
        ' Loop thru them and create a list of Member_Changes which summarizes the list of multiple fields into a dictionary            
        If changedList.Count > 0 Then
            Dim curMemberChanges As List(Of Member_Field_Changes) = New List(Of Member_Field_Changes)

            For idx As Integer = 0 To changedList.Count - 1
                Dim current = DirectCast(changedList.m_Table.Rows.Item(idx), TVCManaged.TVCDataset.Member_Change_NotificationsRow)
                AddToList(curMemberChanges, current.Member_ID, current.Field_Name, current.Requested_Value, status)
            Next
            ' Loop through new formatted list and send                
            If curMemberChanges.Count > 0 Then
                For Each mchange As Member_Field_Changes In curMemberChanges
                    Dim tm = CreateTemplatedMessageSender()(mchange, status, "Fields Approved")
                Next
            End If
        End If
    End Sub

    ' Denied test
    <Test(), Description("MemberFieldsDeniedTestSQL"),
     TestCase()>
    Public Sub MemberChangeNotification_TestDenied()
        ' Get any 'Denied' fields from Member_Change_Notifications            
        Dim status As String = "Denied"
        Dim changedList As New DatabaseObjects.Member_Change_Notifications(pChangeStatus:=status, pOrderBy:="Member_ID")
        ' Loop thru them and create a list of Member_Changes which summarizes the list of multiple fields into a dictionary            
        If changedList.Count > 0 Then
            Dim curMemberChanges As List(Of Member_Field_Changes) = New List(Of Member_Field_Changes)

            For idx As Integer = 0 To changedList.Count - 1
                Dim current = DirectCast(changedList.m_Table.Rows.Item(idx), TVCManaged.TVCDataset.Member_Change_NotificationsRow)
                AddToList(curMemberChanges, current.Member_ID, current.Field_Name, current.Requested_Value, status)
            Next
            ' Loop through new formatted list and send                
            If curMemberChanges.Count > 0 Then
                For Each mchange As Member_Field_Changes In curMemberChanges
                    Dim tm = CreateTemplatedMessageSender()(mchange, status, "Fields Denied")
                Next
            End If
        End If
    End Sub

    Private Function CreateTemplatedMessageSender(mchange As Member_Field_Changes, status As String, statusString As String) As TemplatedMessageSender
        Dim tm As TemplatedMessageSender = New TemplatedMessageSender("Account Notifications", statusString)
        Dim m_TestMember1 = DatabaseObjects.Member.Factory(mchange.Member_ID)
        tm.Recipient = New TemplatedMessageRecipient() With {.MemberID = mchange.Member_ID}
        tm.Placeholder = TemplatePlaceholders.GetConfiguredPlaceholders("AccountNotifications." + status, dboMember:=m_TestMember1)
        tm.Placeholder.Item(status).ReplacementValue = mchange.ToString()
        tm.SenderEmailAddress = "newmcaassociate@tvcmarketing.com"
        tm.SenderFullName = "Motor Club of America"
        tm.OutgoingMessagesDisabled = False
        Return tm
    End Function

    ' 2108 - Mik SQL methods to be called by Business Layer         
    Public Function GetChangedFieldNamesSQL(ByVal specificMemberID As Integer) As String
        Dim sbSQL As New StringBuilder()
        With sbSQL
            .AppendLine("SELECT Member_ID, Field_Name, Requested_Value")
            .AppendLine("FROM Member_Change_Notifications mcn  ")
            .AppendLine("WHERE (mcn.Email_Status = 0 And mcn.Change_Status ='Changed')")
            If specificMemberID > 0 Then
                .AppendLine("AND mcn.Member_ID = " & specificMemberID.ToString())
            End If
            .AppendLine("AND NOT EXISTS(SELECT pfc.Member_ID , pfc.Field_Name")
            .AppendLine("FROM Protected_Field_Changes pfc   ")
            .AppendLine("Where pfc.Member_ID = mcn.Member_ID And pfc.Field_Name = mcn.Field_Name)")
            .AppendLine("Order By Member_ID, Field_Name")
        End With
        Return sbSQL.ToString()
    End Function


    Public Function GetPendingFieldNamesSQL() As String
        Dim sbSQL As New StringBuilder()
        With sbSQL
            .AppendLine("SELECT mcn.Member_ID, mcn.Field_Name, mcn.Requested_Value")
            .AppendLine("FROM Member_Change_Notifications mcn  ")
            .AppendLine("INNER JOIN Protected_Field_Changes pfc On pfc.Member_ID = mcn.Member_ID")
            .AppendLine("WHERE (mcn.Email_Status = 0 ")
            .AppendLine("And mcn.Change_Status ='Pending' ")
            .AppendLine("And pfc.Status ='Pending' ")
            .AppendLine("And pfc.Field_Name = mcn.Field_Name ")
            .AppendLine("And pfc.Response_Date IS NULL) ")
            .AppendLine("Order By mcn.Member_ID, mcn.Field_Name")
        End With
        Return sbSQL.ToString()
    End Function

    Public Function GetApprovedFieldNamesSQL() As String
        Dim sbSQL As New StringBuilder()
        With sbSQL
            .AppendLine("SELECT mcn.Member_ID, mcn.Field_Name, mcn.Requested_Value")
            .AppendLine("FROM Member_Change_Notifications mcn  ")
            .AppendLine("INNER JOIN Protected_Field_Changes pfc On pfc.Member_ID = mcn.Member_ID")
            .AppendLine("WHERE (pfc.Field_Name = mcn.Field_Name ")
            .AppendLine("And mcn.Change_Status ='Approved' ")
            .AppendLine("And pfc.Status ='Approved' ")
            .AppendLine("And mcn.Date_Finalized IS NOT NULL) ")
            .AppendLine("Order By mcn.Member_ID, mcn.Field_Name")
        End With
        Return sbSQL.ToString()
    End Function

    Public Function GetDeniedFieldNamesSQL() As String
        Dim sbSQL As New StringBuilder()
        With sbSQL
            .AppendLine("SELECT mcn.Member_ID, mcn.Field_Name, mcn.Requested_Value")
            .AppendLine("FROM Member_Change_Notifications mcn  ")
            .AppendLine("INNER JOIN Protected_Field_Changes pfc On pfc.Member_ID = mcn.Member_ID")
            .AppendLine("WHERE (pfc.Field_Name = mcn.Field_Name ")
            .AppendLine("And mcn.Change_Status ='Denied' ")
            .AppendLine("And pfc.Status ='Denied' ")
            .AppendLine("And mcn.Date_Finalized IS NOT NULL) ")
            .AppendLine("Order By mcn.Member_ID, mcn.Field_Name")
        End With
        Return sbSQL.ToString()
    End Function

    Public Function SendChangeFieldReminders(Optional ByVal intSpecificMemberID As Integer = 0) As Integer
        Dim dt As DataTable = GetData(GetChangedFieldNamesSQL(intSpecificMemberID))
        ' The rows are now unsorted by Member_ID with various Field_Names and Values            
        For Each row As DataRow In dt.Rows
            RefreshLock()
            DebugWriteLine("Notification: Changed field row:" & row("Member_ID") & row("Field_Name") & row("Requested_Value"))
            'Dim lc As New DatabaseObjects.Legal_Case(CInt(row("ID")))          
            '       Try                 
            '           Dim tm As New TVCManaged.TemplatedMessageSender("Legal Notifications", "Court Date Reminder")                 
            '           tm.Placeholder = TemplatePlaceholders.GetConfiguredPlaceholders("LegalCase.CrtDateReminder", dboLegalCase:=lc)                
            '           tm.Placeholder.Item("CourtDate").ReplacementValue = row("Start_Date").ToString().Replace(" 12:00:00 AM", "")
            ' Drop unspecified time                 '           tm.Placeholder.Item("CourtDateType").ReplacementValue = row("Name")     
            '           tm.Recipient.Member = lc.Member      
            '           tm.Recipient.MessageIndex.LegalCaseID = lc.ID     
            '           tm.SenderEmailAddress = "legal@tvcmarketing.com"           
            '           tm.SenderFullName = "TVC Legal Department"        
            '           tm.Send()           
            '          
            'SendCourtDateReminders += 1     
            '       Catch ex As Exception             
            '           TVCManaged.AppException.Handle(ex)        
            '       End Try     
        Next
    End Function
End Class
