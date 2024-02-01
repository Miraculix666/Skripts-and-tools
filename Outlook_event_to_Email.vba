Public WithEvents ReceivedItems As Outlook.Items

Private Sub Application_Startup()
    Set ReceivedItems = Outlook.Application.Session.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub ReceivedItems_ItemAdd(ByVal Item As Object)
    Dim xMeetingItem As MeetingItem
    Dim xMeetingResponse As MeetingItem
    Dim xForwardMeeting As MeetingItem
    Dim xAppointmentItem As AppointmentItem
    On Error Resume Next
    If TypeOf Item Is MeetingItem Then
        Set xMeetingItem = Item
        Set xAppointmentItem = xMeetingItem.GetAssociatedAppointment(True)
        Set xMeetingResponse = xAppointmentItem.Respond(olMeetingAccepted, True)
        xMeetingResponse.Send
        Set xForwardMeeting = xMeetingItem.Forward
        With xForwardMeeting
            With .Recipients
                .Add "Marius.Albring@gmail.com" ' Ändern Sie die Adresse in Ihre eigene
                .ResolveAll
            End With
            .Send
        End With
    End If
End Sub
