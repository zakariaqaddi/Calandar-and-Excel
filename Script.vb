Sub CreateAppointment()

    Set olOutlook = CreateObject("Outlook.Application")
    Set Namespace = olOutlook.GetNameSpace("MAPI")
    Set oloFolder = Namespace.GetDefaultFolder(9)
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        Description = Cells(i, 6).Value
        StartDate = Cells(i, 10).Value
        
        Set Appointment = oloFolder.items.Add
         
        With Appointment
            .Start = StartDate
            .Subject = Description
            .Save
            
        End With
        
    Next i
    
End Sub
