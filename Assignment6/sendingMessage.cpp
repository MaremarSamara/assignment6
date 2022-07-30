Sub SendMessageMQF( _  
                   strComputerName1 As String, _  
                   strQueueName1 As String, _  
                   strComputerName2 As String, _  
                   strQueueName2 As String, _  
                   strComputerName3 As String, _  
                   strQueueName3 As String _  
                   )  
  
  Dim strFormatName1 As String  
  Dim strFormatName2 As String  
  Dim strFormatName3 As String  
  Dim strMultipleElement As String  
  Dim dest As New MSMQDestination  
  Dim msg As New MSMQMessage  
  
  ' Create multiple-element format name.  
  strFormatName1 = "DIRECT=OS:" & strComputerName1 & "\" & strQueueName1  
  strFormatName2 = "DIRECT=OS:" & strComputerName2 & "\" & strQueueName2  
  strFormatName3 = "DIRECT=OS:" & strComputerName3 & "\" & strQueueName3  
  strMultipleElement = strFormatName1 & "," & strFormatName2 & "," & strFormatName3  
  
  'Set format name of MSMQDestination object.  
  On Error GoTo ErrorHandler  
  dest.FormatName = strMultipleElement  
  MsgBox "Format name is: " + dest.FormatName  
  
  ' Set the message label.  
  msg.Label = "Test Message"  
  
  ' Send the message and close the MSMQDestination object.  
  msg.Send DestinationQueue:=dest  
  dest.Close  
  
  Exit Sub  
  
ErrorHandler:  
  MsgBox "Error " + Hex(Err.Number) + " was returned." _  
         + Chr(13) + Err.Description  
End Sub