Public WithEvents myOlItems As Outlook.Items


Public Sub Application_Startup()

   ' Reference the items in the Inbox. Because myOlItems is declared
   ' "WithEvents" the ItemAdd event will fire below.
   Set myOlItems = Outlook.Session.GetDefaultFolder(olFolderInbox).Items
       

End Sub


Private Sub myOlItems_ItemAdd(ByVal Item As Object)

   ' If it's currently not between 9:00 A.M. and 5:00 P.M.
   If Time() > #9:00:00 AM# Or Time() < #5:00:00 PM# Then

      ' Check to make sure it is an Outlook mail message, otherwise
      ' subsequent code will probably fail depending on what type
      ' of item it is.
      If TypeName(Item) = "MailItem" Then

         ' Forward the item just received
         Set myForward = Item.Forward

         ' Address the message
         myForward.Recipients.Add "joeday@uchicago.edu"

         ' Send it
         myForward.Send

      End If

   End If
   
   If DatePart("m", Now) Mod 3 = 1 Then
       Call RemoveEmail90
    End If

End Sub

Public Sub RemoveEmail90()

Dim olSession As Outlook.Application, olNamespace As NameSpace
Dim olInbox As Outlook.MAPIFolder
Dim i As Integer
Set olSession = New Outlook.Application
Set olNamespace = olSession.GetNamespace("MAPI")
Set olInbox = olNamespace.GetDefaultFolder(olFolderInbox)
Set Delete_Items = olInbox.Items

For i = Delete_Items.Count To 1 Step -1
    If TypeName(Delete_Items.Item(i)) = "MailItem" Then
            If DateDiff("d", Now, Delete_Items.Item(i).ReceivedTime) > 30 Then Delete_Items.Item(i).Delete
    End If
Next

Set olSession = Nothing
Set olNamespace = Nothing
Set olInbox = Nothing
End Sub


Sub MyProc()
    If DatePart("m", Now) Mod 3 = 1 Then
       Call RemoveEmail90
    End If
    
    Application.OnTime TimeSerial(16, 30, 0), "MyProc"
End Sub
