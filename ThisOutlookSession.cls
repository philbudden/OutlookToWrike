VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ChangeSubjectForwardVRSParts(Item As Outlook.MailItem)

If Item.Subject Like "RE:*" Then
  Set MsgReply = Item.Reply
  With MsgReply
      .Subject = "Error Logging Task"
      .HTMLBody = "It looks like you're replying to a task logging distribution group." & vbCrLf & "If it was your intention to log a task with the Analytics team please resend your e-mail after removing 'RE:' from the subject box, otherwise please reply directly to the intended recipient." & vbCrLf & "Please note, no task has been logged with the Analytics team at this time"
      .Body = "It looks like you're replying to a task logging distribution group." & vbCrLf & "If it was your intention to log a task with the Analytics team please resend your e-mail after removing 'RE:' from the subject box, otherwise please reply directly to the intended recipient." & vbCrLf & "Please note, no task has been logged with the Analytics team at this time"
      .Send
  End With

Else: Item.Subject = "[Unassigned - VRS Parts]" & Item.Subject & " @" & Item.Sender & " " & Now
Item.Save

Set myForward = Item.Forward
myForward.BCC = "wrike@wrike.com"

myForward.Send

End If

End Sub

Sub ChangeSubjectForwardDataAdmin(Item As Outlook.MailItem)

If Item.Subject Like "RE:*" Then
  Set MsgReply = Item.Reply
  With MsgReply
      .Subject = "Error Logging Task"
      .HTMLBody = "It looks like you're replying to a task logging distribution group." & vbCrLf & "If it was your intention to log a task with the Analytics team please resend your e-mail after removing 'RE:' from the subject box, otherwise please reply directly to the intended recipient." & vbCrLf & "Please note, no task has been logged with the Analytics team at this time"
      .Body = "It looks like you're replying to a task logging distribution group." & vbCrLf & "If it was your intention to log a task with the Analytics team please resend your e-mail after removing 'RE:' from the subject box, otherwise please reply directly to the intended recipient." & vbCrLf & "Please note, no task has been logged with the Analytics team at this time"
      .Send
  End With

Else: Item.Subject = "[Unassigned - Data Admin]" & Item.Subject & " @" & Item.Sender & " " & Now
Item.Save

Set myForward = Item.Forward
myForward.BCC = "wrike@wrike.com"

myForward.Send

End If

End Sub
