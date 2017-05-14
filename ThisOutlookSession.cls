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
    Item.Subject = "[Unassigned]" & Item.Subject & " @" & Item.Sender & " " & Now
 Item.Save

Set myForward = Item.Forward
myForward.BCC = "wrike@wrike.com"

myForward.Send

End Sub
