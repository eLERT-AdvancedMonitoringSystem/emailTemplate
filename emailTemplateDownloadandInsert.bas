Attribute VB_Name = "Module1"
Sub insertHtml()

'Inspector Represents the window in which an Outlook item is displayed'
Dim insp As Inspector
'set is used to assign a reference to an object'
Set insp = ActiveInspector

If insp.IsWordMail Then
    'create a word document varible'
    Dim wordDoc As Word.Document
    'set is used to assign a reference to an object'
    Set wordDoc = insp.WordEditor
    
    'download the file from github'
    wordDoc.Application.Selection.InsertFile "https://raw.githubusercontent.com/jjenksy/emailTemplate/master/emailHtmlTemplate.html", , False, False, False
    
End If


End Sub
