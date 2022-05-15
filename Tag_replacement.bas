Attribute VB_Name = "Tag_replacement"
Option Explicit
Public wApp As Word.Application
Public wDoc As Word.Document

'remember to enable word object library:
' Tools > References > Microsoft Word Object Library

Public Enum mode
    Replacement
    delete_section
End Enum



'this sub opens the word document
Sub OpenDocument()
    Set wApp = CreateObject("Word.Application")
    wApp.Visible = True
    Dim temppath As String


        
    'put file path below
    temppath = "DOCUMENT PATH"
    Set wDoc = wApp.Documents.Open(temppath)
    
    
    'call the tagreplacement() function to replace/delete any parts of text

End Sub



' tag argument should be the tag placed in word document:
' tags should follow this syntax:
' {{TAG_NAME}} if you will be using the tag as a placeholder for an inputted value or;
' {{TAG_NAME}} ... {{/TAG_NAME}} if you want to name a section to delete if a condition is met.
' 'mode' argument has two options: replace or delete_section
' replacement will replace ONLY the tag with a desired value specified with argument 'replacement_text'
' delete_section will delete a section demarcated with {{TAG_NAME}} ... {{/TAG_NAME}} in the word doc

Sub tagreplacement(tag As Variant, mode, _
    Optional ByVal replacement_text As Variant = "")
    
'the next four lines define the closing tag name


    With wDoc
    
        'finds the tag in the text
        With .Application.Selection.Find
        .Text = tag
        .Wrap = wdFindContinue
        .Replacement.Text = replacement_text
        .Forward = True
        .Execute
        End With
        
        Do While .Application.Selection.Find.Found = True
        
            Select Case mode
            
            'replaces tag with replacement_text argument'
            Case Replacement
            .Application.Selection.Find.Execute Replace:=wdReplaceAll, _
                Forward:=True, Wrap:=wdFindContinue
            


            'deletes a defined section
            Case delete_section
            
            'sets a variable to the closing_tag... i.e adds a "/" after first two chars
            Dim tag_no_open_brackets As String
            Dim closing_tag As String
            tag_no_open_brackets = Right(tag, Len(tag) - 2)
            closing_tag = "{{" + "/" + tag_no_open_brackets
                      
redo:

            'extends selection to the next pair of closed curly brackets
            .Application.Selection.Extend Character:="}"
            .Application.Selection.Extend Character:="}"
            
            'checks that the tag which the selection has been extended to corresponds, if it doesnt, machine goes back up a few lines
            If Right(.Application.Selection.Text, Len(closing_tag)) <> closing_tag _
                Then GoTo redo
        
            .Application.Selection.Delete
            .Application.Selection.Collapse wdCollapseEnd
            .Application.Selection.Find.Execute
            End Select
            
        
        Loop
        
        .Application.Selection.EndOf
        
    End With
End Sub








Sub MAIN()

Call OpenDocument

'call tagreplacement module below!'

End Sub







