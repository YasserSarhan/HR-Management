Access to Word Automation

This Access database demonstrates how Access data can be used as the source for a Word mail merge or for inserting text at bookmarks, a table or formfields  in a Word document.  A simple addresses database is used for the purpose of illustrating the procedures employed.

The method employed for mail merges is to first create a Word data document in the folder in which the database is running.  This is then used as the data source for the mail merge, an approach which I have generally found to be faster than using the Access data direct as the data source.  Either all records can be used for the merge or a subset of records can be selected in the list box on the form.  Options are included for merging to an unsaved new document, for merging and saving the document and for transparently merging to the printer.  Thre is also an option to include a cc list in the merege.  By default this is selected and will, if the merge is to more than one addressee, include a list of the other addressees at the bottom of each letter.  This list is a two-column, single row, Word table whose cells are filled with data from computed columns in the query used to create the data source document.  To exclude the cc list from the merge uncheck the check box on the form.

A single document based on the current record is created either by inserting text at bookmarks in a Word template document or by filling Word formfields.  A Word table can also be filled with data from Access.

To try out the application first extract the relevant AccWord??.mdb file for your version of Access, AccWordMergeLetter.doc, AccWordBookmarkedLetter.dot, AccWordTable.dot and AccWordForm.dot into the same folder.  Open the AccWord??.mdb database, add a few names and addresses and try out the options available with the buttons on the form.

The procedures called by the buttons on the form are contained in the module basWordStuff.  Feel free to use any of the code in your own applications.  Also included in the application is Bill Wilson's excellent BrowseForFileClass module, for which I am once more indebted to Bill.

Unlike the original versions of the mail merge procedures which I posted in messages in Compuserve's MSDEVAPPS and MSOFFICE forums, the procedures in this application use only one instance of Word rather than creating a new instance each time.  To see this in action try merging without saving and, without closing Word, inserting values from the current record in the bookmarked Word document.  You should find both the result of the merge and the letter to the current addressee open in the same instance of Word.  

Ken Sheridan
Stafford, England
kensheridan@dsl.pipex.com
