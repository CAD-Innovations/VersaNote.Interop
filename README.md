# Versa Note Interop
Enables programmatic interaction with the Versa Note add-in for SolidWorks. This interop is supported by Versa Note v1.28.0 and later.

First, get the Versa Note add-in object in SolidWorks using the Versa Note GUID:
```
public VersaNote.Interop.Addin VersaNoteAddin;

object vnObj = iSwApp.GetAddInObject("{BB75C0FD-3720-45E5-8B80-2ACB34900F7E}");
if (vnObj != null)
    VersaNoteAddin = new VersaNote.Interop.Addin(vnObj);
```
Before adding or modifying any notes, get a list of existing drawing notes on the specified sheet:
```
int sheetIndex = 1;
var sheetNotes = VersaNoteAddin.GetAllSheetNotes(sheetIndex);
```
To add a custom note at the end of the notes:
```
var addNoteResult = VersaNoteAddin.AddCustomNote(sheetIndex, "Hello world!");
```

Similarly, you can delete or update custom notes using `DeleteCustomNote` and `UpdateCustomNote` respectively.

By default, the notes displayed on the drawing are refreshed with each API call. If you need to make several API calls, you can suppress the refresh and explicitly refresh the notes when you are finished:
```
var updateNoteResult = VersaNoteAddin.UpdateCustomNote(sheetIndex, sheetNotes.Notes[0].Id, "Updated note text", suppressTableRefresh:true);
var deleteNoteResult = VersaNoteAddin.DeleteCustomNote(sheetIndex, sheetNotes.Notes[1].Id, suppressTableRefresh: true);

var refreshResult = VersaNoteAddin.RefreshNoteTables(sheetIndex);
```

In certain cases, notes with associated linked annotations may need to be linked to a different note (for example, if indented child notes are added to a parent note, e.g. note 7. becomes 7.1)

To get the list of linked annotations on the current sheet:
```
var linkedAnnotations = VersaNoteAddin.GetLinkedAnnotationsOnSheet(sheetIndex);
```

You can iterate through the list of linked annotations to determine if the current sheet contains annotations associated with a given note, and then re-assign those annotations to a new note:
```
var matchingItems = linkedAnnotations.ToList().FirstOrDefault(x => x.NoteId == parentNote.Id);

if (matchingItems.Count() > 0)
    VersaNoteAddin.UpdatedLinkedAnnotationsOnActiveSheet(parentNote.Id, childNote.Id);
```
The note number associated with the linked annotation will be updated automatically.

You can also open the Versa Note Editor from another application:
```
VersaNoteAddin.OpenEditor();
```

Or, use Versa Note to display toast notifications, which can be helpful to avoid overlapping notifications from multiple add-ins:
```
VersaNoteAddin.ShowToastNotification("Hello From Another Application", "We hope you are enjoying Versa Note!", NotificationType.Information);
```