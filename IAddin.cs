namespace VersaNote.Interop
{
    public interface IAddin
    {
        void ShowToastNotification(string title, string message, NotificationType notificationType, int delaySeconds);
        AddCustomNoteResult AddCustomNote(int sheetIndex, string noteText, int noteIndex = -1, bool subNote = false, string parentId = "", bool toogleLinkNote = false, bool suppressTableRefresh = false, bool resetTablePosition = false);
        DeleteCustomNoteResult DeleteCustomNote(int sheetIndex, string noteId, bool suppressTableRefresh = false, bool resetTablePosition = false);
        GetLinkedAnnotationsResult GetLinkedAnnotationsOnSheet(int sheetIndex);
        GetSheetNotesResult GetAllSheetNotes(int sheetIndex);
        void OpenEditor();
        RefreshNoteTablesResult RefreshNoteTables(int sheetIndex, bool resetTablePosition = false);
        UpdatedLinkedAnnotationResult UpdatedLinkedAnnotationsOnActiveSheet(string prevNoteId, string newNoteId);
        UpdatedCustomNoteResult UpdateCustomNote(int sheetIndex, string noteId, string newNoteText, bool toggleSubNote = false, bool toogleLinkNote = false, bool suppressTableRefresh = false, bool resetTablePosition = false);
    }
}