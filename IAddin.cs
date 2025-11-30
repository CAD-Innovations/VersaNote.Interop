namespace VersaNote.Interop
{
    public interface IAddin
    {
        void ShowToastNotification(string title, string message, NotificationType notificationType, int delaySeconds);
        AddCustomNoteResult AddCustomNote(int sheetIndex, string noteText, int noteIndex = -1, bool subNote = false, string parentId = "", bool suppressTableRefresh = false);
        DeleteCustomNoteResult DeleteCustomNote(int sheetIndex, string noteId, bool suppressTableRefresh = false);
        GetSheetNotesResult GetAllSheetNotes(int sheetIndex);
        void OpenEditor();
        RefreshNoteTablesResult RefreshNoteTables(int sheetIndex);
        UpdatedCustomNoteResult UpdateCustomNote(int sheetIndex, string noteId, string newNoteText, bool toggleSubNote = false, bool suppressTableRefresh = false);
    }
}