using System.Collections.Generic;

namespace VersaNote.Interop
{
    public enum NotificationType
    {
        None,
        Information,
        Success,
        Warning,
        Error,
        Notification
    }

    public enum NoteTypes_e
    {
        Blank,
        Custom,
        GTol,
        NewColumn,
        Standard,
        Unknown
    }

    public struct DeleteCustomNoteResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public struct AddCustomNoteResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string NoteId { get; set; }
    }

    public struct GetSheetNotesResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public Note[] Notes { get; set; }
    }

    public struct Note
    {
        public NoteTypes_e NoteType { get; set; }
        public string Id { get; set; }
        public string Text { get; set; }
        public bool Indented { get; set; }
    }

    public struct GetLinkedAnnotationsResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public LinkedAnnotation[] LinkedAnnotations { get; set; }
    }

    public struct LinkedAnnotation
    {
        public string NoteId { get; set; }
        public string Text { get; set; }
    }

    public struct RefreshNoteTablesResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public struct UpdatedLinkedAnnotationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public struct UpdatedCustomNoteResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }
}
