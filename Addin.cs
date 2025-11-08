using System.Linq;
using System.Reflection;

namespace VersaNote.Interop
{
    public class Addin : IAddin
    {
        private object VersaNoteObject;
        private MethodInfo[] methods;

        public Addin(object versaNoteObject)
        {
            VersaNoteObject = versaNoteObject;
            methods = VersaNoteObject.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance);
        }

        /// <summary>
        /// Add a new custom note to the specified drawing sheet.
        /// <para>
        /// Prior to running this method, you should run GetAllSheetNotes() to ensure the insertion position is valid.
        /// If creating a sub note, use GetAllSheetNotes() to find the ID of the parent note.
        /// </para>
        /// </summary>
        /// <param name="sheetIndex">A 1 based sheet number index</param>
        /// <param name="noteText">Note text</param>
        /// <param name="noteIndex">(Optional) A 0 based note index. Specify -1 to add to the end of the notes.</param>
        /// <param name="subNote">(Optional) Set true to create an indented sub-note</param>
        /// <param name="parentId">(Optional) If creating an indented sub-note, you must specify the parent note ID, and the specified note position will be interpreted relative to the parent note (i.e. if the parent note is 3. a noteIndex of 0 would create note 3.1)</param>
        /// <param name="suppressTableRefresh">(Optional) Set true to suppress note table refresh for improved performance when making multiple API calls</param>
        /// <returns>
        /// A <see cref="AddCustomNoteResult"/> structure containing the result of the operation:
        /// <list type="bullet">
        /// <item>
        /// <term><see cref="AddCustomNoteResult.Success"/></term>
        /// <description><c>true</c> if the note was created successfully; otherwise, <c>false</c>.</description>
        /// </item>
        /// <item>
        /// <term><see cref="AddCustomNoteResult.Message"/></term>
        /// <description>
        ///     A descriptive message providing details about the result. May contain
        ///     error information if <see cref="AddCustomNoteResult.Success"/> is <c>false</c>.
        /// </description>
        /// </item>
        /// <item>
        /// <term><see cref="AddCustomNoteResult.NoteId"/></term>
        /// <description>
        ///     The unique identifier of the new custom note, if successful; otherwise, <c>null</c>.
        /// </description>
        /// </item>
        /// </list>
        /// </returns>
        /// <example>
        /// The following example shows how to call this method and inspect the result:
        /// <code language="csharp">
        /// var result = VersaNote.Interop.AddCustomNote(1, "This is the custom note text", 3);
        /// if (result.Success)
        /// {
        ///     Console.WriteLine($"Note created with ID: {result.NoteId}");
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Failed to create note: {result.Message}");
        /// }
        /// </code>
        /// </example>
        public AddCustomNoteResult AddCustomNote(int sheetIndex, string noteText, int noteIndex = -1, bool subNote = false, string parentId = "", bool suppressTableRefresh = false)
        {
            if (subNote == true && parentId == "")
                return new AddCustomNoteResult()
                {
                    Success = false,
                    Message = "A parent ID must be specified if subNote is set to True",
                    NoteId = null
                };

            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(AddCustomNote));
            if (method != null)
            {
                object[] parameters = new object[] { sheetIndex, noteText, noteIndex, subNote, parentId, suppressTableRefresh };
                AddCustomNoteResult AddCustomNoteResult = (AddCustomNoteResult)method.Invoke(VersaNoteObject, parameters);
                return AddCustomNoteResult;
            }

            return new AddCustomNoteResult()
            {
                Success = false,
                Message = $"{MethodInfo.GetCurrentMethod().Name} method not found in Versa Note"
            };
        }

        /// <summary>
        /// Delete the specified custom note from the specified drawing sheet.
        /// </summary>
        /// <param name="sheetIndex">A 1 based sheet number index</param>
        /// <param name="noteId">The unique Id of the note to be deleted</param>
        /// <param name="suppressTableRefresh">(Optional) Set true to suppress note table refresh for improved performance when making multiple API calls</param>
        /// <returns>
        /// A <see cref="DeleteCustomNoteResult"/> structure containing the result of the operation:
        /// <list type="bullet">
        /// <item>
        /// <term><see cref="DeleteCustomNoteResult.Success"/></term>
        /// <description><c>true</c> if the note was created successfully; otherwise, <c>false</c>.</description>
        /// </item>
        /// <item>
        /// <term><see cref="DeleteCustomNoteResult.Message"/></term>
        /// <description>
        ///     A descriptive message providing details about the result. May contain
        ///     error information if <see cref="DeleteCustomNoteResult.Success"/> is <c>false</c>.
        /// </description>
        /// </item>
        /// </list>
        /// </returns>
        /// <example>
        /// The following example shows how to call this method and inspect the result:
        /// <code language="csharp">
        /// var result = VersaNote.Interop.AddCustomNote(1, "This is the custom note text", 3);
        /// if (result.Success)
        /// {
        ///     Console.WriteLine($"Note created with ID: {result.NoteId}");
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Failed to create note: {result.Message}");
        /// }
        /// </code>
        /// </example>
        public DeleteCustomNoteResult DeleteCustomNote(int sheetIndex, string noteId, bool suppressTableRefresh = false)
        {
            // if the note is not custom, return error

            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(DeleteCustomNote));
            if (method != null)
            {
                object[] parameters = new object[] { sheetIndex, noteId, suppressTableRefresh };
                DeleteCustomNoteResult deleteCustomNoteResult = (DeleteCustomNoteResult)method.Invoke(VersaNoteObject, parameters);
                return deleteCustomNoteResult;
            }

            return new DeleteCustomNoteResult()
            {
                Success = false,
                Message = $"{MethodInfo.GetCurrentMethod().Name} method not found in Versa Note"
            };
        }

        /// <summary>
        /// Get the list of all notes the specified drawing sheet.
        /// </summary>
        /// <param name="sheetIndex">A 1 based sheet number index</param>
        /// <returns>
        /// A <see cref="GetSheetNotesResult"/> structure containing the result of the operation.
        /// <para>
        /// The <see cref="GetSheetNotesResult"/> structure includes the following members:
        /// </para>
        /// <list type="bullet">
        ///   <item>
        ///     <term><see cref="GetSheetNotesResult.Success"/></term>
        ///     <description>
        ///     <c>true</c> if the operation completed successfully; otherwise, <c>false</c>.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <term><see cref="GetSheetNotesResult.Message"/></term>
        ///     <description>
        ///     A descriptive message providing details about the result. May contain
        ///     error information if <see cref="GetSheetNotesResult.Success"/> is <c>false</c>.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <term><see cref="GetSheetNotesResult.Notes"/></term>
        ///     <description>
        ///     An array of <see cref="Note"/> objects representing the notes associated
        ///     with the sheet. This array may be empty or <c>null</c> if no notes exist.
        ///     </description>
        ///   </item>
        /// </list>
        /// <para>
        /// Each <see cref="Note"/> object includes:
        /// </para>
        /// <list type="bullet">
        ///   <item>
        ///     <term><see cref="Note.NoteType"/></term>
        ///     <description>
        ///         The type of note, as defined by the <see cref="NoteTypes_e"/> enumeration.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <term><see cref="Note.Id"/></term>
        ///     <description>
        ///         The unique identifier string of the note.
        ///     </description>
        ///   </item>
        ///   <item>
        ///     <term><see cref="Note.Text"/></term>
        ///     <description>
        ///         The value of the note text as it appears on the drawing.
        ///     </description>
        ///    </item>
        /// </list>
        /// </returns>
        public GetSheetNotesResult GetAllSheetNotes(int sheetIndex)
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(GetAllSheetNotes));
            if (method != null)
            {
                object[] parameters = new object[] { sheetIndex };
                GetSheetNotesResult sheetNotesResult = (GetSheetNotesResult)method.Invoke(VersaNoteObject, parameters);
                return sheetNotesResult;
            }

            return new GetSheetNotesResult()
            {
                Success = false,
                Message = $"{MethodInfo.GetCurrentMethod().Name} method not found in Versa Note"
            };
        }

        public void OpenEditor()
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(OpenEditor));
            if (method != null)
            {
                method.Invoke(VersaNoteObject, null);
            }
        }

        public void ShowToastNotification(string title, string message, NotificationType notificationType)
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(ShowToastNotification));
            if (method != null)
            {
                int intNotificationType = (int)notificationType;
                object[] parameters = new object[] { title, message, intNotificationType };
                method.Invoke(VersaNoteObject, parameters);
            }
        }

        /// <summary>
        /// Call this method after note modifications are complete to update the note tables on the specified drawing sheet.
        /// <para>
        /// If this method is not called, the note tables will automatically be updated when the drawing is saved.
        /// </para>
        /// </summary>
        /// <param name="sheetIndex">A 1 based sheet number index</param>
        /// <returns>
        /// A <see cref="RefreshNoteTablesResult"/> structure containing the result of the operation:
        /// <list type="bullet">
        /// <item>
        /// <term><see cref="RefreshNoteTablesResult.Success"/></term>
        /// <description>
        ///     <c>true</c> if the note table(s) were updated successfully; otherwise, <c>false</c>.
        /// </description>
        /// </item>
        /// <item>
        /// <term><see cref="RefreshNoteTablesResult.Message"/></term>
        /// <description>
        ///     A descriptive message providing details about the result. May contain
        ///     error information if <see cref="RefreshNoteTablesResult.Success"/> is <c>false</c>.
        /// </description>
        /// </item>
        /// </list>
        /// </returns>
        /// <example>
        /// The following example shows how to call this method and inspect the result:
        /// <code language="csharp">
        /// var result = VersaNote.Interop.ModifyCustomNote(1, "This is the updated custom note text");
        /// if (result.Success)
        /// {
        ///     Console.WriteLine($"Custom note successfully updated");
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Failed to  update custom note: {result.Message}");
        /// }
        /// </code>
        /// </example>
        public RefreshNoteTablesResult RefreshNoteTables(int sheetIndex)
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(RefreshNoteTables));
            if (method != null)
            {
                object[] parameters = new object[] { sheetIndex };
                RefreshNoteTablesResult commitNoteUpdatesResult = (RefreshNoteTablesResult)method.Invoke(VersaNoteObject, parameters);
                return commitNoteUpdatesResult;
            }

            return new RefreshNoteTablesResult()
            {
                Success = false,
                Message = $"{MethodInfo.GetCurrentMethod().Name} method not found in Versa Note"
            };
        }

        /// <summary>
        /// Modify an existing custom note on the specified drawing sheet.
        /// <para>
        /// Prior to running this method, you should run GetAllSheetNotes() to get the ID of the custom note to modify.
        /// </para>
        /// </summary>
        /// <param name="sheetIndex">A 1 based sheet number index</param>
        /// <param name="noteId">The unique Id of the note to be updated</param>
        /// <param name="newNoteText">Note text</param>
        /// <param name="toggleSubNote">(Optional) Set true to toggle the sub-note indenting status</param>
        /// <param name="suppressTableRefresh">(Optional) Set true to suppress note table refresh for improved performance when making multiple API calls</param>
        /// <returns>
        /// A <see cref="UpdatedCustomNoteResult"/> structure containing the result of the operation:
        /// <list type="bullet">
        /// <item>
        /// <term><see cref="UpdatedCustomNoteResult.Success"/></term>
        /// <description><c>true</c> if the note was updated successfully; otherwise, <c>false</c>.</description>
        /// </item>
        /// <item>
        /// <term><see cref="UpdatedCustomNoteResult.Message"/></term>
        /// <description>
        ///     A descriptive message providing details about the result. May contain
        ///     error information if <see cref="UpdatedCustomNoteResult.Success"/> is <c>false</c>.
        /// </description>
        /// </item>
        /// </list>
        /// </returns>
        /// <example>
        /// The following example shows how to call this method and inspect the result:
        /// <code language="csharp">
        /// var result = VersaNote.Interop.ModifyCustomNote(1, "This is the updated custom note text");
        /// if (result.Success)
        /// {
        ///     Console.WriteLine($"Custom note successfully updated");
        /// }
        /// else
        /// {
        ///     Console.WriteLine($"Failed to  update custom note: {result.Message}");
        /// }
        /// </code>
        /// </example>
        public UpdatedCustomNoteResult UpdateCustomNote(int sheetIndex, string noteId, string newNoteText, bool toggleSubNote = false, bool suppressTableRefresh = false)
        {
            MethodInfo method = methods.FirstOrDefault(x => x.Name == nameof(UpdateCustomNote));
            if (method != null)
            {
                object[] parameters = new object[] { sheetIndex, noteId, newNoteText, toggleSubNote, suppressTableRefresh };
                UpdatedCustomNoteResult updatedCustomNoteResult = (UpdatedCustomNoteResult)method.Invoke(VersaNoteObject, parameters);
                return updatedCustomNoteResult;
            }

            return new UpdatedCustomNoteResult()
            {
                Success = false,
                Message = $"{MethodInfo.GetCurrentMethod().Name} method not found in Versa Note"
            };
        }
    }
}
