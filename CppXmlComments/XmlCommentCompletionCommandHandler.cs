using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CppXmlComments
{
    /// <summary>
    /// The XML comment completion command handler.
    /// </summary>
    class XmlCommentCompletionCommandHandler : IOleCommandTarget
    {
        /// <summary>
        /// The <see cref="ITextView"/>.
        /// </summary>
        private ITextView textView;

        /// <summary>
        /// The <see cref="IOleCommandTarget"/>.
        /// </summary>
        private IOleCommandTarget nextCommandHandler;

        /// <summary>
        /// The <see cref="XmlCommentCompletionCommandHandlerProvider"/>.
        /// </summary>
        private XmlCommentCompletionCommandHandlerProvider provider;

        /// <summary>
        /// The <see cref="ICompletionSession"/>.
        /// </summary>
        private ICompletionSession session;

        /// <summary>
        /// The <see cref="DTE"/>.
        /// </summary>
        private DTE dte;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="textViewAdapter">The <see cref="IVsTextView"/>.</param>
        /// <param name="textView">The <see cref="ITextView"/>.</param>
        /// <param name="dte">The <see cref="DTE"/>.</param>
        /// <param name="provider">The <see cref="XmlCommentCompletionCommandHandlerProvider"/>.</param>
        public XmlCommentCompletionCommandHandler(IVsTextView textViewAdapter, ITextView textView, DTE dte, XmlCommentCompletionCommandHandlerProvider provider)
        {
            // Check if the arguments are valid
            if (textViewAdapter == null || textView == null || dte == null || provider == null)
                return;

            // Initialize
            this.textView = textView;
            this.provider = provider;
            this.dte = dte;

            // Add to command chain and get the next command handler
            try
            {
                textViewAdapter.AddCommandFilter(this, out this.nextCommandHandler);
            }
            catch(Exception)
            {
            }
        }

        /// <summary>
        /// Queries status.
        /// </summary>
        /// <param name="pguidCmdGroup">The command group <see cref="Guid"/>.</param>
        /// <param name="cCmds">The command.</param>
        /// <param name="prgCmds">The array of <see cref="OLECMD"/>.</param>
        /// <param name="pCmdText">The command text.</param>
        /// <returns></returns>
        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            // Pass the command to the next command handler
            if (this.nextCommandHandler == null)
                return VSConstants.E_FAIL;
            try
            {
                return this.nextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }
            catch(Exception)
            {
                return VSConstants.E_FAIL;
            }
        }

        /// <summary>
        /// Triggers completion.
        /// </summary>
        /// <returns>True if successful; otherwise false.</returns>
        private bool TriggerCompletion()
        {
            // If the session is not null, no need to create a new one
            if (this.session != null)
                return true;

            // Check if the text view is null
            if (this.textView == null)
                return false;

            try
            {
                // The caret must be in a non-projection location
                SnapshotPoint? caretPoint = this.textView.Caret.Position.Point.GetPoint(textBuffer => (!textBuffer.ContentType.IsOfType("projection")), PositionAffinity.Predecessor);
                if (!caretPoint.HasValue)
                    return false;

                // Create a session
                this.session = this.provider.CompletionBroker.CreateCompletionSession(this.textView, caretPoint.Value.Snapshot.CreateTrackingPoint(caretPoint.Value.Position, PointTrackingMode.Positive), true);

                // Subscribe to the dismissed event
                this.session.Dismissed += this.OnSessionDismissed;

                // Start the session
                this.session.Start();
            }
            catch(Exception)
            {
                return false;
            }
            
            // Return success
            return true;
        }

        /// <summary>
        /// The session dismissed callback.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The event.</param>
        private void OnSessionDismissed(object sender, EventArgs e)
        {
            // Check if the session is null
            if (this.session == null)
                return;

            // Remove the dismissed event handler
            this.session.Dismissed -= this.OnSessionDismissed;

            // Set the session to null
            this.session = null;
        }

        /// <summary>
        /// Inserts comment at the given point.
        /// </summary>
        /// <param name="currentSnapshotLine">The <see cref="ITextSnapshotLine"/>.</param>
        /// <returns>True if successful; otherwise false.</returns>
        private bool InsertCommentAtPoint(ITextSnapshotLine currentSnapshotLine)
        {
            // Data validation
            if (currentSnapshotLine == null || this.dte == null)
                return false;

            // Initialize
            var currentLine = currentSnapshotLine.GetText();
            var indent = currentLine.Substring(0, currentLine.IndexOf("//"));
            var lineBreak = currentSnapshotLine.GetLineBreakText();
            lineBreak = String.IsNullOrEmpty(lineBreak) ? "\r\n" : lineBreak;
            var textSelection = this.dte.ActiveDocument.Selection as TextSelection;
            var lineNum = textSelection.ActivePoint.Line;
            var offset = textSelection.ActivePoint.LineCharOffset;

            // Move to the end of next line
            textSelection.LineDown();
            textSelection.EndOfLine();

            // Get code element
            var fileCodeModel = this.dte.ActiveDocument.ProjectItem.FileCodeModel;
            CodeElement codeElement = fileCodeModel == null ? null : fileCodeModel.CodeElementFromPoint(textSelection.ActivePoint, vsCMElement.vsCMElementFunction);

            // Add comments for function
            if (codeElement != null && codeElement is CodeFunction)
            {
                // Get function element
                CodeFunction function = codeElement as CodeFunction;

                // Don't generate comment if inside a function body
                var startPoint = function.GetStartPoint();
                if (startPoint != null && startPoint.Line <= lineNum)
                {
                    textSelection.MoveToLineAndOffset(lineNum, offset);
                    return false;
                }

                // Generate comment for function
                StringBuilder sb = new StringBuilder("/ <summary>" + lineBreak + indent + "/// " + lineBreak + indent + "/// </summary>");
                foreach (CodeElement child in codeElement.Children)
                {
                    CodeParameter parameter = child as CodeParameter;
                    if (parameter != null)
                        sb.AppendFormat("{0}{1}/// <param name=\"{2}\"></param>", lineBreak, indent, parameter.Name);
                }

                // If there is the return type is not void, generate a returns element
                if (function.Type.AsString != "void")
                    sb.AppendFormat("{0}{1}/// <returns></returns>", lineBreak, indent);

                // Move to summary element
                textSelection.MoveToLineAndOffset(lineNum, offset);
                textSelection.Insert(sb.ToString());
                textSelection.MoveToLineAndOffset(lineNum, offset);
                textSelection.LineDown();
                textSelection.EndOfLine();
                return true;
            }
            // TODO(acmore): Check if it's inside a function
            // Add summary comment
            else
            {
                // Generate a summary element
                textSelection.MoveToLineAndOffset(lineNum, offset);
                textSelection.Insert(String.Format("/ <summary>{0}{1}/// {0}{1}/// </summary>", lineBreak, indent));

                // Move to summary element
                textSelection.MoveToLineAndOffset(lineNum, offset);
                textSelection.LineDown();
                textSelection.EndOfLine();
                return true;
            }
        }

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="pguidCmdGroup">The command group <see cref="Guid"/>.</param>
        /// <param name="nCmdId">The command.</param>
        /// <param name="nCmdexecopt">The command exec opt.</param>
        /// <param name="pvaIn">The input argument.</param>
        /// <param name="pvaOut">The output argument.</param>
        /// <returns></returns>
        public int Exec(ref Guid pguidCmdGroup, uint nCmdId, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            // Data validation
            if (this.provider == null)
                return VSConstants.E_FAIL;

            try
            {
                // Ensure this function is not called in an automation context
                if (VsShellUtilities.IsInAutomationFunction(this.provider.ServiceProvider))
                {
                    if (this.nextCommandHandler == null)
                        return VSConstants.E_FAIL;
                    return this.nextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);
                }

                // Make a copy of this so we can look at it after forwarding some commands
                uint commandId = nCmdId;
                char key = char.MinValue;

                // Make sure the input is a char before getting it
                if (pguidCmdGroup == VSConstants.VSStd2K && nCmdId == (uint)VSConstants.VSStd2KCmdID.TYPECHAR)
                    key = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);

                // Check for a commit character
                if (nCmdId == (uint)VSConstants.VSStd2KCmdID.RETURN || nCmdId == (uint)VSConstants.VSStd2KCmdID.TAB || key == '>')
                {
                    // Check for a selection
                    if (this.session != null && !this.session.IsDismissed)
                    {
                        // If the selection is fully selected, commit the current session
                        if (this.session.SelectedCompletionSet.SelectionStatus.IsSelected)
                        {
                            // Commit the session
                            this.session.Commit();

                            // Move the caret
                            var textSelection = this.dte.ActiveDocument.Selection as TextSelection;
                            var currentLine = this.textView.TextSnapshot.GetLineFromPosition(this.textView.Caret.Position.BufferPosition.Position).GetText();
                            if (!String.IsNullOrEmpty(currentLine) && textSelection != null)
                            {
                                int index = currentLine.IndexOf('\'');
                                index = index < 0 ? currentLine.IndexOf("[]") : index;
                                index = index < 0 ? currentLine.IndexOf("--->") : index;
                                index = index < 0 ? currentLine.IndexOf('"') : index;
                                index = index < 0 ? currentLine.IndexOf('>') : index;
                                int offset = textSelection.CurrentColumn - 2 - index;
                                textSelection.CharLeft(false, offset);
                            }

                            // Don't add this character to the buffer
                            return VSConstants.S_OK;
                        }

                        // If there is no selection, dismiss the session
                        this.session.Dismiss();
                    }
                    // There is no active session, check if it's inside the XML block
                    else if (nCmdId == (uint)VSConstants.VSStd2KCmdID.RETURN && this.textView != null)
                    {
                        // Initialize
                        var textSnapshotLine = this.textView.TextSnapshot.GetLineFromPosition(this.textView.Caret.Position.BufferPosition.Position);
                        var currentLine = textSnapshotLine.GetText();
                        var lineBreak = textSnapshotLine.GetLineBreakText();
                        lineBreak = lineBreak == "" || lineBreak == null ? "\n" : lineBreak;
                        
                        // If inside the XML block, add  the leading ///
                        if (currentLine != null && currentLine.TrimStart().StartsWith("///") && this.dte != null)
                        {
                            string text = String.Format("{0}{1} ", lineBreak, currentLine.Substring(0, currentLine.IndexOf("///") + 3));
                            var textSelection = this.dte.ActiveDocument.Selection as TextSelection;
                            textSelection.Insert(text);
                            return VSConstants.S_OK;
                        }
                    }
                }
                // Insert comment block for a code block
                else if (!key.Equals(char.MinValue) && key == '/' && this.dte != null)
                {
                    var textSnapshotLine = this.textView.TextSnapshot.GetLineFromPosition(this.textView.Caret.Position.BufferPosition.Position);
                    var currentLine = textSnapshotLine.GetText();
                    if (currentLine.Trim() == "//" && this.InsertCommentAtPoint(textSnapshotLine))
                        return VSConstants.S_OK;
                }

                // Pass along the command so the char is added to the buffer
                int result = this.nextCommandHandler.Exec(ref pguidCmdGroup, nCmdId, nCmdexecopt, pvaIn, pvaOut);

                // Trigger the completion list
                if (!key.Equals(char.MinValue) && key == '<')
                {
                    // If there is no active session, create one
                    if (this.session == null || this.session.IsDismissed)
                        this.TriggerCompletion();

                    // Check if the session was created
                    if (this.session == null)
                        return result;

                    // Filter the session based on current input
                    this.session.Filter();
                    return VSConstants.S_OK;
                }
                else if (commandId == (uint)VSConstants.VSStd2KCmdID.BACKSPACE || commandId == (uint)VSConstants.VSStd2KCmdID.DELETE || char.IsLetterOrDigit(key))
                {
                    // Filter if the session is active
                    if (this.session != null && !this.session.IsDismissed)
                        this.session.Filter();
                    return VSConstants.S_OK;
                }
                return result;
            }
            catch(Exception)
            {
            }

            // Return failure
            return VSConstants.E_FAIL;
        }
    }
}