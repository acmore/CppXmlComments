using EnvDTE;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System;
using System.ComponentModel.Composition;

namespace CppXmlComments
{
    /// <summary>
    /// The XML comment completion command handler provider.
    /// </summary>
    [Export(typeof(IVsTextViewCreationListener))]
    [Name("xml comment completion handler provider")]
    [ContentType("code")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    class XmlCommentCompletionCommandHandlerProvider : IVsTextViewCreationListener
    {
        /// <summary>
        /// The <see cref="IVsEditorAdaptersFactoryService"/>.
        /// </summary>
        [Import]
        internal IVsEditorAdaptersFactoryService AdapterService = null;

        /// <summary>
        /// The <see cref="ICompletionBroker"/>.
        /// </summary>
        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }

        /// <summary>
        /// The <see cref="SVsServiceProvider"/>.
        /// </summary>
        [Import]
        internal SVsServiceProvider ServiceProvider { get; set; }

        /// <summary>
        /// Registers the command handler.
        /// </summary>
        /// <param name="textViewAdapter">The <see cref="IVsTextView"/>.</param>
        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            try
            {
                // Get text view
                ITextView textView = AdapterService.GetWpfTextView(textViewAdapter);
                if (textView == null)
                    return;

                // Create a callback to create the command handler
                Func<XmlCommentCompletionCommandHandler> createCommandHandler = delegate ()
                {
                    var dte = this.ServiceProvider.GetService(typeof(DTE)) as DTE;
                    return new XmlCommentCompletionCommandHandler(textViewAdapter, textView, dte, this);
                };
                textView.Properties.GetOrCreateSingletonProperty(createCommandHandler);
            }
            catch(Exception)
            {
            }
        }
    }
}