using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using System.ComponentModel.Composition;

namespace CppXmlComments
{
    /// <summary>
    /// The XML comment completion source provider.
    /// </summary>
    [Export(typeof(ICompletionSourceProvider))]
    [ContentType("code")]
    [Name("xml comment completion")]
    class XmlCommentCompletionSourceProvider : ICompletionSourceProvider
    {
        /// <summary>
        /// The <see cref="ITextStructureNavigatorSelectorService"/>.
        /// </summary>
        [Import]
        internal ITextStructureNavigatorSelectorService NavigatorService { get; set; }

        /// <summary>
        /// The <see cref="IGlyphService"/>.
        /// </summary>
        [Import]
        internal IGlyphService GlyphService { get; set; }

        /// <summary>
        /// Tries to create a completion source.
        /// </summary>
        /// <param name="textBuffer">The <see cref="ITextBuffer"/>.</param>
        /// <returns>The <see cref="ICompletionSource"/>.</returns>
        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            return new XmlCommentCompletionSource(this, textBuffer);
        }
    }
}