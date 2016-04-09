using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using System.Windows.Media;

namespace CppXmlComments
{
    /// <summary>
    /// The XML comment completion source.
    /// </summary>
    class XmlCommentCompletionSource : ICompletionSource
    {
        /// <summary>
        /// The <see cref="XmlCommentCompletionSourceProvider"/>.
        /// </summary>
        private XmlCommentCompletionSourceProvider sourceProvider;

        /// <summary>
        /// The <see cref="ITextBuffer"/>.
        /// </summary>
        private ITextBuffer textBuffer;

        /// <summary>
        /// The list of <see cref="Completion"/>.
        /// </summary>
        private List<Completion> completionList;

        /// <summary>
        /// The flag indicates whether was disposed.
        /// </summary>
        private bool isDisposed = false;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="sourceProvider">The <see cref="XmlCommentCompletionSourceProvider"/>.</param>
        /// <param name="textBuffer">The <see cref="ITextBuffer"/>.</param>
        public XmlCommentCompletionSource(XmlCommentCompletionSourceProvider sourceProvider, ITextBuffer textBuffer)
        {
            // Initialize
            this.sourceProvider = sourceProvider;
            this.textBuffer = textBuffer;
            ImageSource imageSource = this.sourceProvider.GlyphService.GetGlyph(StandardGlyphGroup.GlyphKeyword, StandardGlyphItem.GlyphItemPublic);

            // Initialize the comment tags
            this.completionList = new List<Completion>();
            this.completionList.Add(new Completion("<!-->", "<!---->", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<![CDATA[>", "<![CDATA[]]>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<c>", "<c></c>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<code>", "<code></code>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<example>", "<example></example>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<exception>", "<exception cref=\"\"></exception>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<include>", "<include file='' path='[@name=\"\"]'/>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<list>", "<list></list>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<para>", "<para></para>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<param>", "<param name=\"\"></param>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<paramref>", "<paramref name=\"\"></paramref>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<permission>", "<permission cref=\"\"></permission>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<remarks>", "<remarks></remarks>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<returns>", "<returns></returns>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<see>", "<see cref=\"\"/>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<seealso>", "<seealso cref=\"\"/>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<typeparam>", "<typeparam name=\"\"></typeparam>", string.Empty, imageSource, string.Empty));
            this.completionList.Add(new Completion("<value>", "<value></value>", string.Empty, imageSource, string.Empty));
        }

        /// <summary>
        /// Augments the completion session.
        /// </summary>
        /// <param name="session">The <see cref="ICompletionSession"/>.</param>
        /// <param name="completionSets">The list of <see cref="CompletionSet"/>.</param>
        public void AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            // Check if it was disposed
            if (this.isDisposed)
                return;

            // Check if the arguments are valid
            if (session == null || completionSets == null)
                return;

            // Check if the text buffer is null
            if (this.textBuffer == null)
                return;

            // Check content type
            if (this.textBuffer.ContentType.TypeName != "C/C++")
                return;

            // Get current point
            SnapshotPoint? currentPoint = session.GetTriggerPoint(this.textBuffer.CurrentSnapshot);
            if (!currentPoint.HasValue)
                return;

            // Get text
            string text = currentPoint.Value.GetContainingLine().GetText();
            if (String.IsNullOrEmpty(text) || !text.TrimStart().StartsWith("///"))
                return;

            // Add to completion set
            try
            {
                var trackingSpan = this.FindTokenSpanAtPosition(session.GetTriggerPoint(this.textBuffer), session);
                if (trackingSpan == null)
                    return;
                completionSets.Add(new CompletionSet(
                    "CppXmlCommentCompletionSet",
                    "XmlComment",
                    trackingSpan,
                    this.completionList,
                    null));
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// Finds the token span at current position.
        /// </summary>
        /// <param name="point">The <see cref="ITrackingPoint"/>.</param>
        /// <param name="session">The <see cref="ICompletionSession"/>.</param>
        /// <returns>The <see cref="ITrackingSpan"/>.</returns>
        private ITrackingSpan FindTokenSpanAtPosition(ITrackingPoint point, ICompletionSession session)
        {
            try
            {
                // Data validation
                if (point == null || session == null)
                    return null;

                // Find the token span at current position
                SnapshotPoint currentPoint = (session.TextView.Caret.Position.BufferPosition) - 1;
                ITextStructureNavigator navigator = this.sourceProvider.NavigatorService.GetTextStructureNavigator(this.textBuffer);
                TextExtent extent = navigator.GetExtentOfWord(currentPoint);
                return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
            }
            catch(Exception)
            {
            }
            return null;
        }

        public void Dispose()
        {
            if (!this.isDisposed)
            {
                try
                {
                    GC.SuppressFinalize(this);
                }
                catch(Exception)
                {
                }
                this.isDisposed = true;
            }
        }
    }
}