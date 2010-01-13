using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using Microsoft.VisualStudio.UI.Undo;

namespace FixMixedTabs
{
    #region InformationBar Factory
    [Export(typeof(IWpfTextViewMarginProvider))]
    [Name(InformationBarMargin.MarginName)]
    [MarginContainer(PredefinedMarginNames.Top)]
    [ContentType("any")]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    internal sealed class MarginFactory : IWpfTextViewMarginProvider
    {
        [Import]
        ITextDocumentFactoryService TextDocumentFactoryService = null;

        [Import]
        IEditorOperationsFactoryService OperationsFactory = null;

        [Import]
        IUndoHistoryRegistry UndoHistoryRegistry = null;

        public IWpfTextViewMargin CreateMargin(IWpfTextViewHost textViewHost, IWpfTextViewMargin containerMargin)
        {
            IWpfTextView view = textViewHost.TextView;

            ITextDocument document;
#if PostBeta2
            if (!TextDocumentFactoryService.TryGetTextDocument(view.TextDataModel.DocumentBuffer, out document))
                return null;
#endif
            if (!view.TextDataModel.DocumentBuffer.Properties.TryGetProperty(typeof(ITextDocument), out document))
                return null;

            return new InformationBarMargin(view, document, OperationsFactory.GetEditorOperations(view), UndoHistoryRegistry.RegisterHistory(view.TextBuffer));
        }
    }
    #endregion
}
