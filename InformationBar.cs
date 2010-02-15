using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;

namespace FixMixedTabs
{
    sealed class InformationBarMargin : ContentControl, IWpfTextViewMargin
    {
        public const string MarginName = "InformationBar";
        private IWpfTextView _textView;
        private bool _isDisposed = false;
        private ITextDocument _document;
        private IEditorOperations _operations;
        private ITextUndoHistory _undoHistory;

        bool dontShowAgain = false;

        public InformationBarMargin(IWpfTextView textView, ITextDocument document, IEditorOperations editorOperations, ITextUndoHistory undoHistory)
        {
            _textView = textView;
            _document = document;
            _operations = editorOperations;
            _undoHistory = undoHistory;

            var informationBar = new InformationBarControl();
            informationBar.Tabify.Click += Tabify;
            informationBar.Untabify.Click += Untabify;
            informationBar.Hide.Click += Hide;
            informationBar.DontShowAgain.Click += DontShowAgain;

            this.Height = 0;
            this.Content = informationBar;
            this.Name = MarginName;

            document.FileActionOccurred += FileActionOccurred;

            // Delay the initial check until the view gets focus
            textView.GotAggregateFocus += GotAggregateFocus;
        }

        void CheckTabsAndSpaces()
        {
            if (dontShowAgain)
                return;

            ITextSnapshot snapshot = _textView.TextDataModel.DocumentBuffer.CurrentSnapshot;

            bool startsWithSpaces = false;
            bool startsWithTabs = false;

            foreach (var line in snapshot.Lines)
            {
                if (line.Length > 0)
                {
                    char firstChar = line.Start.GetChar();
                    if (firstChar == '\t')
                        startsWithTabs = true;
                    else if (firstChar == ' ')
                        startsWithSpaces = true;

                    if (startsWithSpaces && startsWithTabs)
                        break;
                }
            }

            if (startsWithTabs && startsWithSpaces)
                ShowInformationBar();
        }

        #region Event Handlers

        void FileActionOccurred(object sender, TextDocumentFileActionEventArgs e)
        {
            if ((e.FileActionType & FileActionTypes.ContentLoadedFromDisk) != 0 ||
                (e.FileActionType & FileActionTypes.ContentSavedToDisk) != 0)
            {
                CheckTabsAndSpaces();
            }
        }

        void GotAggregateFocus(object sender, EventArgs e)
        {
            _textView.GotAggregateFocus -= GotAggregateFocus;

            CheckTabsAndSpaces();
        }

        #endregion

        #region Hiding and showing the information bar

        void Hide(object sender, RoutedEventArgs e)
        {
            this.CloseInformationBar();
        }

        void DontShowAgain(object sender, RoutedEventArgs e)
        {
            this.dontShowAgain = true;
            this.CloseInformationBar();
        }

        void CloseInformationBar()
        {
            if (this.Height == 0)
                return;

            // Since we're going to be closing, make sure focus is back in the editor
            _textView.VisualElement.Focus();

            ChangeHeightTo(0);
        }

        void ShowInformationBar()
        {
            if (this.Height > 0 || dontShowAgain)
                return;

            ChangeHeightTo(27);
        }

        void ChangeHeightTo(double newHeight)
        {
            if (_textView.Options.GetOptionValue(DefaultWpfViewOptions.EnableSimpleGraphicsId))
            {
                this.Height = newHeight;
            }
            else
            {
                DoubleAnimation animation = new DoubleAnimation(this.Height, newHeight, new Duration(TimeSpan.FromMilliseconds(175)));
                Storyboard.SetTarget(animation, this);
                Storyboard.SetTargetProperty(animation, new PropertyPath(StackPanel.HeightProperty));

                Storyboard storyboard = new Storyboard();
                storyboard.Children.Add(animation);

                storyboard.Begin(this);
            }
        }

        #endregion

        #region Performing Tabify and Untabify

        void Tabify(object sender, RoutedEventArgs e)
        {
            PerformActionInUndo(() =>
            {
                _operations.SelectAll();
                _operations.Tabify();
            });

            this.CloseInformationBar();
        }

        void Untabify(object sender, RoutedEventArgs e)
        {
            PerformActionInUndo(() =>
            {
                _operations.SelectAll();
                _operations.Untabify();

            });

            this.CloseInformationBar();
        }

        void PerformActionInUndo(Action action)
        {
            ITrackingPoint anchor = _textView.TextSnapshot.CreateTrackingPoint(_textView.Selection.AnchorPoint.Position, PointTrackingMode.Positive);
            ITrackingPoint active = _textView.TextSnapshot.CreateTrackingPoint(_textView.Selection.ActivePoint.Position, PointTrackingMode.Positive);
            bool empty = _textView.Selection.IsEmpty;
            TextSelectionMode mode = _textView.Selection.Mode;

            using (var undo = _undoHistory.CreateTransaction("Untabify"))
            {
                _operations.AddBeforeTextBufferChangePrimitive();

                action();

                ITextSnapshot after = _textView.TextSnapshot;

                _operations.SelectAndMoveCaret(new VirtualSnapshotPoint(anchor.GetPoint(after)), 
                                               new VirtualSnapshotPoint(active.GetPoint(after)), 
                                               mode, 
                                               EnsureSpanVisibleOptions.ShowStart);

                _operations.AddAfterTextBufferChangePrimitive();

                undo.Complete();
            }

        }

        #endregion

        private void ThrowIfDisposed()
        {
            if (_isDisposed)
                throw new ObjectDisposedException(MarginName);
        }

        #region IWpfTextViewMargin Members

        public FrameworkElement VisualElement
        {
            get
            {
                ThrowIfDisposed();
                return this;
            }
        }

        #endregion

        #region ITextViewMargin Members

        public double MarginSize
        {
            get
            {
                ThrowIfDisposed();
                return this.ActualHeight;
            }
        }

        public bool Enabled
        {
            get
            {
                ThrowIfDisposed();
                return true;
            }
        }

        public ITextViewMargin GetTextViewMargin(string marginName)
        {
            return (marginName == InformationBarMargin.MarginName) ? (IWpfTextViewMargin)this : null;
        }

        public void Dispose()
        {
            // Nothing to do here.
        }

        #endregion
    }
}
