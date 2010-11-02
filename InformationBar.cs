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
        private ITextDocument _document;
        private IEditorOperations _operations;
        private ITextUndoHistory _undoHistory;

        private bool _isDisposed = false;

        bool _dontShowAgain = false;

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
            textView.Closed += TextViewClosed;

            // Delay the initial check until the view gets focus
            textView.GotAggregateFocus += GotAggregateFocus;
        }

        void DisableInformationBar()
        {
            _dontShowAgain = true;
            this.CloseInformationBar();

            if (_document != null)
            {
                _document.FileActionOccurred -= FileActionOccurred;
                _document = null;
            }

            if (_textView != null)
            {
                _textView.GotAggregateFocus -= GotAggregateFocus;
                _textView.Closed -= TextViewClosed;
                _textView = null;
            }
        }

        void CheckTabsAndSpaces()
        {
            if (_dontShowAgain)
                return;

            ITextSnapshot snapshot = _textView.TextDataModel.DocumentBuffer.CurrentSnapshot;

            int tabSize = _textView.Options.GetOptionValue(DefaultOptions.TabSizeOptionId);

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
                    {
                        // We need to count to make sure there are enough spaces to go into a tab or a tab that follows the spaces
                        int countOfSpaces = 1;
                        for (int i = line.Start + 1; i < line.End; i++)
                        {
                            char ch = snapshot[i];
                            if (ch == ' ')
                            {
                                countOfSpaces++;
                                if (countOfSpaces >= tabSize)
                                {
                                    startsWithSpaces = true;
                                    break;
                                }
                            }
                            else if (ch == '\t')
                            {
                                startsWithSpaces = true;
                                break;
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

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
            if (_dontShowAgain)
                return;

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

        void TextViewClosed(object sender, EventArgs e)
        {
            DisableInformationBar();
        }

        #endregion

        #region Hiding and showing the information bar

        void Hide(object sender, RoutedEventArgs e)
        {
            this.CloseInformationBar();
        }

        void DontShowAgain(object sender, RoutedEventArgs e)
        {
            this.DisableInformationBar();
        }

        void CloseInformationBar()
        {
            if (this.Height == 0 || _dontShowAgain)
                return;

            // Since we're going to be closing, make sure focus is back in the editor
            _textView.VisualElement.Focus();

            ChangeHeightTo(0);
        }

        void ShowInformationBar()
        {
            if (this.Height > 0 || _dontShowAgain)
                return;

            ChangeHeightTo(27);
        }

        void ChangeHeightTo(double newHeight)
        {
            if (_dontShowAgain)
                return;

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
                int tabSize = _textView.Options.GetOptionValue(DefaultOptions.TabSizeOptionId);

                using (ITextEdit edit = _textView.TextBuffer.CreateEdit())
                {
                    foreach (var line in edit.Snapshot.Lines)
                    {
                        bool tabsAfterSpaces = false;
                        int column = 0;
                        int spanLength = 0;
                        int countOfLargestRunOfSpaces = 0;
                        int countOfCurrentRunOfSpaces = 0;

                        for (int i = line.Start; i < line.End; i++)
                        {
                            char ch = edit.Snapshot[i];

                            // Increment column or break, depending on the character
                            if (ch == ' ')
                            {
                                countOfCurrentRunOfSpaces++;
                                countOfLargestRunOfSpaces = Math.Max(countOfLargestRunOfSpaces, countOfCurrentRunOfSpaces);

                                column++;
                                spanLength++;
                            }
                            else if (ch == '\t')
                            {
                                if (countOfLargestRunOfSpaces > 0)
                                    tabsAfterSpaces = true;

                                countOfCurrentRunOfSpaces = 0;

                                column += tabSize - (column % tabSize);
                                spanLength++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // Only do a replace if this will have any effect
                        if (tabsAfterSpaces || countOfLargestRunOfSpaces >= tabSize)
                        {
                            int tabCount = column / tabSize;
                            int spaceCount = column % tabSize;

                            string newWhitespace = string.Format("{0}{1}",
                                                                 new string('\t', tabCount),
                                                                 new string(' ', spaceCount));

                            if (!edit.Replace(new Span(line.Start, spanLength), newWhitespace))
                                return false;
                        }
                    }

                    edit.Apply();
                    return !edit.Canceled;
                }
            });

            this.CloseInformationBar();
        }

        void Untabify(object sender, RoutedEventArgs e)
        {
            PerformActionInUndo(() =>
            {
                int tabSize = _textView.Options.GetOptionValue(DefaultOptions.TabSizeOptionId);

                using (ITextEdit edit = _textView.TextBuffer.CreateEdit())
                {
                    foreach (var line in edit.Snapshot.Lines)
                    {
                        bool hasTabs = false;
                        int column = 0;
                        int spanLength = 0;

                        for (int i = line.Start; i < line.End; i++)
                        {
                            char ch = edit.Snapshot[i];

                            if (ch == '\t')
                            {
                                hasTabs = true;

                                column += tabSize - (column % tabSize);
                                spanLength++;
                            }
                            else if (ch == ' ')
                            {
                                spanLength++;
                                column++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // Only do a replace if this will have any effect
                        if (hasTabs)
                        {
                            string newWhitespace = new string(' ', column);

                            if (!edit.Replace(new Span(line.Start, spanLength), newWhitespace))
                                return false;
                        }
                    }

                    edit.Apply();
                    return !edit.Canceled;
                }
            });

            this.CloseInformationBar();
        }

        void PerformActionInUndo(Func<bool> action)
        {
            ITrackingPoint anchor = _textView.TextSnapshot.CreateTrackingPoint(_textView.Selection.AnchorPoint.Position, PointTrackingMode.Positive);
            ITrackingPoint active = _textView.TextSnapshot.CreateTrackingPoint(_textView.Selection.ActivePoint.Position, PointTrackingMode.Positive);
            bool empty = _textView.Selection.IsEmpty;
            TextSelectionMode mode = _textView.Selection.Mode;

            using (var undo = _undoHistory.CreateTransaction("Untabify"))
            {
                _operations.AddBeforeTextBufferChangePrimitive();

                if (!action())
                {
                    undo.Cancel();
                    return;
                }

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

        #region IWpfTextViewMargin Members

        public FrameworkElement VisualElement
        {
            get
            {
                return this;
            }
        }

        #endregion

        #region ITextViewMargin Members

        public double MarginSize
        {
            get
            {
                return this.ActualHeight;
            }
        }

        public bool Enabled
        {
            get
            {
                return !_dontShowAgain;
            }
        }

        public ITextViewMargin GetTextViewMargin(string marginName)
        {
            return (marginName == InformationBarMargin.MarginName) ? (IWpfTextViewMargin)this : null;
        }

        public void Dispose()
        {
            this.DisableInformationBar();
        }

        #endregion
    }
}
