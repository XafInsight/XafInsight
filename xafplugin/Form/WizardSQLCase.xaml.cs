using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    /// <summary>
    /// Code-behind for the SQL CASE wizard. Provides syntax highlighting, auto-suggestion, and text insertion for building CASE expressions.
    /// </summary>
    public partial class WizardSQLCase : Window
    {
        private readonly WizardSQLCaseViewModel _viewModel;
        private bool _isUpdating;
        private string _currentTypedWord;
        private int _currentWordStart = -1;
        private int _currentWordLength;
        private static readonly string[] SqlKeywords = { "WHERE", "WHEN", "THEN", "ELSE", "END", "AS", "CASE", "INTEGER", "REAL", "TEXT", "BLOB" };
        private static readonly string[] SqlLogicalOperators = { "AND", "OR", "NOT" };
        private static readonly string[] SqlComparisonOperators = { "=", "<>", "<", ">", "<=", ">=" };
        private static readonly string[] SqlSpecialOperators = { "IN", "BETWEEN", "LIKE", "IS", "NULL", "NOT" };
        private static readonly string[] NotAllowKeywords = { "{condition}", "{result}", "{", "}" };

        public WizardSQLCase(WizardSQLCaseViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            _viewModel.RequestClose += (s, result) =>
            {
                DialogResult = result;
                Close();
            };

            SqlRichTextBox.AcceptsReturn = true;
            SqlRichTextBox.Document.LineHeight = 1.0;

            string defaultSql =
                "CASE {optional} \r\n" +
                "    WHEN {condition} THEN {result}\r\n" +
                "    WHEN {condition} THEN {result}\r\n" +
                "    WHEN {condition} THEN {result}\r\n" +
                "    ELSE {result}";

            SqlRichTextBox.Document.Blocks.Clear();
            var paragraph = new Paragraph(new Run(defaultSql))
            {
                Margin = new Thickness(0),
                LineHeight = 1.0
            };
            SqlRichTextBox.Document.Blocks.Add(paragraph);

            HighlightSql();

            SqlRichTextBox.PreviewMouseDoubleClick += SqlRichTextBox_PreviewMouseDoubleClick;
            SqlRichTextBox.PreviewKeyDown += SqlRichTextBox_PreviewKeyDown;
            SuggestionsListBox.KeyDown += SuggestionsListBox_KeyDown;
        }

        private void SqlRichTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (_viewModel.IsSuggestionsVisible && _viewModel.Suggestions.Count > 0)
            {
                switch (e.Key)
                {
                    case Key.Tab:
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = 0;
                        ApplySelectedSuggestion();
                        e.Handled = true;
                        return;
                    case Key.Down:
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = 0;
                        else if (SuggestionsListBox.SelectedIndex < _viewModel.Suggestions.Count - 1)
                            SuggestionsListBox.SelectedIndex++;
                        SuggestionsListBox.ScrollIntoView(SuggestionsListBox.SelectedItem);
                        e.Handled = true;
                        return;
                    case Key.Up:
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = _viewModel.Suggestions.Count - 1;
                        else if (SuggestionsListBox.SelectedIndex > 0)
                            SuggestionsListBox.SelectedIndex--;
                        SuggestionsListBox.ScrollIntoView(SuggestionsListBox.SelectedItem);
                        e.Handled = true;
                        return;
                    case Key.Enter:
                        if (SuggestionsListBox.SelectedIndex >= 0)
                        {
                            ApplySelectedSuggestion();
                            e.Handled = true;
                            return;
                        }
                        break;
                    case Key.Escape:
                        _viewModel.Suggestions.Clear();
                        _viewModel.IsSuggestionsVisible = false;
                        e.Handled = true;
                        return;
                }
            }

            if (e.Key == Key.Enter)
            {
                var caretPos = SqlRichTextBox.CaretPosition;
                string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
                int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, caretPos).Text.Length;

                int lineStart = text.LastIndexOf('\n', Math.Max(0, caretOffset - 1)) + 1;
                int indentLength = 0;
                while (lineStart + indentLength < text.Length &&
                       (text[lineStart + indentLength] == ' ' || text[lineStart + indentLength] == '\t'))
                {
                    indentLength++;
                }

                caretPos.InsertLineBreak();
                if (indentLength > 0)
                {
                    string indent = text.Substring(lineStart, indentLength);
                    caretPos = SqlRichTextBox.CaretPosition;
                    caretPos.InsertTextInRun(indent);
                }

                e.Handled = true;
                NormalizeDocument();
            }
        }

        private void NormalizeDocument()
        {
            if (_isUpdating) return;
            _isUpdating = true;
            try
            {
                int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;
                string fullText = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;

                SqlRichTextBox.Document.Blocks.Clear();
                var paragraph = new Paragraph(new Run(fullText))
                {
                    Margin = new Thickness(0),
                    LineHeight = 1.0
                };
                SqlRichTextBox.Document.Blocks.Add(paragraph);

                HighlightSql();

                var newCaret = GetTextPointerAtOffset(caretOffset);
                if (newCaret != null)
                    SqlRichTextBox.CaretPosition = newCaret;
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void SqlRichTextBox_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Point p = e.GetPosition(SqlRichTextBox);
            TextPointer hit = SqlRichTextBox.GetPositionFromPoint(p, true);
            if (hit == null) return;

            int clickOffset = new TextRange(SqlRichTextBox.Document.ContentStart, hit).Text.Length;
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            if (clickOffset < 0 || clickOffset >= text.Length) return;

            var matches = Regex.Matches(text, @"\{[^}\s]+\}", RegexOptions.None, TimeSpan.FromMilliseconds(100));
            foreach (Match m in matches)
            {
                int start = m.Index;
                int end = m.Index + m.Length;
                if (clickOffset >= start && clickOffset < end)
                {
                    TextPointer sp = GetTextPointerAtOffset(start);
                    TextPointer ep = GetTextPointerAtOffset(end);
                    if (sp != null && ep != null)
                    {
                        SqlRichTextBox.Selection.Select(sp, ep);
                        e.Handled = true;
                    }
                    return;
                }
            }
        }

        private void SqlRichTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating) return;
            _isUpdating = true;
            try
            {
                _viewModel.SqlText = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
                int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;

                HighlightSql();

                TextPointer newCaret = GetTextPointerAtOffset(caretOffset);
                if (newCaret != null)
                    SqlRichTextBox.CaretPosition = newCaret;

                UpdateSuggestions();
                if (_viewModel.IsSuggestionsVisible)
                    PositionSuggestionsPopupAtCaret();
            }
            finally
            {
                _isUpdating = false;
            }
        }

        private void PositionSuggestionsPopupAtCaret()
        {
            var caret = SqlRichTextBox.CaretPosition.GetInsertionPosition(LogicalDirection.Forward);
            Rect rect = caret.GetCharacterRect(LogicalDirection.Forward);

            SuggestionsPopup.Placement = System.Windows.Controls.Primitives.PlacementMode.RelativePoint;
            SuggestionsPopup.PlacementTarget = SqlRichTextBox;
            SuggestionsPopup.HorizontalOffset = rect.Right;
            SuggestionsPopup.VerticalOffset = rect.Bottom;

            if (SuggestionsPopup.IsOpen)
            {
                double x = SuggestionsPopup.HorizontalOffset;
                SuggestionsPopup.HorizontalOffset = x + 0.1;
                SuggestionsPopup.HorizontalOffset = x;
            }
        }

        private void HighlightSql()
        {
            FlowDocument doc = SqlRichTextBox.Document;
            var fullRange = new TextRange(doc.ContentStart, doc.ContentEnd);
            fullRange.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Black);
            fullRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);

            var columnList = _viewModel.Columns.ToList();

            foreach (string word in SqlKeywords) HighlightOccurrences(word, Brushes.Blue, FontWeights.Normal);
            foreach (string op in SqlLogicalOperators) HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            foreach (string op in SqlComparisonOperators) HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            foreach (string op in SqlSpecialOperators) HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            foreach (string reserved in NotAllowKeywords) HighlightOccurrences(reserved, Brushes.Red, FontWeights.Normal);
            HighlightOccurrences("{optional}", Brushes.Gray, FontWeights.Normal);

            foreach (string column in columnList)
            {
                if (!string.IsNullOrWhiteSpace(column))
                    HighlightOccurrences(column, Brushes.Black, FontWeights.Bold);
            }
        }

        private void HighlightOccurrences(string target, Brush foreground, FontWeight fontWeight)
        {
            if (string.IsNullOrEmpty(target)) return;
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            int index = 0;
            while ((index = text.IndexOf(target, index, StringComparison.Ordinal)) >= 0)
            {
                TextPointer start = GetTextPointerAtOffset(index);
                TextPointer end = GetTextPointerAtOffset(index + target.Length);
                if (start != null && end != null)
                {
                    var range = new TextRange(start, end);
                    range.ApplyPropertyValue(TextElement.ForegroundProperty, foreground);
                    range.ApplyPropertyValue(TextElement.FontWeightProperty, fontWeight);
                }
                index += target.Length;
            }
        }

        private TextPointer GetTextPointerAtOffset(int offset)
        {
            TextPointer pointer = SqlRichTextBox.Document.ContentStart;
            int count = 0;
            while (pointer != null)
            {
                if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                {
                    string runText = pointer.GetTextInRun(LogicalDirection.Forward);
                    if (count + runText.Length >= offset)
                        return pointer.GetPositionAtOffset(offset - count);
                    count += runText.Length;
                    pointer = pointer.GetPositionAtOffset(runText.Length);
                }
                else
                {
                    pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
                }
            }
            return null;
        }

        private void UpdateSuggestions()
        {
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;
            caretOffset = Math.Max(0, Math.Min(caretOffset, text.Length));

            int start = caretOffset;
            while (start > 0 && IsIdentChar(text[start - 1])) start--;
            int end = caretOffset;
            while (end < text.Length && IsIdentChar(text[end])) end++;

            string wordUnderCaret = (start < end) ? text.Substring(start, end - start) : string.Empty;
            string typedPrefix = (start < caretOffset) ? text.Substring(start, caretOffset - start) : string.Empty;

            _currentWordStart = start;
            _currentWordLength = end - start;
            _currentTypedWord = typedPrefix;

            var suggestions = new List<string>();
            if (!string.IsNullOrWhiteSpace(typedPrefix))
            {
                void AddMatches(IEnumerable<string> source)
                {
                    foreach (var s in source)
                        if (s.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase) &&
                            !s.Equals(wordUnderCaret, StringComparison.OrdinalIgnoreCase))
                            suggestions.Add(s);
                }

                AddMatches(SqlKeywords);
                AddMatches(SqlLogicalOperators);
                AddMatches(SqlComparisonOperators);
                AddMatches(SqlSpecialOperators);
                AddMatches(_viewModel.Columns.Where(c => !string.IsNullOrEmpty(c)));
            }

            _viewModel.Suggestions.Clear();
            foreach (var s in suggestions.Distinct(StringComparer.OrdinalIgnoreCase))
                _viewModel.Suggestions.Add(s);

            _viewModel.IsSuggestionsVisible = _viewModel.Suggestions.Count > 0;
        }

        private static bool IsIdentChar(char c) =>
            char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '[' || c == ']';

        private void SuggestionsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SuggestionsListBox.SelectedItem is string suggestion && !string.IsNullOrEmpty(_currentTypedWord))
            {
                ReplaceTypedWord(suggestion);
                HighlightSql();
                _viewModel.Suggestions.Clear();
                _viewModel.IsSuggestionsVisible = false;
                _currentTypedWord = string.Empty;
            }
        }

        private void ReplaceTypedWord(string suggestion)
        {
            if (string.IsNullOrEmpty(suggestion) || _currentWordStart < 0) return;

            int startOffset = _currentWordStart;
            int endOffset = _currentWordStart + _currentWordLength;

            TextPointer startPointer = GetTextPointerAtOffset(startOffset);
            TextPointer endPointer = GetTextPointerAtOffset(endOffset);
            if (startPointer != null && endPointer != null)
            {
                SqlRichTextBox.Selection.Select(startPointer, endPointer);
                SqlRichTextBox.Selection.Text = suggestion;
                TextPointer newCaret = GetTextPointerAtOffset(startOffset + suggestion.Length);
                if (newCaret != null)
                    SqlRichTextBox.CaretPosition = newCaret;
            }
        }

        private void ColumnsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ColumnsListBox.SelectedItem is string column && !string.IsNullOrEmpty(column))
            {
                InsertTextAtCaret(column);
                HighlightSql();
            }
        }

        private void InsertTextAtCaret(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            TextPointer caret = SqlRichTextBox.CaretPosition;
            if (caret.GetPointerContext(LogicalDirection.Forward) != TextPointerContext.Text)
                caret = caret.GetInsertionPosition(LogicalDirection.Forward);
            caret.InsertTextInRun(text);
            TextPointer newCaret = caret.GetPositionAtOffset(text.Length);
            if (newCaret != null)
                SqlRichTextBox.CaretPosition = newCaret;
        }

        private void ApplySelectedSuggestion()
        {
            if (SuggestionsListBox.SelectedIndex >= 0)
            {
                string selected = _viewModel.Suggestions[SuggestionsListBox.SelectedIndex];
                ReplaceTypedWord(selected);
                HighlightSql();
                _viewModel.Suggestions.Clear();
                _viewModel.IsSuggestionsVisible = false;
                _currentTypedWord = string.Empty;
            }
        }

        private void SuggestionsListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && SuggestionsListBox.SelectedItem is string)
            {
                ApplySelectedSuggestion();
                SqlRichTextBox.Focus();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                _viewModel.Suggestions.Clear();
                _viewModel.IsSuggestionsVisible = false;
                SqlRichTextBox.Focus();
                e.Handled = true;
            }
        }
    }
}