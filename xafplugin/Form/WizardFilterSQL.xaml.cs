using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using xafplugin.Interfaces;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    /// <summary>
    /// Code‑behind for the SqlCaseQuery window.  This class wires up the WPF
    /// controls defined in the XAML to the view model and implements syntax
    /// highlighting, auto‑suggestion and insertion logic.  The implementation
    /// deliberately avoids third‑party libraries so that the window can be used
    /// directly in a .NET Framework 4.8 VSTO add‑in.
    /// </summary>
    public partial class WizardFilterSQL : Window
    {
        private readonly WizardFilterSQLViewModel _viewModel;
        private readonly IMessageBoxService _dialog = new MessageBoxService();
        private bool _isUpdating;
        private string _currentTypedWord;
        private int _currentWordStart = -1;
        private int _currentWordLength;
        private static readonly string[] SqlKeywords = { "WHERE", "WHEN", "THEN", "ELSE", "END", "AS", "CASE", "INTEGER", "REAL", "TEXT", "BLOB" };
        private static readonly string[] SqlLogicalOperators = { "AND", "OR", "NOT" };
        private static readonly string[] SqlComparisonOperators = { "=", "<>", "<", ">", "<=", ">=" };
        private static readonly string[] SqlSpecialOperators = { "IN", "BETWEEN", "LIKE", "IS", "NULL", "NOT" };
        private static readonly string[] NotAllowKeywords = { "{condition}", "{result}", "{", "}" };


        public WizardFilterSQL(WizardFilterSQLViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;
           
            _viewModel.RequestClose += (s, result) =>
            {
                DialogResult = result;
                Close();
            };

            // Configure RichTextBox defaults
            SqlRichTextBox.AcceptsReturn = true;
            SqlRichTextBox.Document.LineHeight = 1.0; // Tighter line spacing
            
            string defaultSql =
                "    WHERE {condition}\r\n";
            SqlRichTextBox.Document.Blocks.Clear();
            
            // Create a single paragraph with consistent formatting
            Paragraph paragraph = new Paragraph(new Run(defaultSql));
            paragraph.Margin = new Thickness(0); // Remove paragraph spacing
            paragraph.LineHeight = 1.0;
            SqlRichTextBox.Document.Blocks.Add(paragraph);
            
            HighlightSql();

            SqlRichTextBox.PreviewMouseDoubleClick += SqlRichTextBox_PreviewMouseDoubleClick;
            SqlRichTextBox.PreviewKeyDown += SqlRichTextBox_PreviewKeyDown;
            SuggestionsListBox.KeyDown += SuggestionsListBox_KeyDown;
        }

        private void SqlRichTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Handle suggestions navigation/selection if suggestions are visible
            if (_viewModel.IsSuggestionsVisible && _viewModel.Suggestions.Count > 0)
            {
                switch (e.Key)
                {
                    case Key.Tab:
                        // Select first item if nothing is selected
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = 0;
                        ApplySelectedSuggestion();
                        e.Handled = true;
                        return;
                        
                    case Key.Down:
                        // If nothing is selected, select the first item
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = 0;
                        else if (SuggestionsListBox.SelectedIndex < _viewModel.Suggestions.Count - 1)
                            SuggestionsListBox.SelectedIndex++; // Move to next item
                        
                        // Ensure the selected item is visible
                        SuggestionsListBox.ScrollIntoView(SuggestionsListBox.SelectedItem);
                        e.Handled = true;
                        return;
                        
                    case Key.Up:
                        // If nothing is selected, select the last item
                        if (SuggestionsListBox.SelectedIndex == -1)
                            SuggestionsListBox.SelectedIndex = _viewModel.Suggestions.Count - 1;
                        else if (SuggestionsListBox.SelectedIndex > 0)
                            SuggestionsListBox.SelectedIndex--; // Move to previous item
                        
                        // Ensure the selected item is visible
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
                        // If no suggestion is selected, fall through to normal Enter handling
                        break;
                        
                    case Key.Escape:
                        // Hide suggestions
                        _viewModel.Suggestions.Clear();
                        _viewModel.IsSuggestionsVisible = false;
                        e.Handled = true;
                        return;
                }
            }

            if (e.Key == Key.Enter)
            {
                TextPointer caretPos = SqlRichTextBox.CaretPosition;
                
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
                    caretPos = SqlRichTextBox.CaretPosition; // Get updated caret position
                    caretPos.InsertTextInRun(indent);
                }
                
                e.Handled = true; 
                
                NormalizeDocument();
            }
        }

        /// <summary>
        /// Normalizes the document structure by consolidating runs and applying consistent paragraph formatting.
        /// This helps prevent layout shifts when editing.
        /// </summary>
        private void NormalizeDocument()
        {
            if (_isUpdating)
                return;
            
            _isUpdating = true;
            try
            {
                int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;
                
                string fullText = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
                
                SqlRichTextBox.Document.Blocks.Clear();
                
               
                Paragraph paragraph = new Paragraph(new Run(fullText));
                paragraph.Margin = new Thickness(0); 
                paragraph.LineHeight = 1.0;
                SqlRichTextBox.Document.Blocks.Add(paragraph);
                
                HighlightSql();
                
                TextPointer newCaret = GetTextPointerAtOffset(caretOffset);
                if (newCaret != null)
                {
                    SqlRichTextBox.CaretPosition = newCaret;
                }
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

            // Offset bepalen
            int clickOffset = new TextRange(SqlRichTextBox.Document.ContentStart, hit).Text.Length;

            // Volledige tekst
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            if (clickOffset < 0 || clickOffset >= text.Length) return;

            // Zoek tokens {zonder spaties} zoals {waarde}, {waarde2}, {id_123}
            var matches = Regex.Matches(text, @"\{[^}\s]+\}", RegexOptions.None, TimeSpan.FromMilliseconds(100));

            foreach (Match m in matches)
            {
                int start = m.Index;
                int end = m.Index + m.Length; // exclusief

                // Alleen ingrijpen als je "direct op het woord" (binnen het token) dubbelklikt
                if (clickOffset >= start && clickOffset < end)
                {
                    TextPointer sp = GetTextPointerAtOffset(start);
                    TextPointer ep = GetTextPointerAtOffset(end);
                    if (sp != null && ep != null)
                    {
                        SqlRichTextBox.Selection.Select(sp, ep);
                        e.Handled = true; // voorkom standaard-gedrag
                    }
                    return; // klaar
                }
            }
        }



        /// <summary>
        /// Handles the TextChanged event of the RichTextBox.  When the user types
        /// text, this method highlights SQL keywords and column names and computes
        /// auto‑complete suggestions.  The highlighting logic is guarded by
        /// _isUpdating to avoid re‑entrancy when the text is modified in code.
        /// </summary>
        private void SqlRichTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isUpdating)
                return;
            _isUpdating = true;
            try
            {
                // Persist plain text to the view model if desired.  This may be
                // omitted if you prefer to work directly with the FlowDocument.
                _viewModel.SqlText = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;

                // Capture the caret's plain‑text offset so we can restore it after modifying the
                // document.  Using TextRange to measure the number of characters avoids
                // including structural positions (e.g. runs or paragraphs) which can
                // cause offsets to differ from the underlying string.
                int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;

                // Apply syntax highlighting.
                HighlightSql();

                // Restore the caret position after reformatting.  Without this the caret
                // may jump to the beginning of the document when formatting is applied.
                TextPointer newCaret = GetTextPointerAtOffset(caretOffset);
                if (newCaret != null)
                {
                    SqlRichTextBox.CaretPosition = newCaret;
                }

                // Update suggestions based on the word being typed at the caret.
                UpdateSuggestions();

                if (_viewModel.IsSuggestionsVisible)
                {
                    PositionSuggestionsPopupAtCaret();
                }
            }
            finally
            {
                _isUpdating = false;
            }
        }
        private void PositionSuggestionsPopupAtCaret()
        {
            // Gebruik een insertion position voor een stabiele rect
            var caret = SqlRichTextBox.CaretPosition.GetInsertionPosition(LogicalDirection.Forward);
            Rect rect = caret.GetCharacterRect(LogicalDirection.Forward);

            // RelativePoint: offsets zijn RELATIEF aan de RichTextBox
            SuggestionsPopup.Placement = System.Windows.Controls.Primitives.PlacementMode.RelativePoint;
            SuggestionsPopup.PlacementTarget = SqlRichTextBox;
            SuggestionsPopup.HorizontalOffset = rect.Right;
            SuggestionsPopup.VerticalOffset = rect.Bottom;

            // Forceer her-plaatsing als hij al open is
            if (SuggestionsPopup.IsOpen)
            {
                // kleine ‘nudge’ om reflow te triggeren
                double x = SuggestionsPopup.HorizontalOffset;
                SuggestionsPopup.HorizontalOffset = x + 0.1;
                SuggestionsPopup.HorizontalOffset = x;
            }
        }
        /// <summary>
        /// Highlights SQL keywords and column names in the RichTextBox.  Keywords are
        /// coloured blue and column names are rendered in bold.  This routine
        /// clears existing formatting prior to applying new formatting to avoid
        /// accumulating styles over time.  The implementation uses simple string
        /// matching rather than a full parser but is sufficient for CASE expressions.
        /// </summary>
        private void HighlightSql()
        {
            FlowDocument doc = SqlRichTextBox.Document;

            // Reset
            TextRange fullRange = new TextRange(doc.ContentStart, doc.ContentEnd);
            fullRange.ApplyPropertyValue(TextElement.ForegroundProperty, Brushes.Black);
            fullRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Normal);

            // Lijsten
            var keywordList = SqlKeywords;
            var logicalList = SqlLogicalOperators;
            var comparisonList = SqlComparisonOperators;
            var specialList = SqlSpecialOperators;
            var columnList = _viewModel.Columns.ToList();

            string plainText = fullRange.Text;

            // 1) Keywords (blauw)
            foreach (string word in keywordList)
            {
                HighlightOccurrences(word, Brushes.Blue, FontWeights.Normal);
            }

            // 2) Logische operatoren (ook blauw)
            foreach (string op in logicalList)
            {
                HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            }

            // 3) Vergelijkingsoperatoren (ook blauw)
            foreach (string op in comparisonList)
            {
                HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            }

            // 4) Speciale operatoren (ook blauw)
            foreach (string op in specialList)
            {
                HighlightOccurrences(op, Brushes.Blue, FontWeights.Normal);
            }

            foreach (string notAllow in NotAllowKeywords)
            {
                HighlightOccurrences(notAllow, Brushes.Red, FontWeights.Normal);
            }

            HighlightOccurrences("{optional}", Brushes.Gray, FontWeights.Normal);

            // 5) Kolommen (vet)
            foreach (string column in columnList)
            {
                if (!string.IsNullOrWhiteSpace(column))
                {
                    HighlightOccurrences(column, Brushes.Black, FontWeights.Bold);
                }
            }
        }


        /// <summary>
        /// Highlights all occurrences of a target word in the document.  Colour and
        /// font weight can be specified independently.  Matches are case‑sensitive.
        /// </summary>
        /// <param name="target">The word to highlight.</param>
        /// <param name="foreground">Foreground brush to apply.</param>
        /// <param name="fontWeight">Font weight to apply.</param>
        private void HighlightOccurrences(string target, Brush foreground, FontWeight fontWeight)
        {
            if (string.IsNullOrEmpty(target))
                return;
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            int index = 0;
            // Perform a case‑sensitive search.  Only highlight when the text in the
            // editor exactly matches the keyword or column name because SQL identifiers
            // may be case‑sensitive.  Using StringComparison.Ordinal ensures that
            // matching respects the exact character casing.
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

        /// <summary>
        /// Gets a TextPointer corresponding to a character offset within the document.
        /// Offsets are counted as plain text characters.  When iterating through
        /// paragraphs, runs and other elements this method skips structural
        /// boundaries so that the offset corresponds to a position within a run.
        /// Returns null if the offset is beyond the end of the document.
        /// </summary>
        /// <param name="offset">The zero‑based character offset.</param>
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
                    {
                        return pointer.GetPositionAtOffset(offset - count);
                    }
                    else
                    {
                        count += runText.Length;
                        pointer = pointer.GetPositionAtOffset(runText.Length);
                    }
                }
                else
                {
                    pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
                }
            }
            return null;
        }

        /// <summary>
        /// Updates the suggestion list based on the word currently being typed at the
        /// caret.  The suggestion logic is loosely based on the ICollectionView
        /// filtering pattern where a collection is filtered to display matching
        /// items【356508782686960†L21-L31】.  Here we build a filtered list manually and
        /// expose it via the view model.
        /// </summary>
        private void UpdateSuggestions()
        {
            // Retrieve full plain text
            string text = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.Document.ContentEnd).Text;
            int caretOffset = new TextRange(SqlRichTextBox.Document.ContentStart, SqlRichTextBox.CaretPosition).Text.Length;
            if (caretOffset < 0) caretOffset = 0;
            if (caretOffset > text.Length) caretOffset = text.Length;

            // Determine word boundaries
            int start = caretOffset;
            while (start > 0 && IsIdentChar(text[start - 1]))
                start--;
            int end = caretOffset;
            while (end < text.Length && IsIdentChar(text[end]))
                end++;

            string wordUnderCaret = (start < end) ? text.Substring(start, end - start) : string.Empty;
            string typedPrefix = (start < caretOffset) ? text.Substring(start, caretOffset - start) : string.Empty;

            _currentWordStart = start;
            _currentWordLength = end - start;
            _currentTypedWord = typedPrefix;

            var suggestions = new List<string>();

            if (!string.IsNullOrWhiteSpace(typedPrefix))
            {
                // SQL keywords
                foreach (string kw in SqlKeywords)
                {
                    if (kw.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase) &&
                        !kw.Equals(wordUnderCaret, StringComparison.OrdinalIgnoreCase))
                    {
                        suggestions.Add(kw);
                    }
                }

                // Logical operators
                foreach (string op in SqlLogicalOperators)
                {
                    if (op.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase) &&
                        !op.Equals(wordUnderCaret, StringComparison.OrdinalIgnoreCase))
                    {
                        suggestions.Add(op);
                    }
                }

                // Comparison operators
                foreach (string op in SqlComparisonOperators)
                {
                    if (op.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase) &&
                        !op.Equals(wordUnderCaret, StringComparison.OrdinalIgnoreCase))
                    {
                        suggestions.Add(op);
                    }
                }

                // Special operators
                foreach (string op in SqlSpecialOperators)
                {
                    if (op.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase) &&
                        !op.Equals(wordUnderCaret, StringComparison.OrdinalIgnoreCase))
                    {
                        suggestions.Add(op);
                    }
                }

                // Column names
                foreach (string col in _viewModel.Columns)
                {
                    if (!string.IsNullOrEmpty(col) &&
                        col.StartsWith(typedPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        suggestions.Add(col);
                    }
                }
            }

            _viewModel.Suggestions.Clear();
            foreach (var s in suggestions.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                _viewModel.Suggestions.Add(s);
            }
            _viewModel.IsSuggestionsVisible = _viewModel.Suggestions.Count > 0;
        }


        /// <summary>
        /// Handles the MouseDoubleClick event of the SuggestionsListBox.  When the
        /// user double‑clicks a suggestion, the typed word is replaced with the
        /// selected suggestion.  The editor is then highlighted again and the
        /// suggestion list is hidden.
        /// </summary>
        private void SuggestionsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (SuggestionsListBox.SelectedItem is string suggestion && !string.IsNullOrEmpty(_currentTypedWord))
            {
                ReplaceTypedWord(suggestion);
                // After insertion reapply highlighting and hide suggestions.
                HighlightSql();
                _viewModel.Suggestions.Clear();
                _viewModel.IsSuggestionsVisible = false;
                _currentTypedWord = string.Empty;
            }
        }

        /// <summary>
        /// Replaces the last typed word with the specified suggestion.  This is used
        /// when the user selects an item from the auto‑complete list.  The caret is
        /// positioned at the end of the inserted suggestion.
        /// </summary>
        /// <param name="suggestion">The full word to insert.</param>
        private void ReplaceTypedWord(string suggestion)
        {
            if (string.IsNullOrEmpty(suggestion))
                return;
            // Replace the entire word currently under the caret with the chosen suggestion.
            // _currentWordStart and _currentWordLength are maintained by UpdateSuggestions() and
            // represent the plain‑text offsets of the word boundaries.
            if (_currentWordStart < 0)
                return;

            int startOffset = _currentWordStart;
            int endOffset = _currentWordStart + _currentWordLength;
            TextPointer startPointer = GetTextPointerAtOffset(startOffset);
            TextPointer endPointer = GetTextPointerAtOffset(endOffset);
            if (startPointer != null && endPointer != null)
            {
                // Replace the selected range with the suggestion.
                SqlRichTextBox.Selection.Select(startPointer, endPointer);
                SqlRichTextBox.Selection.Text = suggestion;
                // Position the caret at the end of the inserted text.
                TextPointer newCaret = GetTextPointerAtOffset(startOffset + suggestion.Length);
                if (newCaret != null)
                {
                    SqlRichTextBox.CaretPosition = newCaret;
                }
            }
        }

        /// <summary>
        /// Handles double‑clicks on the ColumnsListBox.  When the user double‑clicks
        /// a column name it is inserted into the editor at the caret position.
        /// </summary>
        private void ColumnsListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ColumnsListBox.SelectedItem is string column && !string.IsNullOrEmpty(column))
            {
                InsertTextAtCaret(column);
                HighlightSql();
            }
        }

        /// <summary>
        /// Inserts the specified text into the RichTextBox at the current caret
        /// position.  This method writes into the current run or creates a new run
        /// if necessary.  After insertion the caret is advanced by the length of
        /// the inserted text.
        /// </summary>
        /// <param name="text">The text to insert.</param>
        private void InsertTextAtCaret(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;
            TextPointer caret = SqlRichTextBox.CaretPosition;
            if (caret.GetPointerContext(LogicalDirection.Forward) != TextPointerContext.Text)
            {
                // If we're at a structural position (e.g. start of document) then
                // create a new run by inserting an empty string first.
                caret = caret.GetInsertionPosition(LogicalDirection.Forward);
            }
            caret.InsertTextInRun(text);
            TextPointer newCaret = caret.GetPositionAtOffset(text.Length);
            if (newCaret != null)
            {
                SqlRichTextBox.CaretPosition = newCaret;
            }
        }

        /// <summary>
        /// Determines whether a character is considered part of an SQL identifier for the
        /// purposes of auto‑completion scanning.  Identifiers include letters, digits,
        /// underscores, dots (to allow Table.Column syntax) and square brackets (to
        /// support quoted identifiers like [Column]).  All other characters are
        /// treated as separators that mark word boundaries.
        /// </summary>
        /// <param name="c">The character to classify.</param>
        /// <returns><c>true</c> if the character belongs to an identifier; otherwise
        /// <c>false</c>.</returns>
        private static bool IsIdentChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '[' || c == ']';
        }

        private void ApplySelectedSuggestion()
        {
            if (SuggestionsListBox.SelectedIndex >= 0)
            {
                string selectedSuggestion = _viewModel.Suggestions[SuggestionsListBox.SelectedIndex];
                ReplaceTypedWord(selectedSuggestion);
                // After insertion reapply highlighting and hide suggestions
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
                // Give focus back to the text box
                SqlRichTextBox.Focus();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                // Hide suggestions on Escape
                _viewModel.Suggestions.Clear();
                _viewModel.IsSuggestionsVisible = false;
                // Give focus back to the text box
                SqlRichTextBox.Focus();
                e.Handled = true;
            }
        }
    }
}