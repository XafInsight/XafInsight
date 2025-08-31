using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using xafplugin.Interfaces;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    public partial class WizardFilter : Window
    {
        private readonly WizardFilterViewModel _viewModel;
        private readonly IMessageBoxService _dialog = new MessageBoxService();

        public static readonly string[] SqlLogicalOperators = { "AND", "OR", "AND NOT" };
        public static readonly string[] SqlOperators = { "=", "<>", "<", ">", "<=", ">=", "IN", "LIKE" };
        public static readonly string[] NotAllowKeywords = { "{", "}" };

        private int conditionCount = 1;

        public WizardFilter(WizardFilterViewModel viewModel)
        {
            InitializeComponent();

            btnAddCondition.IsEnabled = false;
            SetupInitialFieldEvents();

            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            DataContext = _viewModel;

            _viewModel.RequestClose += (s, result) =>
            {
                DialogResult = result;
                Close();
            };

            foreach (var op in SqlOperators)
                cboOperator.Items.Add(op);
        }

        private void SetupInitialFieldEvents()
        {
            cboTable.SelectionChanged += Field_ValueChanged;
            cboOperator.SelectionChanged += Field_ValueChanged;
            txtValue.TextChanged += Field_ValueChanged;
            txtName.TextChanged += Field_ValueChanged;
        }

        private void Field_ValueChanged(object sender, EventArgs e) =>
            btnAddCondition.IsEnabled = ValidateFields();

        private bool ValidateFields()
        {
            bool all = true;
            foreach (UIElement element in filterConditionsPanel.Children)
            {
                if (element is GroupBox conditionGroup && conditionGroup.Content is Grid grid)
                {
                    ComboBox columnCombo = null;
                    ComboBox operatorCombo = null;
                    TextBox valueTextBox = null;

                    foreach (UIElement child in grid.Children)
                    {
                        if (child is ComboBox combo)
                        {
                            if (Grid.GetRow(child) == 0) columnCombo = combo;
                            else if (Grid.GetRow(child) == 1) operatorCombo = combo;
                        }
                        else if (child is TextBox tb && Grid.GetRow(child) == 2)
                        {
                            valueTextBox = tb;
                        }
                    }

                    if (columnCombo == null || columnCombo.SelectedItem == null) all = false;
                    if (operatorCombo == null || operatorCombo.SelectedItem == null) all = false;
                    if (valueTextBox == null || string.IsNullOrWhiteSpace(valueTextBox.Text)) all = false;
                }
                else if (element is ComboBox logicalCombo)
                {
                    if (logicalCombo.SelectedItem == null) all = false;
                }
            }
            return all;
        }

        private void btnAddCondition_Click(object sender, RoutedEventArgs e)
        {
            if (conditionCount >= 1)
            {
                var logicalOperatorCombo = CreateLogicalOperatorDropdown();
                filterConditionsPanel.Children.Add(logicalOperatorCombo);
            }

            conditionCount++;

            var group = new GroupBox { Margin = new Thickness(0, 0, 0, 10) };
            var headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var headerText = new TextBlock
            {
                Text = $"Condition {conditionCount}",
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(headerText, 0);

            var deleteButton = new Button
            {
                Content = "X",
                Style = FindResource("DeleteButtonStyle") as Style,
                Tag = conditionCount
            };
            deleteButton.Click += DeleteCondition_Click;
            Grid.SetColumn(deleteButton, 1);

            headerGrid.Children.Add(headerText);
            headerGrid.Children.Add(deleteButton);
            group.Header = headerGrid;

            var grid = new Grid { Margin = new Thickness(5) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var columnLabel = new Label { Content = "Column:", Margin = new Thickness(0, 0, 10, 5), VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(columnLabel, 0);
            Grid.SetColumn(columnLabel, 0);

            var columnCombo = new ComboBox { Margin = new Thickness(0, 0, 0, 5), Padding = new Thickness(5) };
            Grid.SetRow(columnCombo, 0);
            Grid.SetColumn(columnCombo, 1);
            foreach (var col in _viewModel.Columns)
                columnCombo.Items.Add(col);
            columnCombo.SelectionChanged += Field_ValueChanged;

            var operatorLabel = new Label { Content = "Operator:", Margin = new Thickness(0, 0, 10, 5), VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(operatorLabel, 1);
            Grid.SetColumn(operatorLabel, 0);

            var operatorCombo = new ComboBox { Margin = new Thickness(0, 0, 0, 5), Padding = new Thickness(5) };
            Grid.SetRow(operatorCombo, 1);
            Grid.SetColumn(operatorCombo, 1);
            foreach (var op in SqlOperators)
                operatorCombo.Items.Add(op);
            operatorCombo.SelectionChanged += Field_ValueChanged;

            var valueLabel = new Label { Content = "Value:", Margin = new Thickness(0, 0, 10, 5), VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(valueLabel, 2);
            Grid.SetColumn(valueLabel, 0);

            var valueBox = new TextBox { Margin = new Thickness(0, 0, 0, 5), Padding = new Thickness(5) };
            Grid.SetRow(valueBox, 2);
            Grid.SetColumn(valueBox, 1);
            valueBox.TextChanged += Field_ValueChanged;

            grid.Children.Add(columnLabel);
            grid.Children.Add(columnCombo);
            grid.Children.Add(operatorLabel);
            grid.Children.Add(operatorCombo);
            grid.Children.Add(valueLabel);
            grid.Children.Add(valueBox);

            group.Content = grid;
            filterConditionsPanel.Children.Add(group);
            btnAddCondition.IsEnabled = false;
        }

        private ComboBox CreateLogicalOperatorDropdown()
        {
            var logical = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(5),
                HorizontalAlignment = HorizontalAlignment.Center,
                Width = 100,
                Name = "cboLogicalOperator"
            };
            foreach (var op in SqlLogicalOperators)
                logical.Items.Add(op);
            logical.SelectedIndex = 0;
            logical.SelectionChanged += Field_ValueChanged;
            return logical;
        }

        public string BuildFilterConditionString()
        {
            var sb = new StringBuilder();
            bool first = true;

            foreach (UIElement element in filterConditionsPanel.Children)
            {
                if (element is GroupBox group && group.Content is Grid grid)
                {
                    string column = null, op = null, value = null;
                    foreach (UIElement c in grid.Children)
                    {
                        if (c is ComboBox combo)
                        {
                            if (Grid.GetRow(c) == 0 && combo.SelectedItem != null) column = combo.SelectedItem.ToString();
                            else if (Grid.GetRow(c) == 1 && combo.SelectedItem != null) op = combo.SelectedItem.ToString();
                        }
                        else if (c is TextBox tb && Grid.GetRow(c) == 2)
                        {
                            value = tb.Text?.Trim();
                        }
                    }
                    if (!string.IsNullOrEmpty(column) && !string.IsNullOrEmpty(op) && !string.IsNullOrEmpty(value))
                    {
                        string formatted;
                        if (op == "LIKE")
                        {
                            formatted = $"{column} {op} '{value}'";
                        }
                        else if (op == "IN")
                        {
                            formatted = $"{column} {op} ({value})";
                        }
                        else
                        {
                            bool isNumeric = decimal.TryParse(value, out _);
                            formatted = isNumeric ? $"{column} {op} {value}" : $"{column} {op} '{value}'";
                        }
                        sb.Append(formatted);
                    }
                }
                else if (element is ComboBox logicalOp && logicalOp.SelectedItem != null)
                {
                    if (!first)
                    {
                        sb.Append(" " + logicalOp.SelectedItem + " ");
                    }
                }
                first = false;
            }

            return sb.ToString();
        }

        public void btnComplete_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateFields())
            {
                _dialog.ShowWarning("Please fill in all fields before completing.");
                return;
            }
            var filterName = txtName.Text;
            if (string.IsNullOrWhiteSpace(filterName))
            {
                _dialog.ShowWarning("Please enter a filter name.");
                return;
            }

            string filterCondition = BuildFilterConditionString();
            if (string.IsNullOrWhiteSpace(filterCondition))
            {
                _dialog.ShowWarning("Please add at least one valid condition.");
                return;
            }

            if (_viewModel.SetSQLSyntax(filterCondition))
            {
                _viewModel.FilterName = filterName;
                _viewModel.OnRequestClose(true);
                Close();
            }
            else
            {
                _dialog.ShowError("The constructed SQL condition is not valid. Please review your conditions.");
            }
        }

        private void DeleteCondition_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button deleteButton)
            {
                var group = FindParent<GroupBox>(deleteButton);
                if (group == null) return;

                int idx = filterConditionsPanel.Children.IndexOf(group);
                if (idx < 0) return;

                if (idx > 0 && filterConditionsPanel.Children[idx - 1] is ComboBox)
                {
                    filterConditionsPanel.Children.RemoveAt(idx - 1);
                    idx--;
                }
                else if (idx + 1 < filterConditionsPanel.Children.Count &&
                         filterConditionsPanel.Children[idx + 1] is ComboBox)
                {
                    filterConditionsPanel.Children.RemoveAt(idx + 1);
                }

                filterConditionsPanel.Children.RemoveAt(idx);
                conditionCount--;
                RenumberConditions();
                btnAddCondition.IsEnabled = ValidateFields();
            }
        }

        private T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            var parent = System.Windows.Media.VisualTreeHelper.GetParent(child);
            if (parent == null) return null;
            if (parent is T typed) return typed;
            return FindParent<T>(parent);
        }

        private void RenumberConditions()
        {
            int i = 1;
            foreach (UIElement element in filterConditionsPanel.Children)
            {
                if (element is GroupBox group && group.Header is Grid headerGrid)
                {
                    foreach (UIElement h in headerGrid.Children)
                    {
                        if (h is TextBlock tb)
                            tb.Text = $"Condition {i}";
                        else if (h is Button btn)
                            btn.Tag = i;
                    }
                    i++;
                }
            }
        }
    }
}
