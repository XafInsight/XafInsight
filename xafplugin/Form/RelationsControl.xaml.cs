using NLog;
using System.Windows;
using System.Windows.Controls;
using xafplugin.Helpers;
using xafplugin.Modules;
using xafplugin.ViewModels;

namespace xafplugin.Form
{
    public partial class RelationsControl : UserControl
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly RelationsViewModel _viewModel;
        public RelationsControl()
        {
            InitializeComponent();
            var settings = new SettingsProvider();
            var environmentService = new EnvironmentService();
            var dialog = new MessageBoxService();

            if (string.IsNullOrEmpty(environmentService.DatabasePath))
            {
                logger.Warn("Geen databasepad gevonden. Auditfile moet eerst worden ingelezen.");
                MessageBox.Show("Er is geen database beschikbaar. Zorg ervoor dat de auditfile is ingelezen.");
                return;
            }

            this.DataContext = _viewModel = new RelationsViewModel();
            this.Unloaded += RelationsControl_Unloaded;
        }

        private void BtnToevoegen_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.AddRelation())
            {
                MessageBox.Show(
    "Relatie toegevoegd.",
    "Succes",
    MessageBoxButton.OK,
    MessageBoxImage.Information);
            }
        }

        private void BtnAnnuleren_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ClearFields();
        }

        private void BtnVerwijder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is TableRelation relation && DataContext is RelationsViewModel vm)
            {
                var result = MessageBox.Show(
                    $"Weet u zeker dat u deze relatie wilt verwijderen?",
                    "Bevestigen",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    vm.Relations.Remove(relation);
                }
            }
        }

        private void RelationsControl_Unloaded(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is RelationsViewModel vm)
            {
                vm.SaveRelationsToSettings();
            }
        }
    }
}
