using System.Windows;
using OdooAutoCAD.App.ViewModels.Dialogs;

namespace OdooAutoCAD.App.Views.Dialogs;

public partial class ProductMappingDialog : Window
{
    public ProductMappingDialog()
    {
        InitializeComponent();
    }

    public ProductMappingDialog(ProductMappingDialogViewModel viewModel) : this()
    {
        DataContext = viewModel;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is ProductMappingDialogViewModel vm && vm.ApplyMapping())
        {
            DialogResult = true;
            Close();
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
