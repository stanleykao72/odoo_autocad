using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.App.ViewModels.Dialogs;

/// <summary>
/// Display model for a product mapping entry in the management view.
/// </summary>
public class ProductMappingItem
{
    public string AutoCADName { get; set; } = string.Empty;
    public int OdooProductId { get; set; }
    public string OdooProductName { get; set; } = string.Empty;
}

/// <summary>
/// ViewModel for the Product Mapping Dialog.
/// Supports two modes: single mapping (from validation panel) and management (all mappings).
/// </summary>
public partial class ProductMappingDialogViewModel : ObservableObject
{
    private readonly IOdooService _odooService;
    private readonly IBOQProcessor _boqProcessor;

    public ObservableCollection<OdooProduct> OdooProducts { get; } = new();
    public ObservableCollection<ProductMappingItem> ProductMappings { get; } = new();

    // Mode
    [ObservableProperty]
    private bool _isManagementMode;

    // Single mapping mode
    [ObservableProperty]
    private string _autoCADProductName = string.Empty;

    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private bool _isSearching;

    [ObservableProperty]
    private OdooProduct? _selectedProduct;

    // Management mode — new mapping
    [ObservableProperty]
    private string _newAutoCADName = string.Empty;

    [ObservableProperty]
    private OdooProduct? _newSelectedProduct;

    [ObservableProperty]
    private ProductMappingItem? _selectedMapping;

    public bool HasSelectedProduct => SelectedProduct != null;

    public ProductMappingDialogViewModel(IOdooService odooService, IBOQProcessor boqProcessor)
    {
        _odooService = odooService;
        _boqProcessor = boqProcessor;
    }

    partial void OnSelectedProductChanged(OdooProduct? value)
    {
        OnPropertyChanged(nameof(HasSelectedProduct));
    }

    [RelayCommand(CanExecute = nameof(CanSearch))]
    private async Task SearchAsync()
    {
        if (string.IsNullOrWhiteSpace(SearchText)) return;

        IsSearching = true;
        try
        {
            var products = await _odooService.SearchProductsAsync(SearchText);
            OdooProducts.Clear();
            foreach (var p in products)
            {
                OdooProducts.Add(p);
            }
        }
        finally
        {
            IsSearching = false;
        }
    }

    private bool CanSearch() => !string.IsNullOrWhiteSpace(SearchText) && !IsSearching;

    [RelayCommand]
    private void LoadMappings()
    {
        ProductMappings.Clear();
        var mappings = _boqProcessor.GetProductMappings();
        foreach (var kvp in mappings)
        {
            ProductMappings.Add(new ProductMappingItem
            {
                AutoCADName = kvp.Key,
                OdooProductId = kvp.Value,
                OdooProductName = $"Product #{kvp.Value}"
            });
        }
    }

    [RelayCommand]
    private void AddMapping()
    {
        if (string.IsNullOrWhiteSpace(NewAutoCADName) || NewSelectedProduct == null) return;

        _boqProcessor.SetProductMapping(NewAutoCADName, NewSelectedProduct.Id);

        ProductMappings.Add(new ProductMappingItem
        {
            AutoCADName = NewAutoCADName,
            OdooProductId = NewSelectedProduct.Id,
            OdooProductName = NewSelectedProduct.Name
        });

        NewAutoCADName = string.Empty;
        NewSelectedProduct = null;
    }

    [RelayCommand]
    private void RemoveMapping()
    {
        if (SelectedMapping == null) return;

        _boqProcessor.RemoveProductMapping(SelectedMapping.AutoCADName);
        ProductMappings.Remove(SelectedMapping);
        SelectedMapping = null;
    }

    /// <summary>
    /// Apply the selected mapping (single mode). Returns true if a product was selected.
    /// </summary>
    public bool ApplyMapping()
    {
        if (SelectedProduct == null || string.IsNullOrWhiteSpace(AutoCADProductName))
            return false;

        _boqProcessor.SetProductMapping(AutoCADProductName, SelectedProduct.Id);
        return true;
    }
}
