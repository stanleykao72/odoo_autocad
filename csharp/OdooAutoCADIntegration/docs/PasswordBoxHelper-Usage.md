# PasswordBoxHelper Usage Guide

## Overview

`PasswordBoxHelper` is an MVVM-friendly attached behavior that enables two-way binding for WPF PasswordBox controls. Since PasswordBox.Password is not a DependencyProperty, direct binding is not possible. This helper solves that problem while maintaining MVVM principles and security best practices.

## Location

**File**: `src/OdooAutoCAD.App/Helpers/PasswordBoxHelper.cs`

## Features

- ✅ Two-way data binding with ViewModel
- ✅ Thread-safe circular update prevention
- ✅ Compatible with CommunityToolkit.Mvvm
- ✅ Optional clear-on-focus behavior
- ✅ MVVM separation maintained (no code-behind required)

## Basic Usage

### 1. XAML Setup

Add the namespace to your XAML file:

```xaml
<Window xmlns:helpers="clr-namespace:OdooAutoCAD.App.Helpers"
        ...>
```

### 2. Bind to ViewModel Property

#### Option A: Using BoundPassword (Recommended)

```xaml
<PasswordBox helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
```

#### Option B: Using BindPassword + BoundPassword

```xaml
<PasswordBox helpers:PasswordBoxHelper.BindPassword="True"
             helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay}" />
```

### 3. ViewModel Implementation

```csharp
using CommunityToolkit.Mvvm.ComponentModel;

public partial class LoginViewModel : ObservableObject
{
    [ObservableProperty]
    private string _password = string.Empty;
    
    // Property is automatically generated as 'Password'
    // with INotifyPropertyChanged support
}
```

## Advanced Usage

### Clear Password on Focus

Useful for re-entry scenarios:

```xaml
<PasswordBox helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay}"
             helpers:PasswordBoxHelper.ClearPasswordOnFocus="True" />
```

### Complete Example: Odoo Login Form

**XAML (OdooPage.xaml)**:

```xaml
<Page x:Class="OdooAutoCAD.App.Views.Pages.OdooPage"
      xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:helpers="clr-namespace:OdooAutoCAD.App.Helpers"
      xmlns:vm="clr-namespace:OdooAutoCAD.App.ViewModels">
    
    <Page.DataContext>
        <vm:OdooViewModel />
    </Page.DataContext>
    
    <StackPanel Margin="20">
        <TextBlock Text="Odoo Connection" FontSize="20" FontWeight="Bold" Margin="0,0,0,20"/>
        
        <Label Content="Server URL:"/>
        <TextBox Text="{Binding ServerUrl, UpdateSourceTrigger=PropertyChanged}" />
        
        <Label Content="Database:"/>
        <TextBox Text="{Binding Database, UpdateSourceTrigger=PropertyChanged}" />
        
        <Label Content="Username:"/>
        <TextBox Text="{Binding Username, UpdateSourceTrigger=PropertyChanged}" />
        
        <Label Content="Password:"/>
        <PasswordBox helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
        
        <Button Content="Connect" 
                Command="{Binding ConnectCommand}" 
                Margin="0,20,0,0"/>
    </StackPanel>
</Page>
```

**ViewModel (OdooViewModel.cs)**:

```csharp
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.App.ViewModels;

public partial class OdooViewModel : ObservableObject
{
    private readonly IOdooService _odooService;
    
    [ObservableProperty]
    private string _serverUrl = string.Empty;
    
    [ObservableProperty]
    private string _database = string.Empty;
    
    [ObservableProperty]
    private string _username = string.Empty;
    
    [ObservableProperty]
    private string _password = string.Empty;
    
    [ObservableProperty]
    private bool _isConnected;
    
    [ObservableProperty]
    private string _statusMessage = string.Empty;
    
    public OdooViewModel(IOdooService odooService)
    {
        _odooService = odooService;
    }
    
    [RelayCommand]
    private async Task ConnectAsync()
    {
        try
        {
            StatusMessage = "Connecting...";
            
            var success = await _odooService.ConnectAsync(
                ServerUrl, 
                Database, 
                Username, 
                Password);
            
            IsConnected = success;
            StatusMessage = success 
                ? "Connected successfully!" 
                : "Connection failed";
                
            if (success)
            {
                // Clear password after successful connection
                Password = string.Empty;
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Error: {ex.Message}";
            IsConnected = false;
        }
    }
}
```

## Security Considerations

### ✅ Best Practices

1. **Never store plaintext passwords in memory longer than necessary**
   ```csharp
   // Clear password after use
   if (success)
   {
       Password = string.Empty;
   }
   ```

2. **Use SecureString for highly sensitive scenarios** (if required)
   - The current implementation uses string for simplicity
   - Can be extended to support SecureString if needed

3. **Don't log password values**
   ```csharp
   // ❌ NEVER DO THIS
   _logger.LogInformation($"Password: {Password}");
   
   // ✅ DO THIS
   _logger.LogInformation("Attempting connection with provided credentials");
   ```

4. **Store API tokens/passwords in secure storage**
   - Use Windows Credential Manager
   - Or encrypted database (as implemented in SettingsService)

### 🔐 Current Implementation Security

The PasswordBoxHelper:
- ✅ Uses WPF's PasswordBox which masks input visually
- ✅ Prevents circular updates to avoid memory leaks
- ✅ Clears resources when behavior is disabled
- ⚠️ Uses string (not SecureString) for convenience
  - Trade-off: Easier MVVM integration vs. maximum security
  - Acceptable for most enterprise scenarios
  - Can be extended if SecureString is required

## Testing

### Unit Test Example

```csharp
using Xunit;
using OdooAutoCAD.App.ViewModels;

public class OdooViewModelTests
{
    [Fact]
    public void Password_PropertyChanged_RaisesNotification()
    {
        // Arrange
        var vm = new OdooViewModel(Mock.Of<IOdooService>());
        var propertyChanged = false;
        vm.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName == nameof(vm.Password))
                propertyChanged = true;
        };
        
        // Act
        vm.Password = "test123";
        
        // Assert
        Assert.True(propertyChanged);
        Assert.Equal("test123", vm.Password);
    }
}
```

## Troubleshooting

### Issue: Password not updating in ViewModel

**Solution**: Ensure you're using `Mode=TwoWay` and `UpdateSourceTrigger=PropertyChanged`

```xaml
<PasswordBox helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
```

### Issue: Circular update warnings

**Solution**: The helper includes built-in protection. If you still see issues, check that you're not manually updating the PasswordBox.Password in code-behind.

### Issue: Password cleared unexpectedly

**Solution**: Check if `ClearPasswordOnFocus="True"` is enabled

## References

- [Stack Overflow: PasswordBox MVVM Binding](https://stackoverflow.com/questions/1483892/how-to-bind-to-a-passwordbox-in-mvvm)
- [WPF Attached Properties](https://docs.microsoft.com/en-us/dotnet/desktop/wpf/advanced/attached-properties-overview)
- [CommunityToolkit.Mvvm Documentation](https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/)

## Sprint 3 Phase 1 Compatibility

This helper is part of **Sprint 3 Phase 1: Foundation** and is designed to work seamlessly with:

- ✅ OdooSettings from ConfigurationLoader
- ✅ IOdooService.ConnectAsync() method
- ✅ IAutoCADService interface
- ✅ CommunityToolkit.Mvvm framework
- ✅ appsettings.json configuration system

## Version History

- **v6.0.0** - Initial implementation for Sprint 3 Phase 1
  - BoundPassword attached property
  - BindPassword attached property
  - ClearPasswordOnFocus optional enhancement
  - Circular update prevention
