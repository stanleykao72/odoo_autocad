// OdooAutoCAD.App/Services/NavigationService.cs
// Navigation service for page navigation

using System.Windows.Controls;
using Serilog;

namespace OdooAutoCAD.App.Services;

/// <summary>
/// Interface for navigation service.
/// </summary>
public interface INavigationService
{
    Frame? Frame { get; set; }
    void NavigateTo(string pageName);
    void GoBack();
    bool CanGoBack { get; }
}

/// <summary>
/// Navigation service implementation.
/// Uses assembly-qualified type names to resolve page types.
/// </summary>
public class NavigationService : INavigationService
{
    public Frame? Frame { get; set; }

    public bool CanGoBack => Frame?.CanGoBack ?? false;

    public void NavigateTo(string pageName)
    {
        if (Frame == null) return;

        var typeName = $"OdooAutoCAD.App.Views.Pages.{pageName}Page";
        var pageType = typeof(NavigationService).Assembly.GetType(typeName);

        if (pageType != null)
        {
            Frame.Navigate(Activator.CreateInstance(pageType));
        }
        else
        {
            Log.Warning("Page type not found: {TypeName}", typeName);
        }
    }

    public void GoBack()
    {
        if (Frame?.CanGoBack == true)
        {
            Frame.GoBack();
        }
    }
}
