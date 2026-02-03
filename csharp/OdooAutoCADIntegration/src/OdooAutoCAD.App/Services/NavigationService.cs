// OdooAutoCAD.App/Services/NavigationService.cs
// Navigation service for page navigation

using System.Windows.Controls;

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
/// </summary>
public class NavigationService : INavigationService
{
    public Frame? Frame { get; set; }

    public bool CanGoBack => Frame?.CanGoBack ?? false;

    public void NavigateTo(string pageName)
    {
        if (Frame == null) return;

        var pageType = Type.GetType($"OdooAutoCAD.App.Views.Pages.{pageName}Page");
        if (pageType != null)
        {
            Frame.Navigate(Activator.CreateInstance(pageType));
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
