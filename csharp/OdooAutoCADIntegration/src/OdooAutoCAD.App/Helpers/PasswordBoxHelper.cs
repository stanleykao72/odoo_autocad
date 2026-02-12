// OdooAutoCAD.App/Helpers/PasswordBoxHelper.cs
// MVVM PasswordBox binding helper using attached properties

using System.Windows;
using System.Windows.Controls;

namespace OdooAutoCAD.App.Helpers;

/// <summary>
/// Attached behavior helper for binding PasswordBox in MVVM scenarios.
/// Provides a secure way to bind passwords while maintaining MVVM separation.
/// </summary>
/// <remarks>
/// Usage in XAML:
/// <PasswordBox helpers:PasswordBoxHelper.BoundPassword="{Binding Password, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
/// </remarks>
public static class PasswordBoxHelper
{
    private static bool _isUpdating = false;

    #region BoundPassword Attached Property

    /// <summary>
    /// BoundPassword attached property for two-way binding with ViewModel.
    /// </summary>
    public static readonly DependencyProperty BoundPasswordProperty =
        DependencyProperty.RegisterAttached(
            "BoundPassword",
            typeof(string),
            typeof(PasswordBoxHelper),
            new FrameworkPropertyMetadata(
                string.Empty,
                FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                OnBoundPasswordChanged));

    public static string GetBoundPassword(DependencyObject d)
    {
        return (string)d.GetValue(BoundPasswordProperty);
    }

    public static void SetBoundPassword(DependencyObject d, string value)
    {
        d.SetValue(BoundPasswordProperty, value);
    }

    private static void OnBoundPasswordChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not PasswordBox passwordBox)
            return;

        // Prevent circular updates
        if (_isUpdating)
            return;

        // Update PasswordBox when ViewModel property changes
        passwordBox.PasswordChanged -= PasswordBox_PasswordChanged;
        
        _isUpdating = true;
        try
        {
            var newPassword = (string)e.NewValue;
            if (passwordBox.Password != newPassword)
            {
                passwordBox.Password = newPassword;
            }
        }
        finally
        {
            _isUpdating = false;
        }

        passwordBox.PasswordChanged += PasswordBox_PasswordChanged;
    }

    #endregion

    #region BindPassword Attached Property

    /// <summary>
    /// BindPassword attached property to enable/disable binding behavior.
    /// Set to true to activate password binding.
    /// </summary>
    public static readonly DependencyProperty BindPasswordProperty =
        DependencyProperty.RegisterAttached(
            "BindPassword",
            typeof(bool),
            typeof(PasswordBoxHelper),
            new PropertyMetadata(false, OnBindPasswordChanged));

    public static bool GetBindPassword(DependencyObject d)
    {
        return (bool)d.GetValue(BindPasswordProperty);
    }

    public static void SetBindPassword(DependencyObject d, bool value)
    {
        d.SetValue(BindPasswordProperty, value);
    }

    private static void OnBindPasswordChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not PasswordBox passwordBox)
            return;

        if ((bool)e.NewValue)
        {
            passwordBox.PasswordChanged += PasswordBox_PasswordChanged;
        }
        else
        {
            passwordBox.PasswordChanged -= PasswordBox_PasswordChanged;
        }
    }

    #endregion

    #region Event Handlers

    private static void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
    {
        if (sender is not PasswordBox passwordBox)
            return;

        // Prevent circular updates
        if (_isUpdating)
            return;

        _isUpdating = true;
        try
        {
            // Update ViewModel when PasswordBox changes
            SetBoundPassword(passwordBox, passwordBox.Password);
        }
        finally
        {
            _isUpdating = false;
        }
    }

    #endregion

    #region ClearPasswordOnFocus Attached Property (Optional Enhancement)

    /// <summary>
    /// ClearPasswordOnFocus attached property.
    /// When true, clears password when the control receives focus.
    /// </summary>
    public static readonly DependencyProperty ClearPasswordOnFocusProperty =
        DependencyProperty.RegisterAttached(
            "ClearPasswordOnFocus",
            typeof(bool),
            typeof(PasswordBoxHelper),
            new PropertyMetadata(false, OnClearPasswordOnFocusChanged));

    public static bool GetClearPasswordOnFocus(DependencyObject d)
    {
        return (bool)d.GetValue(ClearPasswordOnFocusProperty);
    }

    public static void SetClearPasswordOnFocus(DependencyObject d, bool value)
    {
        d.SetValue(ClearPasswordOnFocusProperty, value);
    }

    private static void OnClearPasswordOnFocusChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not PasswordBox passwordBox)
            return;

        if ((bool)e.NewValue)
        {
            passwordBox.GotFocus += PasswordBox_GotFocus;
        }
        else
        {
            passwordBox.GotFocus -= PasswordBox_GotFocus;
        }
    }

    private static void PasswordBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (sender is PasswordBox passwordBox)
        {
            passwordBox.Clear();
        }
    }

    #endregion
}
