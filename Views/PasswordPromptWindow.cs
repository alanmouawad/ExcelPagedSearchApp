using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelPagedSearchApp.Views;

public sealed class PasswordPromptWindow : Window
{
    private readonly PasswordBox _passwordBox;

    public string? EnteredPassword { get; private set; }

    public PasswordPromptWindow(string fileName, string? message = null)
    {
        Title = "Encrypted Excel File";
        Width = 420;
        Height = 190;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;

        var root = new Grid { Margin = new Thickness(16) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var titleText = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(message)
                ? $"'{fileName}' is password-protected. Enter the workbook password:"
                : message,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 10)
        };
        Grid.SetRow(titleText, 0);
        root.Children.Add(titleText);

        _passwordBox = new PasswordBox
        {
            Height = 28,
            VerticalContentAlignment = VerticalAlignment.Center
        };
        _passwordBox.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter)
                Confirm();
        };
        Grid.SetRow(_passwordBox, 1);
        root.Children.Add(_passwordBox);

        var buttonsPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 14, 0, 0)
        };

        var okButton = new Button
        {
            Content = "Unlock",
            Width = 85,
            Height = 28,
            IsDefault = true,
            Margin = new Thickness(0, 0, 8, 0)
        };
        okButton.Click += (_, _) => Confirm();

        var cancelButton = new Button
        {
            Content = "Cancel",
            Width = 85,
            Height = 28,
            IsCancel = true
        };
        cancelButton.Click += (_, _) =>
        {
            EnteredPassword = null;
            DialogResult = false;
            Close();
        };

        buttonsPanel.Children.Add(okButton);
        buttonsPanel.Children.Add(cancelButton);
        Grid.SetRow(buttonsPanel, 3);
        root.Children.Add(buttonsPanel);

        Content = root;
        Loaded += (_, _) => _passwordBox.Focus();
    }

    private void Confirm()
    {
        EnteredPassword = _passwordBox.Password;
        DialogResult = true;
        Close();
    }
}