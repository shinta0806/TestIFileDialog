using Microsoft.UI.Xaml.Controls;

using TestIFileDialog.ViewModels;

namespace TestIFileDialog.Views;

public sealed partial class MainPage : Page
{
    public MainViewModel ViewModel
    {
        get;
    }

    public MainPage()
    {
        ViewModel = App.GetService<MainViewModel>();
        InitializeComponent();
    }
}
