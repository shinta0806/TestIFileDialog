using Microsoft.UI.Xaml.Controls;

using TestIFileDialog.ViewModels;

namespace TestIFileDialog.Views.MainWindows;

public sealed partial class MainPage : Page
{
	public MainPageViewModel ViewModel
	{
		get;
	}

	public MainPage()
	{
		ViewModel = new();
		InitializeComponent();
	}
}
