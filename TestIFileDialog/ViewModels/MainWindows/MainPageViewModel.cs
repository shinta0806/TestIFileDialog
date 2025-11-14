// ============================================================================
// 
// メインページの ViewModel
// 
// ============================================================================

// ----------------------------------------------------------------------------
// 
// ----------------------------------------------------------------------------

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace TestIFileDialog.ViewModels;

public partial class MainPageViewModel : ObservableRecipient
{
	// ====================================================================
	// コンストラクター
	// ====================================================================

	/// <summary>
	/// メインコンストラクター
	/// </summary>
	public MainPageViewModel()
	{
	}

	// ====================================================================
	// public プロパティー
	// ====================================================================

	// --------------------------------------------------------------------
	// View 通信用のプロパティー
	// --------------------------------------------------------------------

	// --------------------------------------------------------------------
	// コマンド
	// --------------------------------------------------------------------

	#region ButtonFileOpenPickerClickedCommand
	[RelayCommand]
	private async Task ButtonFileOpenPickerClicked()
	{
		try
		{
			FileOpenPicker fileOpenPicker = new();
			InitializeWithWindow.Initialize(fileOpenPicker, App.MainWindow.GetWindowHandle());
			fileOpenPicker.FileTypeFilter.Add(".jpg");
			fileOpenPicker.FileTypeFilter.Add(".png");
			fileOpenPicker.FileTypeFilter.Add("*");

			StorageFile? file = await fileOpenPicker.PickSingleFileAsync();
			if (file == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(file.Path, "開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion
}
