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

using Windows.Storage;

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
			Windows.Storage.Pickers.FileOpenPicker fileOpenPicker = new();
			InitializeWithWindow.Initialize(fileOpenPicker, App.MainWindow.GetWindowHandle());
			fileOpenPicker.FileTypeFilter.Add("*");
			fileOpenPicker.FileTypeFilter.Add(".jpg");
			fileOpenPicker.FileTypeFilter.Add(".png");

			StorageFile? file = await fileOpenPicker.PickSingleFileAsync();
			if (file == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(file.Path, "ファイルを開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion

	#region ButtonFolderPickerClickedCommand
	[RelayCommand]
	private async Task ButtonFolderPickerClicked()
	{
		try
		{
			Windows.Storage.Pickers.FolderPicker folderPicker = new();
			InitializeWithWindow.Initialize(folderPicker, App.MainWindow.GetWindowHandle());

			StorageFolder? folder = await folderPicker.PickSingleFolderAsync();
			if (folder == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(folder.Path, "フォルダーを開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion

	#region ButtonFileSavePickerClickedCommand
	[RelayCommand]
	private async Task ButtonFileSavePickerClicked()
	{
		try
		{
			Windows.Storage.Pickers.FileSavePicker fileSavePicker = new();
			InitializeWithWindow.Initialize(fileSavePicker, App.MainWindow.GetWindowHandle());
			fileSavePicker.FileTypeChoices.Add("JPEG 画像", [".jpg"]);
			fileSavePicker.FileTypeChoices.Add("PNG 画像", [".png"]);

			StorageFile? file = await fileSavePicker.PickSaveFileAsync();
			if (file == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(file.Path, "ファイルを保存");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion

	#region ButtonFileOpenPicker2ClickedCommand
	[RelayCommand]
	private async Task ButtonFileOpenPicker2Clicked()
	{
		try
		{
			Microsoft.Windows.Storage.Pickers.FileOpenPicker fileOpenPicker = new(App.MainWindow.AppWindow.Id);
			fileOpenPicker.FileTypeFilter.Add("*");
			fileOpenPicker.FileTypeFilter.Add(".jpg");
			fileOpenPicker.FileTypeFilter.Add(".png");

			Microsoft.Windows.Storage.Pickers.PickFileResult? file = await fileOpenPicker.PickSingleFileAsync();
			if (file == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(file.Path, "ファイルを開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion
}
