// ============================================================================
// 
// メインページの ViewModel
// 
// ============================================================================

// ----------------------------------------------------------------------------
// 
// ----------------------------------------------------------------------------

using System.Runtime.InteropServices;

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.System.Com;
using Windows.Win32.UI.Shell;
using Windows.Win32.UI.Shell.Common;

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

	#region ButtonIFileOpenDialogFileClickedCommand
	[RelayCommand]
	private async Task ButtonIFileOpenDialogFileClicked()
	{
		try
		{
			String filter = "すべてのファイル|*.*|JPEG 画像|*.jpg;*.jpeg|PNG 画像|*.png";
			UInt32 filterIndex = 1;
			Guid guid = new("F25059A4-FB2C-4A03-B3A0-A8E4A800EEBB");
			FILEOPENDIALOGOPTIONS options = FILEOPENDIALOGOPTIONS.FOS_NOCHANGEDIR | FILEOPENDIALOGOPTIONS.FOS_PATHMUSTEXIST | FILEOPENDIALOGOPTIONS.FOS_FILEMUSTEXIST/* | FILEOPENDIALOGOPTIONS.FOS_ALLOWMULTISELECT*/;
			String[]? pathes = ShowFileOpenDialog(filter, ref filterIndex, options, guid);
			if (pathes == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(String.Join('\n', pathes), "ファイルを開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion

	#region ButtonIFileOpenDialogFolderClickedCommand
	[RelayCommand]
	private async Task ButtonIFileOpenDialogFolderClicked()
	{
		try
		{
			String filter = String.Empty;
			UInt32 filterIndex = 1;
			Guid guid = new("BEBA95F9-C249-4ABA-964F-AFEAE4C1D47B");
			FILEOPENDIALOGOPTIONS options = FILEOPENDIALOGOPTIONS.FOS_NOCHANGEDIR | FILEOPENDIALOGOPTIONS.FOS_PATHMUSTEXIST | FILEOPENDIALOGOPTIONS.FOS_PICKFOLDERS;
			String[]? pathes = ShowFileOpenDialog(filter, ref filterIndex, options, guid);
			if (pathes == null)
			{
				return;
			}

			await App.MainWindow.ShowMessageDialogAsync(String.Join('\n', pathes), "フォルダーを開く");
		}
		catch (Exception ex)
		{
			await App.MainWindow.ShowMessageDialogAsync(ex.Message, "エラー");
		}
	}
	#endregion

	#region ButtonFileOpenPickerClickedCommand
	[RelayCommand]
	private async Task ButtonFileOpenPickerClicked()
	{
		try
		{
			// Microsoft 名前空間との対比のために FileOpenPicker コンストラクターを使用しているが、
			// WindowEx.CreateOpenFilePicker() を使用すれば 1 行削減可能（FolderPicker や FileSavePicker も同様）
			Windows.Storage.Pickers.FileOpenPicker fileOpenPicker = new();
			InitializeWithWindow.Initialize(fileOpenPicker, App.MainWindow.GetWindowHandle());
			fileOpenPicker.FileTypeFilter.Add("*");
			fileOpenPicker.FileTypeFilter.Add(".jpg");
			fileOpenPicker.FileTypeFilter.Add(".png");

			Windows.Storage.StorageFile? file = await fileOpenPicker.PickSingleFileAsync();
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

			Windows.Storage.StorageFolder? folder = await folderPicker.PickSingleFolderAsync();
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

			Windows.Storage.StorageFile? file = await fileSavePicker.PickSaveFileAsync();
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

	#region ButtonFolderPicker2ClickedCommand
	[RelayCommand]
	private async Task ButtonFolderPicker2Clicked()
	{
		try
		{
			Microsoft.Windows.Storage.Pickers.FolderPicker folderPicker = new(App.MainWindow.AppWindow.Id);
			Microsoft.Windows.Storage.Pickers.PickFolderResult? folder = await folderPicker.PickSingleFolderAsync();
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

	#region ButtonFileSavePicker2ClickedCommand
	[RelayCommand]
	private async Task ButtonFileSavePicker2Clicked()
	{
		try
		{
			Microsoft.Windows.Storage.Pickers.FileSavePicker fileSavePicker = new(App.MainWindow.AppWindow.Id);
			fileSavePicker.FileTypeChoices.Add("JPEG 画像", [".jpg"]);
			fileSavePicker.FileTypeChoices.Add("PNG 画像", [".png"]);

			Microsoft.Windows.Storage.Pickers.PickFileResult? file = await fileSavePicker.PickSaveFileAsync();
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

	// ====================================================================
	// private 関数
	// ====================================================================

	/// <summary>
	/// ファイルを開くダイアログを表示
	/// </summary>
	/// <param name="filter">（ShowFileDialogCore() 参照）</param>
	/// <param name="filterIndex">（ShowFileDialogCore() 参照）</param>
	/// <param name="options">（ShowFileDialogCore() 参照）</param>
	/// <param name="guid">（ShowFileDialogCore() 参照）</param>
	/// <returns>選択されたファイル群のパス（キャンセルされた場合は null）</returns>
	private unsafe String[]? ShowFileOpenDialog(String filter, ref UInt32 filterIndex, FILEOPENDIALOGOPTIONS options, Guid guid)
	{
		IFileOpenDialog* openDialog = null;
		IShellItemArray* shellItemArray = null;

		// finally 用の try
		try
		{
			// ダイアログ生成
			HRESULT result = PInvoke.CoCreateInstance(typeof(FileOpenDialog).GUID, null, CLSCTX.CLSCTX_INPROC_SERVER, out openDialog);
			result.ThrowOnFailure();

			// 表示
			if (!ShowFileDialogCore((IFileDialog*)openDialog, filter, ref filterIndex, options, guid))
			{
				return null;
			}

			// 結果取得
			result = openDialog->GetResults(&shellItemArray);
			if (result.Failed)
			{
				return null;
			}
			result = shellItemArray->GetCount(out UInt32 numPathes);
			if (result.Failed)
			{
				return null;
			}
			String[] pathes = new String[numPathes];
			for (UInt32 i = 0; i < numPathes; i++)
			{
				IShellItem* iShellResult;
				result = shellItemArray->GetItemAt(i, &iShellResult);
				if (result.Failed)
				{
					continue;
				}
				String? path = ShellItemToPath(iShellResult);
				iShellResult->Release();
				if (String.IsNullOrEmpty(path))
				{
					continue;
				}
				pathes[i] = path;
			}
			return pathes;
		}
		finally
		{
			if (shellItemArray != null)
			{
				shellItemArray->Release();
			}
			if (openDialog != null)
			{
				openDialog->Release();
			}
		}
	}

	/// <summary>
	/// IFileOpenDialog / IFileSaveDialog の共通表示処理
	/// </summary>
	/// <param name="fileDialog">IFileOpenDialog or IFileSaveDialog</param>
	/// <param name="filter">フィルター "説明|拡張子"（例："画像ファイル|*.jpg;*.jpeg"）、説明には拡張子を含めなくても自動的に OS 側で付与される</param>
	/// <param name="filterIndex">選択フィルター（1 オリジン）</param>
	/// <param name="options">オプション（複数選択など）</param>
	/// <param name="guid">最後にアクセスしたフォルダーなどが GUID 別に記憶される</param>
	/// <returns>true: 決定, false: キャンセル</returns>
	private unsafe Boolean ShowFileDialogCore(IFileDialog* fileDialog, String filter, ref UInt32 filterIndex, FILEOPENDIALOGOPTIONS options, Guid guid)
	{
		// IShellItem.GetDisplayName() が CoTaskMem なのでみんなそれに合わせる
		List<nint> coTaskMemories = [];
		void* shellInitial = null;

		// finally 用の try
		try
		{
			// GUID
			HRESULT result;
			fileDialog->SetClientGuid(guid);

			// オプション
			fileDialog->SetOptions(options);

			// フィルター
			(String[] filters, Boolean checkFilter) = ShowFileDialogCheckFilter(filter);
			if (checkFilter)
			{
				Int32 numFilters = filters.Length / 2;
				nint specs = Marshal.AllocCoTaskMem(Marshal.SizeOf<COMDLG_FILTERSPEC>() * numFilters);
				ReadOnlySpan<COMDLG_FILTERSPEC> specsSpan = new((void*)specs, numFilters);
				coTaskMemories.Add(specs);
				for (Int32 i = 0; i < numFilters; i++)
				{
					(PCWSTR namePcwstr, nint namePtr) = CreateCoMemPcwstr(filters[i * 2]);
					coTaskMemories.Add(namePtr);
					(PCWSTR specPcwstr, nint specPtr) = CreateCoMemPcwstr(filters[i * 2 + 1]);
					coTaskMemories.Add(specPtr);
					COMDLG_FILTERSPEC spec = new() { pszName = namePcwstr, pszSpec = specPcwstr };
					fixed (void* specsSpanPtr = specsSpan[i..])
					{
						Marshal.StructureToPtr(spec, (nint)specsSpanPtr, false);
					}
				}
				fileDialog->SetFileTypes(specsSpan);
				fileDialog->SetFileTypeIndex(filterIndex);
			}

			// ダイアログを表示
			HWND hWnd = (HWND)WindowNative.GetWindowHandle(App.MainWindow);
			result = fileDialog->Show(hWnd);
			if (result.Failed)
			{
				return false;
			}

			// フィルターインデックス書き戻し
			result = fileDialog->GetFileTypeIndex(out UInt32 filterIndexResult);
			if (result.Succeeded)
			{
				filterIndex = filterIndexResult;
			}
			return true;
		}
		finally
		{
			if (shellInitial != null)
			{
				((IShellItem*)shellInitial)->Release();
			}
			foreach (nint ptr in coTaskMemories)
			{
				Marshal.FreeCoTaskMem(ptr);
			}
		}
	}

	/// <summary>
	/// '|' で区切られたフィルター文字列（2 つでペア）を分解
	/// </summary>
	/// <param name="filter"></param>
	/// <returns>filters: フィルター（[n] が説明、[n+1] が拡張子）, check: ペアになっていたか</returns>
	private static (String[], Boolean) ShowFileDialogCheckFilter(String filter)
	{
		String[] filters = filter.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
		Boolean check = filters.Length > 0 && filters.Length % 2 == 0;
		return (filters, check);
	}

	/// <summary>
	/// 背後バッファが持続する PCWSTR を作成
	/// </summary>
	/// <param name="str"></param>
	/// <returns></returns>
	private static unsafe (PCWSTR, nint) CreateCoMemPcwstr(String str)
	{
		nint bufPtr = Marshal.StringToCoTaskMemUni(str);
		PCWSTR pcwstr = new((Char*)bufPtr);
		return (pcwstr, bufPtr);
	}

	/// <summary>
	/// IShellItem からパスを取得
	/// </summary>
	/// <param name="shellResult"></param>
	/// <returns></returns>
	private unsafe String? ShellItemToPath(IShellItem* shellResult)
	{
		PWSTR pathPwstr;
		HRESULT result = shellResult->GetDisplayName(SIGDN.SIGDN_FILESYSPATH, &pathPwstr);
		if (result.Failed)
		{
			return null;
		}
		String path = pathPwstr.ToString();
		Marshal.FreeCoTaskMem((nint)pathPwstr.Value);
		return path;
	}
}
