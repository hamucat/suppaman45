using Microsoft.Win32;
using Parzan.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Parzan
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        MainWindowVM vm;
        suppaman45.UserSettings userSettings;

        public MainWindow()
        {
            InitializeComponent();
            vm = new MainWindowVM();
            DataContext = vm;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //suppaman.exeの場所を確認する
            //settings.jsonでないのはない可能性もあって、その場合それを作るのはこっちの仕事なので
            var jsonPathManager = new JsonPathManager();
            var jsonPath = jsonPathManager.LoadPath();
            if (!File.Exists(jsonPath))
            {
                MessageBox.Show("suppaman45.exeが見つかりませんでした。次に開くダイアログでsuppaman45.exeのある場所を選択してください。","Parzan",MessageBoxButton.OK,MessageBoxImage.Information);
                
                var dialog = new OpenFileDialog();
                dialog.Title = "suppaman45.exeはどこにありますか？";
                dialog.Filter = "アプリケーション|*.exe";
                dialog.FileName = "suppaman45.exe";
                
                if (dialog.ShowDialog() == true)
                {
                    jsonPath = System.IO.Path.GetDirectoryName(dialog.FileName) + @"\settings.json";
                    jsonPathManager.SavePath(jsonPath);
                }else
                {
                    this.Close();
                    return;
                }
            }

            var settingManager = new suppaman45.SettingManager(jsonPath);
            
            //settings.jsonがなければつくる
            if (!File.Exists(jsonPath))
            {
                settingManager.SaveSettings(userSettings = new suppaman45.UserSettings());
            }

            userSettings = settingManager.LoadSettings();

            vm.ReadFileDir = userSettings.ReadFileDir;
            vm.ReadFileExtention = userSettings.ReadFileExtention;
            vm.ReadSheetName = userSettings.ReadSheetName;
            vm.Namedrange = userSettings.NamedRange;
            vm.ReadIgnoreThreshold = userSettings.ReadIgnoreThrethold;
            vm.WriteFilePath = userSettings.WriteFilepath;
            vm.WriteSheetName = userSettings.WriteSheetname;
            vm.WriteTableName = userSettings.WriteTableName;
            vm.ArchiveDirPath = userSettings.ArchiveDirPath;
            vm.ManageSheetName = userSettings.ManageSheetName;
            vm.UnprocessedDatesRangeName = userSettings.UnprocessedDatesRangeName;
        }

        //読み込みフォルダを開く
        private void ReadFileDir_ReferenceButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                vm.ReadFileDir = dialog.FileName + "\\";
            }
            
        }

        //書き込みファイルを開く
        private void WriteFilePath_ReferenceButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.IsFolderPicker= false;
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                vm.WriteFilePath = dialog.FileName;
            }
        }

        //アーカイブフォルダの選択
        private void ArchiveDirPath_ReferenceButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                vm.ArchiveDirPath = dialog.FileName + "\\";
            }
        }
    }
}
