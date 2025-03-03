using System.Management;
using MetadataExtractor;
using MetadataExtractor.Formats.Exif;
using Directory = System.IO.Directory;
using IniParser;
using IniParser.Model;

namespace ImageClassification
{
    public partial class frmMain : Form
    {
        private static string _targetRoot = "D:\\A_ImageClassification";
        private readonly ListBox _logBox;
        private readonly ProgressBar _progressBar;
        private readonly Label _progressLabel;
        private readonly ManualResetEventSlim _processingEvent = new(true);
        private ConfigManager _configManager = null;

        //private Dictionary<string,string>  _dic

        private string _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");

        // 目標存放目錄
        private int _totalFiles = 0;

        private int _processedFiles = 0;

        // 事件初始為可執行狀態
        private (string Make, string Model, string Date) lastJpgExifInfo = ("Unknown", "Unknown", "Unknown");

        // 設定程式視窗屬性與初始化控制項
        public frmMain()
        {
            InitializeComponent();
            _configManager = new ConfigManager("config.ini");
            _targetRoot = _configManager.GetOutputPath();
            this.Text = "SD 卡自動分類器";
            this.Width = 600;
            this.Height = 450;

            // 設定日誌顯示區域
            _logBox = new ListBox() { Dock = DockStyle.Top, Height = 300 };
            _progressBar = new ProgressBar() { Dock = DockStyle.Top, Height = 30 };
            _progressLabel = new Label() { Dock = DockStyle.Top, Height = 20, TextAlign = ContentAlignment.MiddleCenter };

            // 將控制項添加到表單中
            this.Controls.Add(_progressLabel);
            this.Controls.Add(_progressBar);
            this.Controls.Add(_logBox);

            // 啟動磁碟機監控
            StartDriveWatcher();
        }

        // 開始監聽磁碟插入事件
        private void StartDriveWatcher()
        {
            Task.Run(() =>
            {
                while (true) // 持續監聽
                {
                    using (var watcher = new ManagementEventWatcher(
                               new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2")))
                    {
                        watcher.EventArrived += new EventArrivedEventHandler(OnVolumeInserted);
                        watcher.Start();

                        Log("正在監聽磁碟機插入...");

                        // 讓監聽器持續等待事件
                        watcher.WaitForNextEvent();
                    }

                    Thread.Sleep(500); // 避免 CPU 過載
                }
            });
        }

        // 當新磁碟機插入時觸發
        private void OnVolumeInserted(object sender, EventArrivedEventArgs e)
        {
            var driveLetter = e.NewEvent["DriveName"]?.ToString();
            if (string.IsNullOrEmpty(driveLetter))
            {
                return;
            }

            Log($"偵測到新磁碟機: {driveLetter}");

            // **檢查是否正在處理其他磁碟，避免多次觸發**
            if (!_processingEvent.IsSet)
            {
                Log($"目前正在處理其他磁碟，跳過 {driveLetter}。");
                return;
            }

            // **阻擋後續的磁碟偵測**
            _processingEvent.Reset();

            // **開啟新執行緒處理磁碟內容 (避免UI卡住)**
            Task.Factory.StartNew(() => ProcessDrive(driveLetter), TaskCreationOptions.LongRunning);
        }

        // 處理偵測到的磁碟
        private async void ProcessDrive(string driveLetter)
        {
            var dcimPath = Path.Combine(driveLetter, "DCIM");
            if (!Directory.Exists(dcimPath))
            {
                Log("未找到 DCIM 資料夾，跳過。");
                _processingEvent.Set();
                return;
            }

            var files = Directory.GetFiles(dcimPath, "*.*", SearchOption.AllDirectories).ToList();
            _totalFiles = files.Count;
            _processedFiles = 0;
            UpdateProgress(0);

            var tasks = files.Select(file => Task.Run(() => ProcessFile(file, driveLetter))).ToList();

            try
            {
                await Task.WhenAll(tasks).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Log($"發生錯誤: {ex.Message}");
            }
            finally
            {
                Log("所有檔案處理完成！");
                UpdateProgress(100);
                OpenFolder(_targetRoot);
                _processingEvent.Set();
            }
        }

        private void ProcessFile(string filePath, string driveLetter)
        {
            if (filePath.ToLower().EndsWith(".jpg"))
            {
                ProcessJpeg(filePath, driveLetter);
            }
            else
            {
                if (!ProcessRaw(filePath, driveLetter))
                {
                    ProcessOtherFile(filePath, driveLetter);
                }
            }
        }

        private bool ProcessRaw(string filePath, string driveLetter)
        {
            try
            {
                var directories = ImageMetadataReader.ReadMetadata(filePath);
                var exifIfd0 = directories.OfType<ExifIfd0Directory>().FirstOrDefault();
                var exifSubIfd = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();

                var cameraMake = exifIfd0?.GetDescription(ExifIfd0Directory.TagMake)?.Trim() ?? "Unknown";
                var cameraModel = exifIfd0?.GetDescription(ExifIfd0Directory.TagModel)?.Trim() ?? "Unknown";
                if (cameraModel.StartsWith(cameraMake, StringComparison.OrdinalIgnoreCase))
                {
                    cameraModel = cameraModel.Substring(cameraMake.Length).Trim();
                }

                var dateTakenRaw = exifSubIfd?.GetDescription(ExifSubIfdDirectory.TagDateTime);
                var dateTaken = "Unknown";
                if (DateTime.TryParseExact(dateTakenRaw, "yyyy:MM:dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                {
                    dateTaken = dt.ToString("yyyy-MM-dd");
                }

                var destination = Path.Combine(_targetRoot, cameraMake, cameraModel, dateTaken);
                Directory.CreateDirectory(destination);
                File.Copy(filePath, Path.Combine(destination, Path.GetFileName(filePath)), true);
                Log($"分類 RAW: {filePath} -> {destination}");
                return true;
            }
            catch
            {
                return false;
            }
        }

        // 處理 JPG 檔案，並根據 EXIF 資訊分類
        private void ProcessJpeg(string filePath, string driveLetter)
        {
            try
            {
                if (!Directory.Exists(driveLetter))
                {
                    Log("磁碟已被移除，中止 JPEG 處理！");
                    return;
                }

                // 讀取圖片的 EXIF 資訊
                var directories = ImageMetadataReader.ReadMetadata(filePath);
                var exifSubIfd = directories.OfType<ExifSubIfdDirectory>().FirstOrDefault();
                var exifIfd0 = directories.OfType<ExifIfd0Directory>().FirstOrDefault();

                // 取得相機品牌與型號
                var cameraMake = exifIfd0?.GetDescription(ExifIfd0Directory.TagMake)?.Trim() ?? "Unknown";
                var cameraModel = exifIfd0?.GetDescription(ExifIfd0Directory.TagModel)?.Trim() ?? "Unknown";

                // **去除品牌名稱**
                if (cameraModel.StartsWith(cameraMake, StringComparison.OrdinalIgnoreCase))
                {
                    cameraModel = cameraModel.Substring(cameraMake.Length).Trim();
                }

                // 取得拍攝日期
                var dateTakenRaw = exifSubIfd?.GetDescription(ExifSubIfdDirectory.TagDateTime);
                // 可能部分機器沒有TagDateTime，用DitiTime再抓
                dateTakenRaw = exifSubIfd?.GetDescription(ExifSubIfdDirectory.TagDateTimeDigitized);
                var dateTaken = "Unknown";

                if (DateTime.TryParseExact(dateTakenRaw, "yyyy:MM:dd HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out DateTime dt))
                {
                    dateTaken = dt.ToString("yyyy-MM-dd");
                }

                // 更新最後處理的 JPG EXIF 資訊
                lastJpgExifInfo = (cameraMake, cameraModel, dateTaken);

                // 將檔案複製到對應的目標資料夾
                var destination = Path.Combine(_targetRoot, cameraMake, cameraModel, dateTaken);
                Directory.CreateDirectory(destination);
                File.Copy(filePath, Path.Combine(destination, Path.GetFileName(filePath)), true);

                Log($"分類: {filePath} -> {destination}");
            }
            catch (IOException)
            {
                Log($"錯誤: 檔案 {filePath} 無法存取，可能已移除。");
            }
            catch (Exception ex)
            {
                Log($"錯誤: {ex.Message}");
            }
            finally
            {
                UpdateProgressBar();
            }
        }

        // 處理其他檔案（非 JPG）
        private void ProcessOtherFile(string filePath, string driveLetter)
        {
            try
            {
                if (!Directory.Exists(driveLetter))
                {
                    Log("磁碟已被移除，中止其他檔案處理！");
                    return;
                }

                // 使用最後一個處理的 JPG 的 EXIF 資訊
                string cameraMake = lastJpgExifInfo.Make;
                string cameraModel = lastJpgExifInfo.Model;
                var dateTaken = File.GetCreationTime(filePath).ToString("yyyy-MM-dd");
                // 如果未處理過 JPG，則使用檔案建立日期
                if (cameraMake == "Unknown" && cameraModel == "Unknown")
                {
                    dateTaken = File.GetCreationTime(filePath).ToString("yyyy-MM-dd");
                }

                // 將檔案複製到對應的目標資料夾
                var destination = Path.Combine(_targetRoot, cameraMake, cameraModel, dateTaken);
                Directory.CreateDirectory(destination);
                File.Copy(filePath, Path.Combine(destination, Path.GetFileName(filePath)), true);

                Log($"分類: {filePath} -> {destination}");
            }
            catch (IOException)
            {
                Log($"錯誤: 檔案 {filePath} 無法存取，可能已移除。");
            }
            catch (Exception ex)
            {
                Log($"錯誤: {ex.Message}");
            }
            finally
            {
                UpdateProgressBar();
            }
        }

        // 開啟指定資料夾
        private void OpenFolder(string folderPath)
        {
            try
            {
                // 使用 explorer 開啟指定資料夾
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
            }
            catch (Exception ex)
            {
                Log($"無法開啟資料夾: {ex.Message}");
            }
        }

        // 更新進度條的狀態
        private void UpdateProgressBar()
        {
            Interlocked.Increment(ref _processedFiles);

            if (_totalFiles == 0) return;

            var progress = Math.Min(100, (int)((_processedFiles / (double)_totalFiles) * 100));

            // **減少 UI 更新頻率 (只更新每 10%)**
            if (progress % 10 == 0 || progress == 100)
            {
                UpdateProgress(progress);
            }
        }

        // 更新進度顯示
        private void UpdateProgress(int percentage)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() =>
                {
                    _progressBar.Value = percentage;
                    _progressLabel.Text = $"{_processedFiles} / {_totalFiles} 已處理";
                }));
            }
            else
            {
                _progressBar.Value = percentage;
                _progressLabel.Text = $"{_processedFiles} / {_totalFiles} 已處理";
            }
        }

        // 記錄日誌並顯示

        // 記錄日誌並顯示
        private void Log(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() =>
                {
                    _logBox.Items.Add(message);
                    // 自動捲動日誌區域
                    _logBox.SelectedIndex = _logBox.Items.Count - 1;
                }));
            }
            else
            {
                _logBox.Items.Add(message);
                // 自動捲動日誌區域
                _logBox.SelectedIndex = _logBox.Items.Count - 1;
            }
        }
    }
}

public class ConfigManager
{
    private readonly string _configFile;
    private readonly FileIniDataParser _parser;
    private IniData _data;

    public ConfigManager(string configFile)
    {
        _configFile = configFile;
        _parser = new FileIniDataParser();

        if (!File.Exists(_configFile))
        {
            // 建立預設設定
            _data = new IniData();
            _data["Settings"]["OutputPath"] = "D:\\A_ImageClassification";
            Save();
        }
        else
        {
            // 讀取設定檔
            _data = _parser.ReadFile(_configFile);
        }
    }

    public string GetOutputPath()
    {
        return _data["Settings"]["OutputPath"];
    }

    public void SetOutputPath(string path)
    {
        _data["Settings"]["OutputPath"] = path;
        Save();
    }

    public void Save()
    {
        _parser.WriteFile(_configFile, _data);
    }
}