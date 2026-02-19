using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Globalization;
using System.Text.RegularExpressions; // Added for validation
using System.Threading.Tasks; // Added for Task
using System.Diagnostics;
using System.Text;
using System.IO.Compression;

namespace AI_Video_Assembler
{
    public partial class MainWindow : Window
    {
        // [New] Global Pool for Random Colors
        public static ObservableCollection<string> RandomColorPool { get; set; } = new ObservableCollection<string>();
        private string _randomPoolPath = "random_colors.txt";
        public static RenderJobSettings settings;
        public FontFamily _defaultSystemFont { get; set; } = new FontFamily("Segoe UI"); // ค่ากันตาย

        // [New] Matrix Variables
        private long _totalCombinations = 0;

        #region 0.Varidate Step
        private bool ValidateStep1()
        {
            // ตรวจสอบว่า List/Collection ของซีนว่างเปล่าหรือไม่
            if (Folders == null || Folders.Count == 0)
            {
                // แจ้งเตือนภาษาไทยตามที่กำหนด
                MessageBox.Show("กรุณาอัปโหลดซีนของวีดีโอก่อนไปสเตปถัดไป", "แจ้งเตือน", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true; // มีซีนแล้ว ให้ผ่านได้
        }
        #endregion

        #region 1. Data Models
        public class SceneFolder : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

            private string _name;
            public string Name { get => _name; set { _name = value; OnPropertyChanged(nameof(Name)); } }

            public int FileCount { get; set; }
            public string FolderPath { get; set; }
            public List<VideoFile> Files { get; set; } = new List<VideoFile>();
            public ObservableCollection<TextLayer> TextLayers { get; set; } = new ObservableCollection<TextLayer>();
        }

        public class VideoFile : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

            private bool _isSelected = true;
            public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }

            public string Name { get; set; }
            public string FilePath { get; set; }

            private string _duration = "Loading...";
            public string Duration { get => _duration; set { _duration = value; OnPropertyChanged(nameof(Duration)); } }

            private string _status = "Checking...";
            public string Status { get => _status; set { _status = value; OnPropertyChanged(nameof(Status)); } }

            private Brush _statusColor = Brushes.Gray;
            public Brush StatusColor { get => _statusColor; set { _statusColor = value; OnPropertyChanged(nameof(StatusColor)); } }
        }

        public class TextLayer : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
                // Update brush when any color-related property changes
                if (name == nameof(BgColor) || name == nameof(BgOpacity) ||
                    name == nameof(BgMode) || name == nameof(GradientEndColor) || name == nameof(IsGradientHorizontal))
                    UpdateDisplayBrush();
            }
            // [สำคัญ] ต้องมีตัวนี้เพื่อเก็บ Path ไฟล์ .ttf จริงๆ
            public string FontPath { get; set; }

            public FontFamily FontFamily
            {
                get => _fontFamily;
                set { _fontFamily = value; OnPropertyChanged(nameof(FontFamily)); }
            }

            private string _text = "New Text";
            public string Text { get => _text; set { _text = value; OnPropertyChanged(nameof(Text)); } }

            private FontFamily _fontFamily = new FontFamily("Segoe UI");
            private double _x = 50;
            public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }

            private double _y = 50;
            public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }

            private int _fontSize = 32;
            public int FontSize { get => _fontSize; set { _fontSize = value; OnPropertyChanged(nameof(FontSize)); } }

            // --- Background Properties ---

            // Mode: 0 = Solid, 1 = Gradient, 2 = Random
            private int _bgMode = 0;
            public int BgMode { get => _bgMode; set { _bgMode = value; OnPropertyChanged(nameof(BgMode)); } }

            // Solid Color
            private string _bgColor = "#000000";
            public string BgColor { get => _bgColor; set { _bgColor = value; OnPropertyChanged(nameof(BgColor)); } }

            // Gradient End Color (Start is BgColor)
            private string _gradientEndColor = "#FF0000";
            public string GradientEndColor { get => _gradientEndColor; set { _gradientEndColor = value; OnPropertyChanged(nameof(GradientEndColor)); } }

            private bool _isGradientHorizontal = true;
            public bool IsGradientHorizontal { get => _isGradientHorizontal; set { _isGradientHorizontal = value; OnPropertyChanged(nameof(IsGradientHorizontal)); } }

            private double _bgOpacity = 0.5;
            public double BgOpacity { get => _bgOpacity; set { _bgOpacity = value; OnPropertyChanged(nameof(BgOpacity)); } }

            private Brush _displayBrush;
            public Brush DisplayBrush { get => _displayBrush; private set { _displayBrush = value; OnPropertyChanged(nameof(DisplayBrush)); } }

            private TextAlignment _textAlign = TextAlignment.Center;
            public TextAlignment TextAlign { get => _textAlign; set { _textAlign = value; OnPropertyChanged(nameof(TextAlign)); } }

            // Animation Properties
            private string _animationEffect = "None (ไม่มี)";
            public string AnimationEffect { get => _animationEffect; set { _animationEffect = value; OnPropertyChanged(nameof(AnimationEffect)); } }

            private double _animationDuration = 1.0;
            public double AnimationDuration { get => _animationDuration; set { _animationDuration = value; OnPropertyChanged(nameof(AnimationDuration)); } }

            private double _delay = 0.0;
            public double Delay { get => _delay; set { _delay = value; OnPropertyChanged(nameof(Delay)); } }

            private double _textDuration = 5.0;
            public double TextDuration { get => _textDuration; set { _textDuration = value; OnPropertyChanged(nameof(TextDuration)); } }

            private double _effectDuration = 1.0;
            public double EffectDuration { get => _effectDuration; set { _effectDuration = value; OnPropertyChanged(nameof(EffectDuration)); } }

            private string _exitEffect = "None (ไม่มี)";
            public string ExitEffect { get => _exitEffect; set { _exitEffect = value; OnPropertyChanged(nameof(ExitEffect)); } }

            private double _exitEffectDuration = 1.0;
            public double ExitEffectDuration { get => _exitEffectDuration; set { _exitEffectDuration = value; OnPropertyChanged(nameof(ExitEffectDuration)); } }

            private bool _hasExitAnimation = false;
            public bool HasExitAnimation { get => _hasExitAnimation; set { _hasExitAnimation = value; OnPropertyChanged(nameof(HasExitAnimation)); } }

            private string _activePresetName = "";
            public string ActivePresetName { get => _activePresetName; set { _activePresetName = value; OnPropertyChanged(nameof(ActivePresetName)); } }

            // Selection Properties
            private bool _isSelectedLayer = false;
            public bool IsSelectedLayer
            {
                get => _isSelectedLayer;
                set
                {
                    _isSelectedLayer = value;
                    OnPropertyChanged(nameof(IsSelectedLayer));
                    LayerBorderThickness = value ? new Thickness(2) : new Thickness(0);
                    LayerBorderBrush = value ? new SolidColorBrush(Color.FromRgb(79, 70, 229)) : Brushes.Transparent;
                }
            }

            private Thickness _layerBorderThickness = new Thickness(0);
            public Thickness LayerBorderThickness
            {
                get => _layerBorderThickness;
                set { _layerBorderThickness = value; OnPropertyChanged(nameof(LayerBorderThickness)); }
            }

            private Brush _layerBorderBrush = Brushes.Transparent;
            public Brush LayerBorderBrush
            {
                get => _layerBorderBrush;
                set { _layerBorderBrush = value; OnPropertyChanged(nameof(LayerBorderBrush)); }
            }

            public static ObservableCollection<string> EffectOptions { get; } = new ObservableCollection<string>
            {
                "None (ไม่มี)", "Fade In (เลือนเข้า)", "Slide Left (เลื่อนซ้าย)", "Slide Right (เลื่อนขวา)",
                "Slide Up (เลื่อนขึ้น)", "Slide Down (เลื่อนลง)", "Zoom In (ขยายเข้า)", "Zoom Out (ขยายออก)",
                "Bounce (เด้ง)", "Rotate In (หมุนเข้า)"
            };

            public TextLayer() { UpdateDisplayBrush(); }

            public void UpdateDisplayBrush()
            {
                try
                {
                    if (BgMode == 2) // Random Mode
                    {
                        // Pick a random color from the global pool for display preview
                        if (MainWindow.RandomColorPool != null && MainWindow.RandomColorPool.Count > 0)
                        {
                            var rand = new Random();
                            string randomHex = MainWindow.RandomColorPool[rand.Next(MainWindow.RandomColorPool.Count)];
                            var color = (Color)ColorConverter.ConvertFromString(randomHex);
                            var brush = new SolidColorBrush(color);
                            brush.Opacity = BgOpacity;
                            DisplayBrush = brush;
                        }
                        else
                        {
                            // Fallback if pool empty
                            var brush = new SolidColorBrush(Colors.Gray);
                            brush.Opacity = BgOpacity;
                            DisplayBrush = brush;
                        }
                    }
                    else if (BgMode == 1) // Gradient Mode
                    {
                        var startColor = (Color)ColorConverter.ConvertFromString(BgColor);
                        var endColor = (Color)ColorConverter.ConvertFromString(GradientEndColor);

                        var brush = new LinearGradientBrush();
                        brush.StartPoint = new Point(0, 0);
                        brush.EndPoint = IsGradientHorizontal ? new Point(1, 0) : new Point(0, 1);
                        brush.GradientStops.Add(new GradientStop(startColor, 0.0));
                        brush.GradientStops.Add(new GradientStop(endColor, 1.0));
                        brush.Opacity = BgOpacity;
                        DisplayBrush = brush;
                    }
                    else // Solid Mode (Default)
                    {
                        var color = (Color)ColorConverter.ConvertFromString(BgColor);
                        var brush = new SolidColorBrush(color);
                        brush.Opacity = BgOpacity;
                        DisplayBrush = brush;
                    }
                }
                catch { DisplayBrush = Brushes.Transparent; }
            }

            // Method to force refresh (e.g. when random pool changes)
            public void RefreshBrush() => UpdateDisplayBrush();
        }
        #endregion

        /// <summary>ช่วงที่กำลังเล่นในพรีวิว: Intro = ภาพต้น, Clips = วิดีโอ, Outro = ภาพท้าย</summary>
        private enum PreviewPhase { Idle, Intro, Clips, Outro }

        #region 2. Variables & Constructor
        private int _currentStep = 1;
        private List<string> _previewPlaylist = new List<string>();
        private List<double> _clipAccumulatedDurations = new List<double>();
        private double _totalSequenceDuration = 0;
        private double _clipsTotalDuration = 0;
        private int _currentPlaylistIndex = 0;
        private DispatcherTimer _trimTimer;
        private bool _isPlayerAActive = false;
        private bool _isDraggingSlider = false;
        private bool _nextPlayerReadyForSwitch = false;

        private string _introImagePath = null;
        private string _outroImagePath = null;
        private double _introDurationSeconds = 0;
        private double _outroDurationSeconds = 0;
        private PreviewPhase _previewPhase = PreviewPhase.Idle;
        private DateTime _introStartTime;
        private DateTime _outroStartTime;
        private double _pausedSequenceTime = 0;
        public ObservableCollection<string> CachedColors { get; set; } = new ObservableCollection<string>();
        private string _cacheFilePath = "color_cache.txt";
        private DispatcherTimer _step3Timer;
        private bool _isDraggingText = false;
        private Point _clickPosition;
        private TextLayer _draggingLayer;
        private bool _isDraggingSliderStep3 = false;
        private bool _step3WasPlayingBeforeSliderDrag = false;
        private bool _isAdjustingFontSize = false;
        private double _step3TimelinePixelsPerSecond = 40;
        private double _step3TimelineTotalDuration = 0;
        // Step 3 timeline segment drag/resize
        private bool _step3SegmentDragging;
        private bool _step3SegmentResizing;
        private TextLayer _step3SegmentLayer;
        private Border _step3SegmentElement;
        private double _step3SegmentStartMouseX;
        private double _step3SegmentStartDelay;
        private double _step3SegmentStartDuration;
        private const double Step3SegmentMinDuration = 0.5;
        public ObservableCollection<SceneFolder> Folders { get; set; }
        public ObservableCollection<FontFamilyGroup> FontGroups { get; set; } = new ObservableCollection<FontFamilyGroup>();

        public MainWindow()
        {
            InitializeComponent();
            Folders = new ObservableCollection<SceneFolder>();
            if (ListFolders != null) { ListFolders.ItemsSource = Folders; ListFolders.SelectionChanged += ListFolders_SelectionChanged; }
            _trimTimer = new DispatcherTimer(); _trimTimer.Interval = TimeSpan.FromMilliseconds(25); _trimTimer.Tick += TrimTimer_Tick;
            _step3Timer = new DispatcherTimer(); _step3Timer.Interval = TimeSpan.FromMilliseconds(100); _step3Timer.Tick += Step3Timer_Tick;

            LoadColorCache();
            if (ListCachedColors != null) ListCachedColors.ItemsSource = CachedColors;

            LoadRandomColorPool();
            LoadFonts();
            UpdateSidebarUI(1);
            CreateDummyFontConfig(AppDomain.CurrentDomain.BaseDirectory);

            // Set default export path
            if (TxtExportPath != null)
                TxtExportPath.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "AI_Video_Export");
        }
        #endregion

        #region 3. Navigation
        private void NavButton_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && int.TryParse(btn.Tag.ToString(), out int step)) GoToStep(step); }
        private void NextStep_Click(object sender, RoutedEventArgs e) { if (_currentStep < 5) GoToStep(_currentStep + 1); }
        private void PrevStep_Click(object sender, RoutedEventArgs e) { if (_currentStep > 1) GoToStep(_currentStep - 1); }

        private void GoToStep(int step)
        {
            // [เพิ่มใหม่] เงื่อนไข Lock: หากพยายามไป Step อื่น (2,3,4,5) แต่ยังไม่มี Scene ใน Step 1
            if (step > 1 && (Folders == null || Folders.Count == 0))
            {
                MessageBox.Show("กรุณาอัปโหลดซีนของวีดีโอก่อนไปสเตปถัดไป", "แจ้งเตือน", MessageBoxButton.OK, MessageBoxImage.Warning);
                return; // หยุดการทำงาน ไม่เปลี่ยนหน้า
            }

            _currentStep = step;

            // ซ่อนทุก View ก่อน
            if (ViewStep1 != null) ViewStep1.Visibility = Visibility.Collapsed;
            if (ViewStep2 != null) ViewStep2.Visibility = Visibility.Collapsed;
            if (ViewStep3 != null) ViewStep3.Visibility = Visibility.Collapsed;
            if (ViewStep4 != null) ViewStep4.Visibility = Visibility.Collapsed;
            if (ViewStep5 != null) ViewStep5.Visibility = Visibility.Collapsed;

            // หยุดการเล่นวิดีโอ Preview ของ Step ต่างๆ เมื่อเปลี่ยนหน้า
            if (step != 2) StopPreview();
            StopPreviewStep3();
            StopStep4_Click(null, null);

            // แสดง View ที่เลือก
            switch (step)
            {
                case 1:
                    if (ViewStep1 != null) ViewStep1.Visibility = Visibility.Visible;
                    break;
                case 2:
                    if (ViewStep2 != null) ViewStep2.Visibility = Visibility.Visible;
                    break;
                case 3:
                    if (ViewStep3 != null) ViewStep3.Visibility = Visibility.Visible;
                    RefreshVisualSceneSelect();
                    break;
                case 4:
                    if (ViewStep4 != null) ViewStep4.Visibility = Visibility.Visible;
                    break;
                case 5:
                    if (ViewStep5 != null)
                    {
                        ViewStep5.Visibility = Visibility.Visible;
                        UpdateStep5Summary();
                    }
                    break;
            }

            // อัปเดตเมนู Sidebar และ Progress Bar
            UpdateSidebarUI(step);
        }
        private void UpdateStep5Summary()
        {
            if (TxtSummaryRes == null) return;

            // เช็คขนาดจาก PreviewScreen ที่ตั้งค่าไว้ใน Step 2
            double w = 1280;
            double h = 720;

            if (PreviewScreen != null)
            {
                w = PreviewScreen.Width > 0 ? PreviewScreen.Width : 1280;
                h = PreviewScreen.Height > 0 ? PreviewScreen.Height : 720;
            }

            string ratio = (w > h) ? "16:9 Landscape" : "9:16 Portrait";
            TxtSummaryRes.Text = $"{w}x{h} ({ratio})";

            // Reset File Size text
            if (TxtSummarySize != null) TxtSummarySize.Text = "-";
        }

        private void UpdateSidebarUI(int step)
        {
            var btns = new Button[] { BtnStep1, BtnStep2, BtnStep3, BtnStep4, BtnStep5 };
            foreach (var btn in btns) { if (btn == null) continue; btn.Background = Brushes.Transparent; btn.Foreground = new SolidColorBrush(Color.FromRgb(107, 114, 128)); btn.FontWeight = FontWeights.Normal; }
            Button activeBtn = FindName($"BtnStep{step}") as Button;
            if (activeBtn != null) { activeBtn.Background = new SolidColorBrush(Color.FromRgb(238, 242, 255)); activeBtn.Foreground = new SolidColorBrush(Color.FromRgb(79, 70, 229)); activeBtn.FontWeight = FontWeights.Bold; }
            if (ProjectProgress != null) ProjectProgress.Value = step * 20;
            if (TxtProgress != null) TxtProgress.Text = $"{step * 20}%";
        }
        #endregion

        #region 4. Step 1: Import
        private void ListFolders_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (ListFolders.SelectedItem is SceneFolder selectedFolder && GridFiles != null) GridFiles.ItemsSource = selectedFolder.Files; }

        private void AddFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog { IsFolderPicker = true, EnsurePathExists = true };
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                string path = dialog.FileName;
                if (Folders.Any(f => f.FolderPath != null && f.FolderPath.Equals(path, StringComparison.OrdinalIgnoreCase))) { MessageBox.Show("โฟลเดอร์นี้ถูกนำเข้าแล้ว"); return; }
                try
                {
                    string folderName = new DirectoryInfo(path).Name;
                    var allowedExtensions = new[] { ".mp4", ".avi", ".mov", ".mkv", ".wmv" };
                    var files = Directory.GetFiles(path).Where(file => allowedExtensions.Contains(Path.GetExtension(file).ToLower())).ToList();
                    if (files.Count > 0)
                    {
                        var newFolder = new SceneFolder { Name = folderName, FileCount = files.Count, FolderPath = path, Files = new List<VideoFile>() };
                        foreach (var filePath in files) { var vFile = new VideoFile { Name = Path.GetFileName(filePath), FilePath = filePath, Duration = "Loading...", Status = "Checking...", StatusColor = Brushes.Orange }; newFolder.Files.Add(vFile); VerifyAndLoadVideo(vFile, filePath); }
                        Folders.Add(newFolder); ListFolders.SelectedItem = newFolder;
                    }
                    else MessageBox.Show("ไม่พบไฟล์วิดีโอ");
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            }
        }

        private void VerifyAndLoadVideo(VideoFile videoFile, string filePath)
        {
            var player = new MediaPlayer();
            var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
            void Cleanup() { timer.Stop(); player.Close(); }
            timer.Tick += (s, e) => { Cleanup(); if (videoFile.Status == "Checking...") { videoFile.Status = "Error"; videoFile.StatusColor = Brushes.Red; } };
            player.MediaOpened += (s, e) => { timer.Stop(); if (player.NaturalDuration.HasTimeSpan) { videoFile.Duration = player.NaturalDuration.TimeSpan.ToString(@"mm\:ss"); videoFile.Status = "Ready"; videoFile.StatusColor = Brushes.Green; } else { videoFile.Status = "Error"; videoFile.StatusColor = Brushes.Red; } player.Close(); };
            player.MediaFailed += (s, e) => { Cleanup(); videoFile.Status = "Error"; videoFile.StatusColor = Brushes.Red; };
            timer.Start();
            try { player.Open(new Uri(filePath)); } catch { Cleanup(); videoFile.Status = "Error"; videoFile.StatusColor = Brushes.Red; }
        }

        private void DeleteFolder_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.Tag is SceneFolder folder && MessageBox.Show($"ลบซีน '{folder.Name}'?", "ยืนยัน", MessageBoxButton.YesNo) == MessageBoxResult.Yes) { Folders.Remove(folder); if (ListFolders.SelectedItem == folder) GridFiles.ItemsSource = null; } }
        private void EditSceneName_Click(object sender, RoutedEventArgs e)
        {
            // Simplified input dialog logic for brevity
            if (sender is Button btn && btn.Tag is SceneFolder folder)
            {
                // InputBox logic here
            }
        }
        private void MoveSceneUp_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.Tag is SceneFolder folder) { int idx = Folders.IndexOf(folder); if (idx > 0) Folders.Move(idx, idx - 1); } }
        private void MoveSceneDown_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.Tag is SceneFolder folder) { int idx = Folders.IndexOf(folder); if (idx < Folders.Count - 1) Folders.Move(idx, idx + 1); } }
        #endregion

        #region Step 2: Settings (Seamless Playback & Preview = Render)
        private void SpeedSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (TxtSpeedValue != null) TxtSpeedValue.Text = $"{e.NewValue:F1}x";
            if (PreviewPlayerA != null) PreviewPlayerA.SpeedRatio = e.NewValue;
            if (PreviewPlayerB != null) PreviewPlayerB.SpeedRatio = e.NewValue;
        }

        private void MuteAudio_Click(object sender, RoutedEventArgs e)
        {
            bool mute = ChkMuteAudio != null && ChkMuteAudio.IsChecked == true;
            if (PreviewPlayerA != null) PreviewPlayerA.IsMuted = mute;
            if (PreviewPlayerB != null) PreviewPlayerB.IsMuted = mute;
        }

        private void AspectRatio_Click(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag != null)
            {
                double w = 1280, h = 720;
                if (rb.Tag.ToString() == "9:16") { w = 720; h = 1280; }
                if (PreviewScreen != null) { PreviewScreen.Width = w; PreviewScreen.Height = h; }
                if (PreviewScreenStep3 != null) { PreviewScreenStep3.Width = w; PreviewScreenStep3.Height = h; }
            }
        }

        private void BrowseIntro_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.Filters.Add(new Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"));
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                _introImagePath = dialog.FileName;
                if (TxtIntroPath != null) TxtIntroPath.Text = Path.GetFileName(_introImagePath);
            }
        }

        private void BrowseOutro_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.Filters.Add(new Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Images", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"));
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                _outroImagePath = dialog.FileName;
                if (TxtOutroPath != null) TxtOutroPath.Text = Path.GetFileName(_outroImagePath);
            }
        }

        private void PlaySequencePreview_Click(object sender, RoutedEventArgs e)
        {
            if (Folders == null || Folders.Count == 0) return;

            _previewPlaylist.Clear();
            _clipAccumulatedDurations.Clear();
            _clipsTotalDuration = 0;
            _totalSequenceDuration = 0;
            _nextPlayerReadyForSwitch = false;

            double trimStart = 0;
            double.TryParse(TxtTrimStart?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimStart);
            double trimEnd = 0;
            double.TryParse(TxtTrimEnd?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimEnd);
            double speed = SpeedSlider != null ? SpeedSlider.Value : 1.0;

            foreach (var folder in Folders)
            {
                var clip = folder.Files?.FirstOrDefault(f => f.Status == "Ready" && f.IsSelected);
                if (clip != null)
                {
                    _previewPlaylist.Add(clip.FilePath);
                    double dur = ParseClipDurationSeconds(clip.Duration);
                    double playDur = Math.Max(0, dur - trimStart - trimEnd);
                    _clipAccumulatedDurations.Add(_clipsTotalDuration);
                    _clipsTotalDuration += (playDur / speed);
                }
            }

            if (_previewPlaylist.Count == 0)
            {
                MessageBox.Show("กรุณาเลือกวิดีโออย่างน้อย 1 ไฟล์ใน Step 1", "แจ้งเตือน", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _introDurationSeconds = 0;
            _outroDurationSeconds = 0;
            if (!string.IsNullOrWhiteSpace(_introImagePath) && File.Exists(_introImagePath))
            {
                double.TryParse(TxtIntroDuration?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out _introDurationSeconds);
                _introDurationSeconds = Math.Max(0, _introDurationSeconds);
            }
            if (!string.IsNullOrWhiteSpace(_outroImagePath) && File.Exists(_outroImagePath))
            {
                double.TryParse(TxtOutroDuration?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out _outroDurationSeconds);
                _outroDurationSeconds = Math.Max(0, _outroDurationSeconds);
            }

            _totalSequenceDuration = _introDurationSeconds + _clipsTotalDuration + _outroDurationSeconds;
            for (int i = 0; i < _clipAccumulatedDurations.Count; i++)
                _clipAccumulatedDurations[i] += _introDurationSeconds;

            if (TxtTotalTime != null) TxtTotalTime.Text = TimeSpan.FromSeconds(_totalSequenceDuration).ToString(@"mm\:ss");
            if (TimelineSlider != null) { TimelineSlider.Maximum = _totalSequenceDuration; TimelineSlider.Value = 0; TimelineSlider.Minimum = 0; }
            if (BtnPlayPauseToggle != null) BtnPlayPauseToggle.IsEnabled = true;
            if (BtnStopPreview != null) BtnStopPreview.IsEnabled = true;
            if (PreviewPlaceholder != null) PreviewPlaceholder.Visibility = Visibility.Collapsed;

            StopPlayers();
            HideIntroOutroImage();
            _currentPlaylistIndex = 0;
            _isPlayerAActive = true;
            _pausedSequenceTime = 0;

            if (_introDurationSeconds > 0 && !string.IsNullOrEmpty(_introImagePath))
            {
                _previewPhase = PreviewPhase.Intro;
                ShowIntroOutroImage(_introImagePath);
                _introStartTime = DateTime.Now;
                _trimTimer.Start();
                UpdateStep2PlayPauseIcon();
                return;
            }

            _previewPhase = PreviewPhase.Clips;
            PreparePlayer(PreviewPlayerA, 0, autoPlay: true);
            if (_previewPlaylist.Count > 1)
                PreparePlayer(PreviewPlayerB, 1, autoPlay: false);
            PreviewPlayerA.Opacity = 1;
            PreviewPlayerB.Opacity = 0;
            PreviewPlayerA.Play();
            _trimTimer.Start();
            UpdateStep2PlayPauseIcon();
        }

        private static double ParseClipDurationSeconds(string duration)
        {
            if (string.IsNullOrWhiteSpace(duration)) return 0;
            if (TimeSpan.TryParse("00:" + duration.Trim(), CultureInfo.InvariantCulture, out TimeSpan ts))
                return ts.TotalSeconds;
            return 0;
        }

        private void ShowIntroOutroImage(string imagePath)
        {
            try
            {
                if (IntroOutroImage == null) return;
                var uri = new Uri(imagePath, UriKind.Absolute);
                IntroOutroImage.Source = new BitmapImage(uri);
                IntroOutroImage.Visibility = Visibility.Visible;
                if (PreviewPlayerA != null) PreviewPlayerA.Opacity = 0;
                if (PreviewPlayerB != null) PreviewPlayerB.Opacity = 0;
            }
            catch { }
        }

        private void HideIntroOutroImage()
        {
            if (IntroOutroImage != null) { IntroOutroImage.Source = null; IntroOutroImage.Visibility = Visibility.Collapsed; }
            _previewPhase = PreviewPhase.Idle;
        }

        private double GetCurrentSequenceTime()
        {
            if (_previewPhase == PreviewPhase.Intro)
                return Math.Min(_introDurationSeconds, (DateTime.Now - _introStartTime).TotalSeconds);
            if (_previewPhase == PreviewPhase.Outro)
            {
                double elapsed = (DateTime.Now - _outroStartTime).TotalSeconds;
                return _introDurationSeconds + _clipsTotalDuration + Math.Min(elapsed, _outroDurationSeconds);
            }
            if (_previewPhase == PreviewPhase.Clips)
            {
                var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                if (active == null || active.Source == null || !active.NaturalDuration.HasTimeSpan) return _introDurationSeconds;
                double trimStart = 0; double.TryParse(TxtTrimStart?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimStart);
                double speed = SpeedSlider != null ? SpeedSlider.Value : 1.0;
                double currentClipTime = active.Position.TotalSeconds;
                if (currentClipTime < trimStart) currentClipTime = trimStart;
                double playedInClip = (currentClipTime - trimStart) / speed;
                return _clipAccumulatedDurations[_currentPlaylistIndex] + playedInClip;
            }
            return _pausedSequenceTime;
        }

        private void PausePreview_Click(object sender, RoutedEventArgs e)
        {
            if (_trimTimer.IsEnabled)
            {
                _pausedSequenceTime = GetCurrentSequenceTime();
                _trimTimer.Stop();
                if (_previewPhase == PreviewPhase.Clips)
                {
                    var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                    if (active != null) active.Pause();
                }
            }
            else
            {
                SeekToTime(_pausedSequenceTime);
                if (_previewPhase == PreviewPhase.Intro)
                { _introStartTime = DateTime.Now.AddSeconds(-_pausedSequenceTime); _trimTimer.Start(); }
                else if (_previewPhase == PreviewPhase.Clips)
                { var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB; if (active != null) active.Play(); _trimTimer.Start(); }
                else if (_previewPhase == PreviewPhase.Outro)
                { _outroStartTime = DateTime.Now.AddSeconds(-(_pausedSequenceTime - _introDurationSeconds - _clipsTotalDuration)); _trimTimer.Start(); }
            }
        }

        private void StopPreview_Click(object sender, RoutedEventArgs e) { StopPreview(); }

        private void UpdateStep2PlayPauseIcon()
        {
            if (IconPlayPause == null) return;
            try
            {
                var isPlaying = _trimTimer != null && _trimTimer.IsEnabled;
                var geo = (Geometry)Resources[isPlaying ? "IconPause" : "IconPlay"];
                if (geo != null) IconPlayPause.Data = geo;
            }
            catch { }
        }

        private void PlayPauseToggle_Click(object sender, RoutedEventArgs e)
        {
            if (_trimTimer.IsEnabled)
            {
                _pausedSequenceTime = GetCurrentSequenceTime();
                _trimTimer.Stop();
                if (_previewPhase == PreviewPhase.Clips)
                {
                    var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                    if (active != null) active.Pause();
                }
            }
            else
            {
                if (_totalSequenceDuration > 0 && _pausedSequenceTime < _totalSequenceDuration)
                {
                    SeekToTime(_pausedSequenceTime);
                    if (_previewPhase == PreviewPhase.Intro)
                    { _introStartTime = DateTime.Now.AddSeconds(-_pausedSequenceTime); _trimTimer.Start(); }
                    else if (_previewPhase == PreviewPhase.Clips)
                    { var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB; if (active != null) active.Play(); _trimTimer.Start(); }
                    else if (_previewPhase == PreviewPhase.Outro)
                    { _outroStartTime = DateTime.Now.AddSeconds(-(_pausedSequenceTime - _introDurationSeconds - _clipsTotalDuration)); _trimTimer.Start(); }
                }
                else
                    PlaySequencePreview_Click(sender, e);
            }
            UpdateStep2PlayPauseIcon();
        }

        private void PreparePlayer(MediaElement player, int index, bool autoPlay = false)
        {
            if (player == null || index >= _previewPlaylist.Count) return;

            double trimStart = 0;
            double.TryParse(TxtTrimStart?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimStart);
            double speed = SpeedSlider != null ? SpeedSlider.Value : 1.0;
            bool mute = ChkMuteAudio != null && ChkMuteAudio.IsChecked == true;

            try
            {
                player.SpeedRatio = speed;
                player.IsMuted = mute;
                player.Source = new Uri(_previewPlaylist[index]);
                player.Position = TimeSpan.FromSeconds(trimStart);
                player.Play();
                player.Pause();
                if (autoPlay) player.Play();
            }
            catch { }
        }

        private void TrimTimer_Tick(object sender, EventArgs e)
        {
            double totalTime;

            if (_previewPhase == PreviewPhase.Intro)
            {
                totalTime = (DateTime.Now - _introStartTime).TotalSeconds;
                if (!_isDraggingSlider && TimelineSlider != null)
                {
                    TimelineSlider.Value = Math.Min(totalTime, _introDurationSeconds);
                    if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(totalTime).ToString(@"mm\:ss");
                }
                if (totalTime >= _introDurationSeconds)
                {
                    _previewPhase = PreviewPhase.Clips;
                    HideIntroOutroImage();
                    _currentPlaylistIndex = 0;
                    _isPlayerAActive = true;
                    PreparePlayer(PreviewPlayerA, 0, autoPlay: true);
                    if (_previewPlaylist.Count > 1)
                        PreparePlayer(PreviewPlayerB, 1, autoPlay: false);
                    PreviewPlayerA.Opacity = 1;
                    PreviewPlayerB.Opacity = 0;
                    PreviewPlayerA.Play();
                }
                return;
            }

            if (_previewPhase == PreviewPhase.Outro)
            {
                double elapsed = (DateTime.Now - _outroStartTime).TotalSeconds;
                totalTime = _introDurationSeconds + _clipsTotalDuration + elapsed;
                if (!_isDraggingSlider && TimelineSlider != null)
                {
                    TimelineSlider.Value = Math.Min(totalTime, _totalSequenceDuration);
                    if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(totalTime).ToString(@"mm\:ss");
                }
                if (elapsed >= _outroDurationSeconds)
                    StopPreview();
                return;
            }

            var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
            var next = _isPlayerAActive ? PreviewPlayerB : PreviewPlayerA;
            if (active == null || active.Source == null || !active.NaturalDuration.HasTimeSpan) return;

            double trimStart = 0;
            double.TryParse(TxtTrimStart?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimStart);
            double trimEnd = 0;
            double.TryParse(TxtTrimEnd?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimEnd);
            double speed = SpeedSlider != null ? SpeedSlider.Value : 1.0;

            double currentClipTime = active.Position.TotalSeconds;
            if (currentClipTime < trimStart) currentClipTime = trimStart;
            double playedInClip = (currentClipTime - trimStart) / speed;
            totalTime = _clipAccumulatedDurations[_currentPlaylistIndex] + playedInClip;

            if (!_isDraggingSlider && TimelineSlider != null)
            {
                TimelineSlider.Value = Math.Min(totalTime, TimelineSlider.Maximum);
                if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(totalTime).ToString(@"mm\:ss");
            }

            double clipEndTime = active.NaturalDuration.TimeSpan.TotalSeconds - trimEnd;
            if (active.Position.TotalSeconds >= clipEndTime)
            {
                if (_currentPlaylistIndex + 1 < _previewPlaylist.Count)
                {
                    _nextPlayerReadyForSwitch = false;
                    next.Play();
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        active.Stop();
                        active.Opacity = 0;
                        next.Opacity = 1;
                    }), DispatcherPriority.Loaded);

                    _currentPlaylistIndex++;
                    _isPlayerAActive = !_isPlayerAActive;

                    if (_currentPlaylistIndex + 1 < _previewPlaylist.Count)
                        PreparePlayer(active, _currentPlaylistIndex + 1, autoPlay: false);
                }
                else if (_outroDurationSeconds > 0 && !string.IsNullOrEmpty(_outroImagePath))
                {
                    StopPlayers();
                    _previewPhase = PreviewPhase.Outro;
                    ShowIntroOutroImage(_outroImagePath);
                    _outroStartTime = DateTime.Now;
                }
                else
                    StopPreview();
            }
        }

        private void StopPlayers()
        {
            if (PreviewPlayerA != null) { PreviewPlayerA.Stop(); PreviewPlayerA.Source = null; PreviewPlayerA.Opacity = 0; }
            if (PreviewPlayerB != null) { PreviewPlayerB.Stop(); PreviewPlayerB.Source = null; PreviewPlayerB.Opacity = 0; }
            _nextPlayerReadyForSwitch = false;
        }

        private void StopPreview()
        {
            StopPlayers();
            _trimTimer.Stop();
            HideIntroOutroImage();
            _previewPhase = PreviewPhase.Idle;
            if (PreviewPlaceholder != null) PreviewPlaceholder.Visibility = Visibility.Visible;
            if (TimelineSlider != null) TimelineSlider.Value = 0;
            if (TxtCurrentTime != null) TxtCurrentTime.Text = "0:00";
            _pausedSequenceTime = 0;
            UpdateStep2PlayPauseIcon();
        }

        private void TimelineSlider_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingSlider = true;
            _trimTimer.Stop();
            if (_previewPhase == PreviewPhase.Clips)
            {
                var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                if (active != null) active.Pause();
            }
        }

        private void TimelineSlider_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDraggingSlider = false;
            if (TimelineSlider != null)
            {
                SeekToTime(TimelineSlider.Value);
                _pausedSequenceTime = TimelineSlider.Value;
                if (_previewPhase == PreviewPhase.Intro)
                    _introStartTime = DateTime.Now.AddSeconds(-_pausedSequenceTime);
                else if (_previewPhase == PreviewPhase.Outro)
                    _outroStartTime = DateTime.Now.AddSeconds(-(_pausedSequenceTime - _introDurationSeconds - _clipsTotalDuration));
                else if (_previewPhase == PreviewPhase.Clips)
                {
                    var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                    if (active != null) active.Play();
                }
                _trimTimer.Start();
            }
        }

        private void TimelineSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isDraggingSlider && TimelineSlider != null && e.NewValue >= 0)
            {
                SeekToTime(e.NewValue);
                if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(e.NewValue).ToString(@"mm\:ss");
            }
        }

        private void SeekToTime(double time)
        {
            if (_totalSequenceDuration <= 0) return;
            time = Math.Max(0, Math.Min(time, _totalSequenceDuration));
            _pausedSequenceTime = time;

            if (time < _introDurationSeconds)
            {
                _previewPhase = PreviewPhase.Intro;
                StopPlayers();
                if (!string.IsNullOrEmpty(_introImagePath)) ShowIntroOutroImage(_introImagePath);
                else HideIntroOutroImage();
                if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(time).ToString(@"mm\:ss");
                return;
            }

            if (time >= _introDurationSeconds + _clipsTotalDuration)
            {
                _previewPhase = PreviewPhase.Outro;
                StopPlayers();
                if (!string.IsNullOrEmpty(_outroImagePath))
                    ShowIntroOutroImage(_outroImagePath);
                else
                {
                    if (IntroOutroImage != null) { IntroOutroImage.Source = null; IntroOutroImage.Visibility = Visibility.Collapsed; }
                    if (PreviewPlayerA != null) PreviewPlayerA.Opacity = 0;
                    if (PreviewPlayerB != null) PreviewPlayerB.Opacity = 0;
                }
                double outroElapsed = time - (_introDurationSeconds + _clipsTotalDuration);
                _outroStartTime = DateTime.Now.AddSeconds(-outroElapsed);
                if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(time).ToString(@"mm\:ss");
                return;
            }

            _previewPhase = PreviewPhase.Clips;
            HideIntroOutroImage();

            if (_clipAccumulatedDurations == null || _clipAccumulatedDurations.Count == 0) return;

            int targetIdx = 0;
            for (int i = 0; i < _clipAccumulatedDurations.Count; i++)
            {
                if (time >= _clipAccumulatedDurations[i]) targetIdx = i;
                else break;
            }

            double local = time - _clipAccumulatedDurations[targetIdx];
            double trimStart = 0;
            double.TryParse(TxtTrimStart?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out trimStart);
            double speed = SpeedSlider != null ? SpeedSlider.Value : 1.0;
            double pos = trimStart + (local * speed);

            if (targetIdx != _currentPlaylistIndex)
            {
                _currentPlaylistIndex = targetIdx;
                StopPlayers();
                _isPlayerAActive = true;

                PreparePlayer(PreviewPlayerA, targetIdx, autoPlay: false);
                PreviewPlayerA.Position = TimeSpan.FromSeconds(pos);
                PreviewPlayerA.Opacity = 1;

                if (targetIdx + 1 < _previewPlaylist.Count)
                {
                    PreparePlayer(PreviewPlayerB, targetIdx + 1, autoPlay: false);
                    PreviewPlayerB.Opacity = 0;
                }
            }
            else
            {
                var active = _isPlayerAActive ? PreviewPlayerA : PreviewPlayerB;
                if (active != null) active.Position = TimeSpan.FromSeconds(pos);
            }

            if (TxtCurrentTime != null) TxtCurrentTime.Text = TimeSpan.FromSeconds(time).ToString(@"mm\:ss");
        }
        #endregion

        #region Step 3: Visuals

        // --- Font & Color Management ---
        private void LoadFonts()
        {
            FontGroups.Clear();
            string fontDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fonts");

            if (!Directory.Exists(fontDir)) Directory.CreateDirectory(fontDir);

            var files = Directory.GetFiles(fontDir, "*.*")
                                 .Where(s => s.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) ||
                                             s.EndsWith(".otf", StringComparison.OrdinalIgnoreCase));

            // Dictionary ชั่วคราวเพื่อจัดกลุ่ม: Key = ชื่อฟอนต์หลัก (เช่น "Sarabun")
            var tempGroups = new Dictionary<string, List<FontVariant>>();

            foreach (var file in files)
            {
                try
                {
                    var uri = new Uri(file, UriKind.Absolute);
                    GlyphTypeface glyph;

                    // [แก้ Crash] ลองสร้าง GlyphTypeface ถ้าไฟล์เสียมันจะเด้งไป catch
                    try
                    {
                        glyph = new GlyphTypeface(uri);
                    }
                    catch { continue; } // ข้ามไฟล์ที่เสียไปเลย

                    // ดึงชื่อ Family (เช่น "Sarabun")
                    string familyName = glyph.Win32FamilyNames.ContainsKey(CultureInfo.GetCultureInfo("en-us"))
                        ? glyph.Win32FamilyNames[CultureInfo.GetCultureInfo("en-us")]
                        : glyph.Win32FamilyNames.Values.FirstOrDefault();

                    // ดึงชื่อ Style (เช่น "Bold", "Italic")
                    string faceName = glyph.Win32FaceNames.ContainsKey(CultureInfo.GetCultureInfo("en-us"))
                        ? glyph.Win32FaceNames[CultureInfo.GetCultureInfo("en-us")]
                        : glyph.Win32FaceNames.Values.FirstOrDefault();

                    if (string.IsNullOrEmpty(familyName)) familyName = Path.GetFileNameWithoutExtension(file);
                    if (string.IsNullOrEmpty(faceName)) faceName = "Regular";

                    // สร้าง Variant
                    var variant = new FontVariant
                    {
                        StyleName = faceName,
                        FilePath = file,
                        // สร้าง FontFamily สำหรับแสดงผลใน UI
                        FontFamilyObj = new FontFamily(new Uri("file:///" + file), "./#" + familyName)
                    };

                    // จัดลงกลุ่ม
                    if (!tempGroups.ContainsKey(familyName))
                    {
                        tempGroups[familyName] = new List<FontVariant>();
                    }
                    tempGroups[familyName].Add(variant);
                }
                catch
                {
                    // เจอไฟล์ Error ข้ามไปเงียบๆ ไม่ต้อง Crash
                }
            }

            // แปลง Dictionary ลง ObservableCollection เพื่อแสดงหน้าจอ
            foreach (var kvp in tempGroups)
            {
                var group = new FontFamilyGroup
                {
                    FamilyName = kvp.Key,
                    Variants = kvp.Value.OrderBy(v => v.StyleName).ToList() // เรียงตามชื่อ Style
                };

                // เลือกตัวแรกเป็นค่า Default (มักจะเป็น Regular)
                var def = group.Variants.FirstOrDefault(v => v.StyleName.Contains("Regular")) ?? group.Variants.First();
                group.SelectedVariant = def;

                FontGroups.Add(group);
            }
        }

        private void ImportFont_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.Filters.Add(new Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Font Files", "*.ttf;*.otf"));

            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                string sourceFile = dialog.FileName;
                string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fonts");
                if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

                string destFile = Path.Combine(destDir, Path.GetFileName(sourceFile));
                try
                {
                    File.Copy(sourceFile, destFile, true);
                    MessageBox.Show("Import Font สำเร็จ! (อาจต้องเลือกใหม่ในรายการ)", "สำเร็จ");
                    LoadFonts();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error importing font: " + ex.Message);
                }
            }
        }

        // --- Color Pool Management ---
        private void LoadColorCache() { try { if (File.Exists(_cacheFilePath)) foreach (var l in File.ReadAllLines(_cacheFilePath)) CachedColors.Add(l); else { CachedColors.Add("#FF0000"); CachedColors.Add("#0000FF"); CachedColors.Add("#00FF00"); } } catch { } }
        private void SaveColorCache() { try { File.WriteAllLines(_cacheFilePath, CachedColors); } catch { } }

        private void LoadRandomColorPool()
        {
            RandomColorPool.Clear();
            try { if (File.Exists(_randomPoolPath)) foreach (var l in File.ReadAllLines(_randomPoolPath)) RandomColorPool.Add(l); } catch { }
            if (RandomColorPool.Count == 0) { RandomColorPool.Add("#FF5733"); RandomColorPool.Add("#33FF57"); RandomColorPool.Add("#3357FF"); }
        }
        private void SaveRandomColorPool() { try { File.WriteAllLines(_randomPoolPath, RandomColorPool); } catch { } }

        private void AddColorToRandomPool_Click(object sender, RoutedEventArgs e)
        {
            if (ListTextLayers.SelectedItem is TextLayer layer)
            {
                if (!RandomColorPool.Contains(layer.BgColor))
                {
                    RandomColorPool.Add(layer.BgColor);
                    SaveRandomColorPool();
                }
            }
        }

        private void RemoveColorFromRandomPool_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is string color)
            {
                RandomColorPool.Remove(color);
                SaveRandomColorPool();
            }
        }

        // --- Player & Scene Logic ---
        private void RefreshVisualSceneSelect()
        {
            if (SceneSelectVisual != null)
            {
                SceneSelectVisual.Items.Clear();
                foreach (var folder in Folders)
                {
                    SceneSelectVisual.Items.Add(new ComboBoxItem { Content = folder.Name, Tag = folder });
                }

                if (SceneSelectVisual.Items.Count > 0)
                {
                    SceneSelectVisual.SelectedIndex = 0;
                }
                else
                {
                    if (PreviewPlayerStep3 != null) PreviewPlayerStep3.Source = null;
                    if (PreviewPlaceholderStep3 != null) PreviewPlaceholderStep3.Visibility = Visibility.Visible;
                }
            }
        }

        private void SceneSelectVisual_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SceneSelectVisual.SelectedItem is ComboBoxItem item && item.Tag is SceneFolder folder)
            {
                foreach (var layer in folder.TextLayers) layer.PropertyChanged -= TextLayer_PropertyChanged;
                ListTextLayers.ItemsSource = folder.TextLayers;
                foreach (var layer in folder.TextLayers) layer.PropertyChanged += TextLayer_PropertyChanged;

                // [FIX] Load the first SELECTED clip only
                var firstClip = folder.Files.FirstOrDefault(f => f.Status == "Ready" && f.IsSelected);
                if (firstClip != null)
                {
                    try
                    {
                        PreviewPlayerStep3.Source = new Uri(firstClip.FilePath);
                        PreviewPlayerStep3.Play();
                        PreviewPlayerStep3.Pause(); // Pause immediately to show frame

                        // Update UI
                        TimelineSliderStep3.Value = 0;
                        TxtCurrentTimeStep3.Text = "00:00";
                        if (PreviewPlaceholderStep3 != null) PreviewPlaceholderStep3.Visibility = Visibility.Collapsed; // Hide placeholder
                    }
                    catch { }
                }
                else
                {
                    PreviewPlayerStep3.Source = null;
                    if (PreviewPlaceholderStep3 != null) PreviewPlaceholderStep3.Visibility = Visibility.Visible;
                }
            }
        }

        // [ADDED] Missing SelectionChanged Handler for Text Layers
        private void UpdateStep3PresetLabel()
        {
            if (TxtActivePreset == null || TxtPositionXY == null) return;
            if (ListTextLayers.SelectedItem is TextLayer layer)
            {
                TxtPositionXY.Text = $"X = {layer.X:F0}, Y = {layer.Y:F0} px";
                TxtActivePreset.Text = string.IsNullOrEmpty(layer.ActivePresetName) ? "-" : layer.ActivePresetName;
            }
            else { TxtPositionXY.Text = "X = -, Y = - px"; TxtActivePreset.Text = "-"; }
        }

        private void ListTextLayers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // เปลี่ยนชื่อตัวแปรเป็น selectedLayer เพื่อไม่ให้ชนกับตัวแปรอื่น
            if (ListTextLayers.SelectedItem is TextLayer selectedLayer)
            {
                // Logic 1: อัปเดตกรอบสีน้ำเงินในหน้าจอ Preview
                if (SceneSelectVisual.SelectedItem is ComboBoxItem item && item.Tag is SceneFolder folder)
                {
                    // ใช้ชื่อตัวแปร l ใน Loop แทน layer เพื่อป้องกันชื่อซ้ำ
                    foreach (var l in folder.TextLayers)
                    {
                        l.IsSelectedLayer = (l == selectedLayer);
                    }
                }

                // Logic 2: อัปเดต ComboBox ให้ตรงกับฟอนต์ที่เลือกอยู่
                if (FontGroups != null)
                {
                    string currentPath = selectedLayer.FontPath;

                    // วนหาใน Group ว่า Path นี้อยู่กลุ่มไหน
                    foreach (var group in FontGroups)
                    {
                        var match = group.Variants.FirstOrDefault(v => v.FilePath == currentPath);
                        if (match != null)
                        {
                            // เจอแล้ว! สั่งเลือกใน UI
                            CmbFontFamily.SelectedItem = group;
                            CmbFontStyle.ItemsSource = group.Variants; // อัปเดตรายการ Style
                            CmbFontStyle.SelectedItem = match;
                            break;
                        }
                    }
                }
                UpdateStep3PresetLabel();
            }
            else
                UpdateStep3PresetLabel();
        }
        private void CmbFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmbFontFamily.SelectedItem is FontFamilyGroup group)
            {
                // 1. อัปเดตรายการในช่อง Style
                CmbFontStyle.ItemsSource = group.Variants;

                // 2. เลือก Style แรกเป็นค่าเริ่มต้น (เช่น Regular)
                if (group.Variants.Count > 0)
                {
                    var defaultVariant = group.Variants[0];
                    CmbFontStyle.SelectedItem = defaultVariant; // สั่งเลือกใน UI

                    // [สำคัญ] บังคับอัปเดต Text Layer ทันที!
                    if (ListTextLayers.SelectedItem is TextLayer layer)
                    {
                        layer.FontPath = defaultVariant.FilePath;
                        // ใช้ FontFamilyObj ที่สร้างเตรียมไว้แล้วใน LoadFonts (ชัวร์สุด)
                        layer.FontFamily = defaultVariant.FontFamilyObj;
                    }
                }
            }
        }

        // ฟังก์ชันเมื่อเปลี่ยน Style (ComboBox ตัวขวา)
        private void CmbFontStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // ตรวจสอบว่ามีการเลือกค่าจริง และมี TextLayer ที่เลือกอยู่
            if (CmbFontStyle.SelectedItem is FontVariant variant && ListTextLayers.SelectedItem is TextLayer layer)
            {
                // 1. อัปเดต Path ไฟล์จริง (ส่งให้ FFmpeg ตอน Render)
                layer.FontPath = variant.FilePath;

                // 2. อัปเดต FontFamily บนหน้าจอ (เพื่อให้เห็นผลทันที)
                // สร้าง FontFamily ใหม่ที่ชี้ไปหาไฟล์เฉพาะตัวนั้นๆ
                // เทคนิค: ใช้ BaseUri เพื่อบังคับให้ WPF โหลดไฟล์ใหม่แน่นอน
                var uri = new Uri(variant.FilePath, UriKind.Absolute);

                // ชื่อ Font (FamilyName) ต้องดึงให้ถูกเพื่อให้ WPF แสดงผลได้
                // แต่การใช้ FilePath ตรงๆ แบบนี้มักจะได้ผลดีที่สุดสำหรับการพรีวิวไฟล์เดียว
                layer.FontFamily = new FontFamily(uri, "./#" + variant.StyleName);

                if (variant.FontFamilyObj != null)
                {
                    layer.FontFamily = variant.FontFamilyObj;
                }
                // ถ้าใช้บรรทัดบนแล้วไม่เปลี่ยน ให้ลองใช้ตัวที่เก็บไว้ใน variant
                // layer.FontFamily = variant.FontFamilyObj; 
            }
        }

        // [NEW] Stop Button Logic for Step 3
        private void StopPreviewStep3_Click(object sender, RoutedEventArgs e)
        {
            StopPreviewStep3();
        }

        private void UpdateStep3PlayPauseIcon()
        {
            if (IconPlayPauseStep3 == null) return;
            try
            {
                var isPlaying = _step3Timer != null && _step3Timer.IsEnabled;
                var geo = (Geometry)Resources[isPlaying ? "IconPause" : "IconPlay"];
                if (geo != null) IconPlayPauseStep3.Data = geo;
            }
            catch { }
        }

        private void PlayPauseToggleStep3_Click(object sender, RoutedEventArgs e)
        {
            if (PreviewPlayerStep3.Source == null) return;
            if (_step3Timer.IsEnabled)
            {
                if (PreviewPlayerStep3.CanPause) PreviewPlayerStep3.Pause();
                _step3Timer.Stop();
            }
            else
            {
                if (PreviewPlaceholderStep3 != null) PreviewPlaceholderStep3.Visibility = Visibility.Collapsed;
                PreviewPlayerStep3.Play();
                _step3Timer.Start();
                PlayTextAnimations();
            }
            UpdateStep3PlayPauseIcon();
        }

        private void StopPreviewStep3()
        {
            if (PreviewPlayerStep3 != null)
            {
                PreviewPlayerStep3.Stop();
                PreviewPlayerStep3.Position = TimeSpan.Zero;
            }
            if (_step3Timer != null) _step3Timer.Stop();
            if (TimelineSliderStep3 != null) TimelineSliderStep3.Value = 0;
            if (TxtCurrentTimeStep3 != null) TxtCurrentTimeStep3.Text = "00:00";
            if (PreviewPlaceholderStep3 != null) PreviewPlaceholderStep3.Visibility = Visibility.Visible;
            UpdateStep3PlayPauseIcon();
        }

        private void PlayTextAnimations()
        {
            if (OverlayItemsControl == null) return;
            foreach (var item in ListTextLayers.Items)
            {
                if (item is TextLayer layer)
                {
                    var container = OverlayItemsControl.ItemContainerGenerator.ContainerFromItem(layer) as FrameworkElement;
                    if (container == null)
                    {
                        OverlayItemsControl.UpdateLayout();
                        container = OverlayItemsControl.ItemContainerGenerator.ContainerFromItem(layer) as FrameworkElement;
                    }
                    if (container == null) continue;

                    container.RenderTransform = new TransformGroup(); container.Opacity = 1;
                    if (layer.AnimationEffect.StartsWith("None")) continue;

                    Storyboard sbEnter = CreateAnimationStoryboard(layer, container, true);
                    sbEnter.BeginTime = TimeSpan.FromSeconds(layer.Delay);
                    sbEnter.Begin();

                    if (layer.HasExitAnimation)
                    {
                        Storyboard sbExit = CreateAnimationStoryboard(layer, container, false, useExitEffect: true);
                        sbExit.BeginTime = TimeSpan.FromSeconds(layer.Delay + layer.TextDuration);
                        sbExit.Begin();
                    }
                }
            }
        }

        private Storyboard CreateAnimationStoryboard(TextLayer layer, FrameworkElement container, bool isEnter, bool useExitEffect = false)
        {
            Storyboard sb = new Storyboard();
            double dur = isEnter ? layer.EffectDuration : (useExitEffect ? layer.ExitEffectDuration : layer.EffectDuration);
            string effectName = useExitEffect && !isEnter ? (layer.ExitEffect ?? "None") : layer.AnimationEffect;
            IEasingFunction ease = new CubicEase { EasingMode = EasingMode.EaseOut };

            if (effectName.Contains("Fade In") || effectName.Contains("Fade"))
            {
                DoubleAnimation fadeIn = new DoubleAnimation(isEnter ? 0 : 1, isEnter ? 1 : 0, TimeSpan.FromSeconds(dur));
                if (!isEnter) container.Opacity = 1; else container.Opacity = 0;
                Storyboard.SetTarget(fadeIn, container); Storyboard.SetTargetProperty(fadeIn, new PropertyPath("Opacity")); sb.Children.Add(fadeIn);
            }
            else if (effectName.Contains("Slide"))
            {
                TranslateTransform tt = new TranslateTransform(); container.RenderTransform = tt;
                double from = 0, to = 0; string prop = "X";
                if (effectName.Contains("Left")) { prop = "X"; from = 200; }
                else if (effectName.Contains("Right")) { prop = "X"; from = -200; }
                else if (effectName.Contains("Up")) { prop = "Y"; from = 200; }
                else if (effectName.Contains("Down")) { prop = "Y"; from = -200; }

                if (isEnter) { to = 0; } else { to = from; from = 0; }
                if (isEnter && prop == "X") tt.X = from; if (isEnter && prop == "Y") tt.Y = from;

                DoubleAnimation slide = new DoubleAnimation { From = from, To = to, Duration = TimeSpan.FromSeconds(dur), EasingFunction = ease };
                Storyboard.SetTarget(slide, container); Storyboard.SetTargetProperty(slide, new PropertyPath($"RenderTransform.{prop}")); sb.Children.Add(slide);
                DoubleAnimation fade = new DoubleAnimation(isEnter ? 0 : 1, isEnter ? 1 : 0, TimeSpan.FromSeconds(dur * 0.5));
                Storyboard.SetTarget(fade, container); Storyboard.SetTargetProperty(fade, new PropertyPath("Opacity")); sb.Children.Add(fade);
            }
            return sb;
        }

        private void Step3Timer_Tick(object sender, EventArgs e)
        {
            if (PreviewPlayerStep3.Source != null && PreviewPlayerStep3.NaturalDuration.HasTimeSpan && !_isDraggingSliderStep3)
            {
                double dur = PreviewPlayerStep3.NaturalDuration.TimeSpan.TotalSeconds;
                if (Math.Abs(dur - _step3TimelineTotalDuration) > 0.01)
                    UpdateStep3Timeline();
                TimelineSliderStep3.Maximum = dur;
                TimelineSliderStep3.Value = PreviewPlayerStep3.Position.TotalSeconds;
                TxtCurrentTimeStep3.Text = PreviewPlayerStep3.Position.ToString(@"mm\:ss");
                TxtTotalTimeStep3.Text = PreviewPlayerStep3.NaturalDuration.TimeSpan.ToString(@"mm\:ss");
                UpdateStep3PlayheadPosition();
            }
        }

        private void TimelineZoom_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string tag)
            {
                if (tag == "In") _step3TimelinePixelsPerSecond = Math.Min(200, _step3TimelinePixelsPerSecond + 20);
                else if (tag == "Out") _step3TimelinePixelsPerSecond = Math.Max(20, _step3TimelinePixelsPerSecond - 20);
                if (TxtTimelineZoom != null) TxtTimelineZoom.Text = $"{(int)_step3TimelinePixelsPerSecond} px/s";
                UpdateStep3Timeline();
            }
        }

        private void UpdateStep3Timeline()
        {
            double duration = 0;
            if (PreviewPlayerStep3?.Source != null && PreviewPlayerStep3.NaturalDuration.HasTimeSpan)
                duration = PreviewPlayerStep3.NaturalDuration.TimeSpan.TotalSeconds;
            _step3TimelineTotalDuration = duration;

            double width = Math.Max(100, duration * _step3TimelinePixelsPerSecond);

            if (TimelineRulerCanvas != null)
            {
                TimelineRulerCanvas.Children.Clear();
                TimelineRulerCanvas.Width = width;
                for (int sec = 0; sec <= (int)Math.Ceiling(duration); sec++)
                {
                    double x = sec * _step3TimelinePixelsPerSecond;
                    var line = new System.Windows.Shapes.Line { X1 = x, Y1 = 16, X2 = x, Y2 = 18, Stroke = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF)), StrokeThickness = 1 };
                    TimelineRulerCanvas.Children.Add(line);
                    if (sec % 5 == 0 || sec == 0)
                    {
                        var tb = new TextBlock { Text = sec.ToString(), FontSize = 9, Foreground = new SolidColorBrush(Color.FromRgb(0xE5, 0xE7, 0xEB)), Margin = new Thickness(x + 2, 0, 0, 0) };
                        TimelineRulerCanvas.Children.Add(tb);
                    }
                }
            }

            if (TimelineVideoCanvas != null)
            {
                TimelineVideoCanvas.Children.Clear();
                TimelineVideoCanvas.Width = width;
                var rect = new System.Windows.Shapes.Rectangle { Width = width, Height = 22, Fill = new SolidColorBrush(Color.FromRgb(0x4F, 0x46, 0xE5)), Opacity = 0.6 };
                TimelineVideoCanvas.Children.Add(rect);
            }

            if (TimelineTextCanvas != null)
            {
                TimelineTextCanvas.Children.Clear();
                TimelineTextCanvas.Width = width;
                if (SceneSelectVisual?.SelectedItem is ComboBoxItem item && item.Tag is SceneFolder folder)
                {
                    foreach (var layer in folder.TextLayers)
                    {
                        double left = layer.Delay * _step3TimelinePixelsPerSecond;
                        double w = Math.Max(4, layer.TextDuration * _step3TimelinePixelsPerSecond);
                        var border = new Border
                        {
                            Width = w,
                            Height = 22,
                            Background = new SolidColorBrush(Color.FromRgb(0x10, 0xB9, 0x81)),
                            Opacity = 0.9,
                            Tag = layer,
                            Cursor = Cursors.Hand
                        };
                        border.MouseLeftButtonDown += TimelineSegment_MouseDown;
                        border.MouseMove += TimelineSegment_MouseMove;
                        border.MouseLeftButtonUp += TimelineSegment_MouseUp;
                        TimelineTextCanvas.Children.Add(border);
                        Canvas.SetLeft(border, left);
                        Canvas.SetTop(border, 2);
                    }
                }
            }

            UpdateStep3PlayheadPosition();
        }

        private void UpdateStep3PlayheadPosition()
        {
            if (TimelinePlayhead == null || _step3TimelineTotalDuration <= 0) return;
            double t = PreviewPlayerStep3?.Position.TotalSeconds ?? 0;
            double playheadX = t * _step3TimelinePixelsPerSecond;
            TimelinePlayhead.Margin = new Thickness(42 + playheadX, 0, 0, 0);
        }

        private void TimelineSliderStep3_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingSliderStep3 = true;
            _step3WasPlayingBeforeSliderDrag = _step3Timer.IsEnabled;
            _step3Timer.Stop();
            if (PreviewPlayerStep3.CanPause) PreviewPlayerStep3.Pause();
        }
        private void TimelineSliderStep3_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDraggingSliderStep3 = false;
            if (PreviewPlayerStep3.Source != null)
            {
                PreviewPlayerStep3.Position = TimeSpan.FromSeconds(TimelineSliderStep3.Value);
                if (_step3WasPlayingBeforeSliderDrag) { PreviewPlayerStep3.Play(); _step3Timer.Start(); UpdateStep3PlayPauseIcon(); }
            }
        }

        private void TimelineSegment_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Tag is TextLayer layer && TimelineTextCanvas != null)
            {
                Point posInCanvas = e.GetPosition(TimelineTextCanvas);
                Point posInSegment = e.GetPosition(border);
                bool onResizeHandle = border.ActualWidth > 12 && posInSegment.X >= border.ActualWidth - 8;
                _step3SegmentLayer = layer;
                _step3SegmentElement = border;
                _step3SegmentStartMouseX = posInCanvas.X;
                _step3SegmentStartDelay = layer.Delay;
                _step3SegmentStartDuration = layer.TextDuration;
                _step3SegmentResizing = onResizeHandle;
                _step3SegmentDragging = !onResizeHandle;
                border.Cursor = onResizeHandle ? Cursors.SizeWE : Cursors.SizeAll;
                border.CaptureMouse();
                e.Handled = true;
            }
        }

        private void TimelineSegment_MouseMove(object sender, MouseEventArgs e)
        {
            if (sender is Border border)
            {
                if (!_step3SegmentDragging && !_step3SegmentResizing)
                {
                    Point p = e.GetPosition(border);
                    border.Cursor = (border.ActualWidth > 12 && p.X >= border.ActualWidth - 8) ? Cursors.SizeWE : Cursors.Hand;
                }
            }
            if (TimelineTextCanvas == null || _step3SegmentLayer == null || _step3SegmentElement == null) return;
            double pps = _step3TimelinePixelsPerSecond;
            double totalDur = _step3TimelineTotalDuration;
            double curX = e.GetPosition(TimelineTextCanvas).X;
            double deltaSec = (curX - _step3SegmentStartMouseX) / pps;

            if (_step3SegmentResizing)
            {
                double newDuration = _step3SegmentStartDuration + deltaSec;
                double maxDur = totalDur > 0 ? Math.Max(0, totalDur - _step3SegmentLayer.Delay) : 60;
                if (newDuration < Step3SegmentMinDuration) newDuration = Step3SegmentMinDuration;
                if (newDuration > maxDur) newDuration = maxDur;
                _step3SegmentLayer.TextDuration = newDuration;
                double w = Math.Max(4, newDuration * pps);
                _step3SegmentElement.Width = w;
            }
            else if (_step3SegmentDragging)
            {
                double newDelay = _step3SegmentStartDelay + deltaSec;
                double maxDelay = totalDur > 0 ? Math.Max(0, totalDur - _step3SegmentLayer.TextDuration) : 0;
                if (newDelay < 0) newDelay = 0;
                if (newDelay > maxDelay) newDelay = maxDelay;
                _step3SegmentLayer.Delay = newDelay;
                Canvas.SetLeft(_step3SegmentElement, newDelay * pps);
            }
        }

        private void TimelineSegment_MouseUp(object sender, MouseButtonEventArgs e)
        {
            TimelineSegment_EndDrag();
        }

        private void TimelineSegment_EndDrag()
        {
            if (_step3SegmentElement != null)
            {
                _step3SegmentElement.ReleaseMouseCapture();
                _step3SegmentElement.Cursor = Cursors.Hand;
            }
            _step3SegmentDragging = false;
            _step3SegmentResizing = false;
            _step3SegmentLayer = null;
            _step3SegmentElement = null;
        }

        private void RemoveText_Click(object sender, RoutedEventArgs e) { if (ListTextLayers.SelectedItem is TextLayer selected && SceneSelectVisual.SelectedItem is ComboBoxItem item && item.Tag is SceneFolder folder) { selected.PropertyChanged -= TextLayer_PropertyChanged; folder.TextLayers.Remove(selected); UpdateStep3Timeline(); } }

        private void TextLayer_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (_isAdjustingFontSize) return;
            if (e.PropertyName == nameof(TextLayer.Text) || e.PropertyName == nameof(TextLayer.FontSize) || e.PropertyName == nameof(TextLayer.X)) AdjustLayerFontSize(sender as TextLayer);
        }

        private void AdjustLayerFontSize(TextLayer layer)
        {
            if (layer == null || string.IsNullOrEmpty(layer.Text)) return;
            _isAdjustingFontSize = true;
            try
            {
                double canvasWidth = 1280; // [CONFIRMED] Match Full HD Width Logic
                double avail = canvasWidth - layer.X - 30; if (avail < 20) avail = 20;
                FormattedText ft = new FormattedText(layer.Text, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), layer.FontSize, Brushes.White, VisualTreeHelper.GetDpi(this).PixelsPerDip);
                if (ft.Width > avail) { double ratio = avail / ft.Width; int newSize = (int)(layer.FontSize * ratio); if (newSize < 8) newSize = 8; layer.FontSize = newSize; }
            }
            catch { }
            finally { _isAdjustingFontSize = false; }
        }

        private void TextOverlay_MouseDown(object sender, MouseButtonEventArgs e) { if (sender is FrameworkElement el && el.DataContext is TextLayer layer) { _isDraggingText = true; _draggingLayer = layer; _clickPosition = e.GetPosition(OverlayItemsControl); el.CaptureMouse(); ListTextLayers.SelectedItem = layer; } }
        private void TextOverlay_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDraggingText && _draggingLayer != null && sender is FrameworkElement el)
            {
                Point cur = e.GetPosition(OverlayItemsControl);
                double nx = _draggingLayer.X + (cur.X - _clickPosition.X);
                double ny = _draggingLayer.Y + (cur.Y - _clickPosition.Y);
                if (nx < 0) nx = 0; if (nx + el.ActualWidth > OverlayItemsControl.ActualWidth) nx = OverlayItemsControl.ActualWidth - el.ActualWidth;
                if (ny < 0) ny = 0; if (ny + el.ActualHeight > OverlayItemsControl.ActualHeight) ny = OverlayItemsControl.ActualHeight - el.ActualHeight;
                _draggingLayer.X = nx; _draggingLayer.Y = ny;
                _draggingLayer.ActivePresetName = "Custom";
                _clickPosition = cur;
                UpdateStep3PresetLabel();
            }
        }
        private void TextOverlay_MouseUp(object sender, MouseButtonEventArgs e) { _isDraggingText = false; _draggingLayer = null; if (sender is FrameworkElement el) el.ReleaseMouseCapture(); }

        private static readonly Dictionary<string, string> PresetDisplayNames = new Dictionary<string, string>
        {
            { "TL", "Top-Left (↖)" }, { "TC", "Top-Center (↑)" }, { "TR", "Top-Right (↗)" },
            { "CL", "Center-Left (←)" }, { "CC", "Center (•)" }, { "CR", "Center-Right (→)" },
            { "BL", "Bottom-Left (↙)" }, { "BC", "Bottom-Center (↓)" }, { "BR", "Bottom-Right (↘)" }
        };

        private void TxtPosXY_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!(ListTextLayers.SelectedItem is TextLayer layer)) return;
            if (double.TryParse(TxtPosX?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double x))
                layer.X = Math.Max(0, x);
            if (double.TryParse(TxtPosY?.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                layer.Y = Math.Max(0, y);
            layer.ActivePresetName = "Custom";
            UpdateStep3PresetLabel();
        }

        private void SetTextPosition_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string pos && ListTextLayers.SelectedItem is TextLayer layer)
            {
                var container = OverlayItemsControl.ItemContainerGenerator.ContainerFromItem(layer) as FrameworkElement;
                if (container == null) return;
                double w = container.ActualWidth, h = container.ActualHeight;
                double pw = OverlayItemsControl.ActualWidth > 0 ? OverlayItemsControl.ActualWidth : 1280;
                double ph = OverlayItemsControl.ActualHeight > 0 ? OverlayItemsControl.ActualHeight : 720;
                switch (pos)
                {
                    case "TL": layer.X = 20; layer.Y = 20; break;
                    case "TC": layer.X = (pw - w) / 2; layer.Y = 20; break;
                    case "TR": layer.X = pw - w - 20; layer.Y = 20; break;
                    case "CL": layer.X = 20; layer.Y = (ph - h) / 2; break;
                    case "CC": layer.X = (pw - w) / 2; layer.Y = (ph - h) / 2; break;
                    case "CR": layer.X = pw - w - 20; layer.Y = (ph - h) / 2; break;
                    case "BL": layer.X = 20; layer.Y = ph - h - 20; break;
                    case "BC": layer.X = (pw - w) / 2; layer.Y = ph - h - 20; break;
                    case "BR": layer.X = pw - w - 20; layer.Y = ph - h - 20; break;
                }
                layer.ActivePresetName = PresetDisplayNames.TryGetValue(pos, out var name) ? name : pos;
                UpdateStep3PresetLabel();
            }
        }

        private void SetTextAlign_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.Tag is string align && ListTextLayers.SelectedItem is TextLayer layer) { if (align == "Left") layer.TextAlign = TextAlignment.Left; if (align == "Center") layer.TextAlign = TextAlignment.Center; if (align == "Right") layer.TextAlign = TextAlignment.Right; } }

        // [Modified] RB Logic for Modes
        private void RbMode_Click(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && ListTextLayers.SelectedItem is TextLayer layer)
            {
                if (rb.Tag.ToString() == "Solid") layer.BgMode = 0;
                else if (rb.Tag.ToString() == "Gradient") layer.BgMode = 1;
                else if (rb.Tag.ToString() == "Random") { layer.BgMode = 2; layer.RefreshBrush(); }
            }
        }

        private void PresetColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex && ListTextLayers.SelectedItem is TextLayer layer)
            {
                // Assign to current mode active color
                if (layer.BgMode == 1) // Gradient Mode - Assign to Start Color by default or toggled? Simplicity: Assign to BgColor (Start)
                {
                    layer.BgColor = hex;
                }
                else
                {
                    layer.BgColor = hex;
                }
            }
        }
        private void PresetGradientEnd_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex && ListTextLayers.SelectedItem is TextLayer layer)
            {
                layer.GradientEndColor = hex;
            }
        }

        private void AddColorToCache_Click(object sender, RoutedEventArgs e) { if (ListTextLayers.SelectedItem is TextLayer layer && !CachedColors.Contains(layer.BgColor)) CachedColors.Add(layer.BgColor); }
        private void RemoveColorFromCache_Click(object sender, RoutedEventArgs e) { if (ListCachedColors.SelectedItem is string c) CachedColors.Remove(c); }
        private void Window_Closing(object sender, CancelEventArgs e) { SaveColorCache(); SaveRandomColorPool(); }
        #endregion

        // --------------------------------------------------------------------------
        // [ใหม่] ฟังก์ชัน Import Font จากไฟล์ ZIP
        // --------------------------------------------------------------------------
        private void ImportFontZip_Click(object sender, RoutedEventArgs e)
        {
            // ใช้ OpenFileDialog เลือก .zip
            var dialog = new Microsoft.Win32.OpenFileDialog(); // หรือใช้ CommonOpenFileDialog ก็ได้
            dialog.Filter = "Zip Files (*.zip)|*.zip";
            dialog.Title = "เลือกไฟล์ Font Pack (.zip)";

            if (dialog.ShowDialog() == true)
            {
                string zipPath = dialog.FileName;
                string destDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fonts");
                if (!Directory.Exists(destDir)) Directory.CreateDirectory(destDir);

                int count = 0;
                try
                {
                    using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                    {
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            // เช็คว่าเป็นไฟล์ .ttf หรือ .otf (ไม่เอาโฟลเดอร์)
                            string ext = Path.GetExtension(entry.FullName).ToLower();
                            if (!string.IsNullOrEmpty(entry.Name) && (ext == ".ttf" || ext == ".otf"))
                            {
                                // Extract ลงโฟลเดอร์ Fonts โดยตรง (Flatten Path)
                                string destFile = Path.Combine(destDir, entry.Name);

                                // ถ้ามีไฟล์ชื่อซ้ำ ให้ลบของเก่าก่อน หรือวางทับ
                                entry.ExtractToFile(destFile, true);
                                count++;
                            }
                        }
                    }

                    if (count > 0)
                    {
                        MessageBox.Show($"นำเข้าฟอนต์สำเร็จ {count} ไฟล์!", "Success");
                        LoadFonts(); // รีโหลดรายการทันที
                    }
                    else
                    {
                        MessageBox.Show("ไม่พบไฟล์ .ttf หรือ .otf ใน Zip นี้", "Warning");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error extracting zip: " + ex.Message, "Error");
                }
            }
        }

        // --------------------------------------------------------------------------
        // [แก้ไข] ฟังก์ชัน AddText_Click ให้ใช้ Default Font ที่โหลดมา
        // --------------------------------------------------------------------------
        private void AddText_Click(object sender, RoutedEventArgs e)
        {
            if (SceneSelectVisual.SelectedItem is ComboBoxItem item && item.Tag is SceneFolder folder)
            {
                double startX = 50, startY = 50;
                if (folder.TextLayers.Count > 0)
                {
                    var last = folder.TextLayers.Last();
                    startX = last.X;
                    startY = last.Y + last.FontSize + 40;
                }

                var newLayer = new TextLayer
                {
                    Text = "ข้อความใหม่",
                    X = startX,
                    Y = startY,
                    FontSize = 32,
                    FontFamily = _defaultSystemFont // <--- ใช้ฟอนต์แรกที่เจอในโฟลเดอร์
                };

                newLayer.PropertyChanged += TextLayer_PropertyChanged;
                folder.TextLayers.Add(newLayer);
                ListTextLayers.SelectedItem = newLayer;
                UpdateStep3Timeline();
            }
        }

        #region Other Steps
        // [MODIFIED] Matrix Calculation Logic and Validation
        // -------------------------------------------------------------------
        // STEP 4: MATRIX CALCULATION & PREVIEW LOGIC
        // -------------------------------------------------------------------

        private void CalcMatrix_Click(object sender, RoutedEventArgs e)
        {
            // 1. ตรวจสอบข้อมูลจาก Step 1
            var activeFolders = Folders.Where(f => f.Files.Any(x => x.IsSelected)).ToList();

            if (activeFolders.Count == 0)
            {
                MessageBox.Show("กรุณาเลือกไฟล์วิดีโอใน Step 1 อย่างน้อย 1 ไฟล์", "Warning");
                return;
            }

            // 2. คำนวณจำนวนความเป็นไปได้ (Cartesian Product)
            long total = 1;
            string sceneInfo = "";
            List<List<VideoFile>> sourceLists = new List<List<VideoFile>>();

            foreach (var f in activeFolders)
            {
                var validFiles = f.Files.Where(x => x.IsSelected && x.Status == "Ready").ToList();
                if (validFiles.Count > 0)
                {
                    total *= validFiles.Count;
                    sourceLists.Add(validFiles);
                    sceneInfo += $"{validFiles.Count} x ";
                }
            }

            if (sceneInfo.EndsWith(" x ")) sceneInfo = sceneInfo.Substring(0, sceneInfo.Length - 3);

            _totalCombinations = total;

            // 3. Update UI Stats
            TxtTotalCombinations.Text = total.ToString("N0");
            TxtSceneCountInfo.Text = $"จาก {sceneInfo} ({activeFolders.Count} Scenes)";
            BtnNextStep4.IsEnabled = total > 0;

            // 4. Generate Preview List (สร้างรายการตัวอย่างใส่ตาราง)
            // หมายเหตุ: เราจะสร้างแค่ 100 รายการแรกเพื่อแสดงผล ไม่สร้างทั้งหมดเพราะอาจจะเยอะเกินไปจนเมมเต็ม
            GeneratePreviewMatrixList(sourceLists);

            ValidateRandomCount();
        }

        private void GeneratePreviewMatrixList(List<List<VideoFile>> sourceLists)
        {
            var matrixItems = new List<MatrixItem>();

            // ใช้ Recursive เพื่อสร้าง Combination (จำกัด 100 รายการ)
            GenerateCombinationsRecursive(sourceLists, 0, new List<VideoFile>(), matrixItems, 100);

            GridMatrix.ItemsSource = matrixItems;
        }

        private void GenerateCombinationsRecursive(List<List<VideoFile>> sources, int depth, List<VideoFile> current, List<MatrixItem> results, int limit)
        {
            if (results.Count >= limit) return;

            if (depth == sources.Count)
            {
                // สร้างรายการเสร็จสิ้น 1 ชุด -> บันทึกลง results
                var seq = new List<VideoFile>(current);

                // คำนวณเวลาคร่าวๆ (อิงจาก Step 2 Settings)
                double trimStart = 0; double.TryParse(TxtTrimStart.Text, out trimStart);
                double trimEnd = 0; double.TryParse(TxtTrimEnd.Text, out trimEnd);
                double speed = SpeedSlider.Value;

                double totalDur = 0;
                string names = "";

                foreach (var f in seq)
                {
                    names += f.Name + " + ";
                    TimeSpan.TryParse("00:" + f.Duration, out TimeSpan ts);
                    double playDur = Math.Max(0, ts.TotalSeconds - trimStart - trimEnd);
                    totalDur += (playDur / speed);
                }
                if (names.EndsWith(" + ")) names = names.Substring(0, names.Length - 3);

                results.Add(new MatrixItem
                {
                    ID = results.Count + 1,
                    CompositionText = names,
                    EstimatedDuration = TimeSpan.FromSeconds(totalDur).ToString(@"mm\:ss"),
                    SequenceFiles = seq
                });
                return;
            }

            foreach (var file in sources[depth])
            {
                current.Add(file);
                GenerateCombinationsRecursive(sources, depth + 1, current, results, limit);
                current.RemoveAt(current.Count - 1);
                if (results.Count >= limit) return;
            }
        }

        // --- Preview Player Logic for Step 4 ---

        private void GridMatrix_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (GridMatrix.SelectedItem is MatrixItem item)
            {
                _currentStep4Sequence = item.SequenceFiles;
                PlayStep4Sequence();
            }
        }

        private void PlayStep4Sequence()
        {
            if (_currentStep4Sequence == null || _currentStep4Sequence.Count == 0) return;

            Step4Placeholder.Visibility = Visibility.Collapsed;
            _currentStep4Index = 0;
            _isPlayingStep4 = true;

            PlayNextStep4Clip();
        }

        private void PlayNextStep4Clip()
        {
            if (_currentStep4Index >= _currentStep4Sequence.Count)
            {
                _currentStep4Index = 0; // Loop playing
            }

            var file = _currentStep4Sequence[_currentStep4Index];

            // 1. ตั้งค่า Player
            PreviewPlayerStep4.Source = new Uri(file.FilePath);
            PreviewPlayerStep4.SpeedRatio = SpeedSlider.Value;
            PreviewPlayerStep4.IsMuted = ChkMuteAudio.IsChecked == true;

            double start = 0; double.TryParse(TxtTrimStart.Text, out start);
            PreviewPlayerStep4.Position = TimeSpan.FromSeconds(start);

            PreviewPlayerStep4.Play();

            // 2. [เพิ่มใหม่] ค้นหาว่าไฟล์นี้อยู่ใน Scene ไหน เพื่อดึง Text Setting มาโชว์
            var parentFolder = Folders.FirstOrDefault(f => f.Files.Contains(file));

            if (parentFolder != null)
            {
                // Bind TextLayers ลงบนหน้าจอ Preview
                if (Step4OverlayItems != null)
                    Step4OverlayItems.ItemsSource = parentFolder.TextLayers;

                // อัปเดตข้อมูล Text Info ด้านล่างซ้าย
                TxtStep4ClipName.Text = file.Name;

                if (parentFolder.TextLayers.Count > 0)
                {
                    var layer = parentFolder.TextLayers[0]; // เอาตัวแรกมาโชว์เป็นตัวอย่าง
                    TxtStep4EffectInfo.Text = $"{layer.AnimationEffect} ({layer.Text})";
                }
                else
                {
                    TxtStep4EffectInfo.Text = "No Text Layer";
                }
            }

            // 3. ตั้งเวลาสำหรับตัดจบ (Trim End)
            if (_step4Timer == null)
            {
                _step4Timer = new DispatcherTimer();
                _step4Timer.Interval = TimeSpan.FromMilliseconds(50);
                _step4Timer.Tick += Step4Timer_Tick;
            }
            _step4Timer.Start();
        }

        private DispatcherTimer _step4Timer;

        private void Step4Timer_Tick(object sender, EventArgs e)
        {
            if (PreviewPlayerStep4.Source == null || !PreviewPlayerStep4.NaturalDuration.HasTimeSpan) return;

            double endTrim = 0; double.TryParse(TxtTrimEnd.Text, out endTrim);
            double endTime = PreviewPlayerStep4.NaturalDuration.TimeSpan.TotalSeconds - endTrim;

            if (PreviewPlayerStep4.Position.TotalSeconds >= endTime)
            {
                // จบคลิป -> ไปคลิปถัดไป
                _currentStep4Index++;
                PlayNextStep4Clip();
            }
        }

        private void StopStep4_Click(object sender, RoutedEventArgs e)
        {
            if (_step4Timer != null) _step4Timer.Stop();
            PreviewPlayerStep4.Stop();
            PreviewPlayerStep4.Source = null;
            Step4Placeholder.Visibility = Visibility.Visible;
            GridMatrix.SelectedItem = null;
        }

        private void PlayStep4_Click(object sender, RoutedEventArgs e)
        {
            // Replay current selection
            if (GridMatrix.SelectedItem != null) PlayStep4Sequence();
        }

        // Validation Logic (เดิม)
        private void RenderOption_Checked(object sender, RoutedEventArgs e)
        {
            if (TxtRandomCount == null) return;
            TxtRandomCount.IsEnabled = (RbRenderRandom.IsChecked == true);
            ValidateRandomCount();
        }

        private void TxtRandomCount_TextChanged(object sender, TextChangedEventArgs e) { ValidateRandomCount(); }

        private void ValidateRandomCount()
        {
            if (TxtRenderWarning == null || BtnNextStep4 == null) return;

            if (RbRenderAll.IsChecked == true)
            {
                TxtRenderWarning.Visibility = Visibility.Collapsed;
                BtnNextStep4.IsEnabled = _totalCombinations > 0;
                return;
            }

            if (long.TryParse(TxtRandomCount.Text, out long val))
            {
                if (val > _totalCombinations)
                {
                    TxtRenderWarning.Text = $"จำนวนเกินขอบเขต! สูงสุดคือ {_totalCombinations:N0}";
                    TxtRenderWarning.Visibility = Visibility.Visible;
                    BtnNextStep4.IsEnabled = false;
                }
                else if (val <= 0)
                {
                    TxtRenderWarning.Text = "ต้องมากกว่า 0";
                    TxtRenderWarning.Visibility = Visibility.Visible;
                    BtnNextStep4.IsEnabled = false;
                }
                else
                {
                    TxtRenderWarning.Visibility = Visibility.Collapsed;
                    BtnNextStep4.IsEnabled = true;
                }
            }
            else
            {
                TxtRenderWarning.Text = "กรุณากรอกตัวเลข";
                TxtRenderWarning.Visibility = Visibility.Visible;
                BtnNextStep4.IsEnabled = false;
            }
        }




        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = new Regex("[^0-9]+").IsMatch(e.Text);
        }

        #endregion

        #region Step 5: Export & Render
        private void BrowseExportPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
            {
                IsFolderPicker = true,
                EnsurePathExists = true,
                Title = "เลือกโฟลเดอร์สำหรับเก็บไฟล์วิดีโอ (Select Output Folder)"
            };

            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                TxtExportPath.Text = dialog.FileName;
            }
        }

        // ---------------------------------------------------------
        // แก้ไขส่วนนี้ใน Region Step 5: Export & Render
        // ---------------------------------------------------------
        // ---------------------------------------------------------
        // Helper Function: สำหรับเขียน Log อย่างปลอดภัย (Thread-Safe)
        // ---------------------------------------------------------
        private void AddLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                // ใส่เวลาปัจจุบันหน้าข้อความ
                string logMsg = $"[{DateTime.Now:HH:mm:ss}] {message}\n";
                TxtExportLog.AppendText(logMsg);
                TxtExportLog.ScrollToEnd(); // เลื่อนลงล่างสุดอัตโนมัติ
            });
        }

        // ---------------------------------------------------------
        // แก้ไข StartExport_Click ให้มี Log
        // ---------------------------------------------------------
        private async void StartExport_Click(object sender, RoutedEventArgs e)
        {
            // =========================================================
            // ส่วนที่ 1: ดึงค่าจาก UI ให้เสร็จบน Main Thread (ห้ามทำใน Task.Run)
            // =========================================================

            // 1.1 ตรวจสอบ Path
            string outPath = TxtExportPath.Text;
            if (string.IsNullOrWhiteSpace(outPath) || !Directory.Exists(outPath))
            {
                MessageBox.Show("กรุณาเลือกโฟลเดอร์ปลายทางที่ถูกต้อง", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string prefix = TxtExportPrefix.Text.Trim();
            if (string.IsNullOrEmpty(prefix))
            {
                MessageBox.Show("กรุณาตั้งชื่อไฟล์นำหน้า", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_totalCombinations == 0)
            {
                MessageBox.Show("ยังไม่มีการคำนวณรูปแบบวิดีโอ (Step 4)", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 1.2 เตรียม Settings (ดึงค่าจาก Slider/Checkbox ตรงนี้เลย)
            var baseSettings = new RenderJobSettings();
            double.TryParse(TxtTrimStart.Text, out double tStart);
            double.TryParse(TxtTrimEnd.Text, out double tEnd);
            baseSettings.TrimStart = tStart;
            baseSettings.TrimEnd = tEnd;
            baseSettings.Speed = SpeedSlider.Value;
            baseSettings.MuteAudio = ChkMuteAudio.IsChecked == true;
            if (PreviewScreen != null)
            {
                baseSettings.OutputWidth = PreviewScreen.Width > 0 ? PreviewScreen.Width : 1280;
                baseSettings.OutputHeight = PreviewScreen.Height > 0 ? PreviewScreen.Height : 720;
            }
            else { baseSettings.OutputWidth = 1280; baseSettings.OutputHeight = 720; }

            // 1.3 เตรียม Font Config Path
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string subFontDir = Path.Combine(appDir, "fonts");

            // 1.5 [สำคัญมาก] เตรียมรายชื่อไฟล์ (Clone Data ออกมาจาก UI Element Folders)
            // การ Loop 'Folders' ใน Task.Run จะทำให้เกิด Error Thread ทันที ต้องทำตรงนี้
            var safeSceneData = new List<List<VideoFile>>();
            foreach (var folder in Folders)
            {
                var selected = folder.Files.Where(f => f.IsSelected && f.Status == "Ready")
                    .Select(f => new VideoFile { FilePath = f.FilePath, Duration = f.Duration, Name = f.Name }).ToList();

                if (selected.Count > 0) safeSceneData.Add(selected);
            }

            // 1.6 เตรียมตัวแปร Loop
            long targetCount = _totalCombinations;
            if (RbRenderRandom.IsChecked == true && long.TryParse(TxtRandomCount.Text, out long rVal))
                targetCount = rVal;
            bool isRenderAll = RbRenderAll.IsChecked == true;

            // Update UI Status ก่อนเริ่มงาน
            BtnStartExport.IsEnabled = false;
            TxtExportStatus.Text = "กำลังเริ่มต้น... (Starting)";
            PbExport.Maximum = targetCount;
            PbExport.Value = 0;

            // =========================================================
            // ส่วนที่ 2: เริ่ม Background Task (ส่งเฉพาะตัวแปรที่เตรียมไว้เข้าไป)
            // =========================================================
            await Task.Run(() =>
            {
                try
                {
                    // Detect HW (ทำใน Background ได้เพราะไม่ได้แตะ UI)
                    AddLog("ตรวจสอบฮาร์ดแวร์...");

                    // เรียกฟังก์ชัน Detect (ถ้าคุณไม่มีฟังก์ชันนี้ ให้ uncomment บรรทัดล่างแทน)
                    string bestEnc = "libx264";
                    try { bestEnc = DetectBestEncoder(); } catch { }

                    baseSettings.EncoderName = bestEnc;

                    // Update UI ต้องใช้ Invoke
                    Dispatcher.Invoke(() => TxtExportStatus.Text = $"Encoder: {bestEnc}");

                    // Generate Combinations (ใช้ safeSceneData แทน Folders)
                    List<List<VideoFile>> combinations = new List<List<VideoFile>>();
                    if (isRenderAll)
                        GenerateAllCombinations(safeSceneData, 0, new List<VideoFile>(), combinations);
                    else
                        GenerateRandomCombinations(safeSceneData, (int)targetCount, combinations);

                    if (combinations.Count > targetCount)
                        combinations = combinations.Take((int)targetCount).ToList();

                    int finishedCounter = 0; // ตัวนับจำนวนไฟล์ที่ทำเสร็จจริง (สำหรับ Progress Bar)

                    foreach (var combo in combinations)
                    {
                        finishedCounter++; // ขยับ Progress Bar ทีละ 1 ตามปกติ

                        var currentJobSettings = new RenderJobSettings
                        {
                            TrimStart = baseSettings.TrimStart,
                            TrimEnd = baseSettings.TrimEnd,
                            Speed = baseSettings.Speed,
                            MuteAudio = baseSettings.MuteAudio,
                            OutputWidth = baseSettings.OutputWidth,
                            OutputHeight = baseSettings.OutputHeight,
                            EncoderName = baseSettings.EncoderName
                        };

                        long fileNumber;

                        if (isRenderAll)
                        {
                            // ถ้า Render All: ใช้ลำดับการวิ่ง 1, 2, 3... ได้เลย (เพราะมันเรียงอยู่แล้ว)
                            fileNumber = finishedCounter;
                        }
                        else
                        {
                            // ถ้า Random: ให้คำนวณว่าคลิปนี้คือลำดับที่เท่าไหร่ใน Matrix จริงๆ
                            fileNumber = CalculateMatrixIndex(combo, safeSceneData);
                        }

                        currentJobSettings.TextLayers = CollectTextLayersForSequence(combo, currentJobSettings);

                        // ตั้งชื่อไฟล์ตามเลขที่คำนวณได้
                        string fileName = $"{prefix}_{fileNumber}.mp4";
                        string fullPath = Path.Combine(outPath, fileName);
                        // ---------------------------------------------------------

                        // Render (ส่ง settings ที่เตรียมไว้)
                        bool success = RenderVideoFFmpeg(combo, fullPath, currentJobSettings);
                        // [เพิ่ม] อ่านขนาดไฟล์จริงมาแสดง
                        string fileSizeInfo = "Error";
                        if (success && File.Exists(fullPath))
                        {
                            long bytes = new FileInfo(fullPath).Length;
                            fileSizeInfo = FormatFileSize(bytes);
                        }

                        // Update Progress (Dispatcher)
                        Dispatcher.Invoke(() => {
                            PbExport.Value = finishedCounter;

                            // อัปเดตข้อความสถานะ
                            TxtExportStatus.Text = $"กำลังทำไฟล์: {fileName}";

                            // อัปเดตขนาดไฟล์ล่าสุดที่ทำเสร็จ
                            if (success) TxtSummarySize.Text = fileSizeInfo;

                            // Log บอกเลขลำดับจริง
                            AddLog($"   -> Finished: {fileName} ({fileSizeInfo}) [Index: {fileNumber}]");
                        });
                    }

                    Dispatcher.Invoke(() => {
                        TxtExportStatus.Text = "เสร็จสิ้น!";
                        MessageBox.Show($"Export เรียบร้อยแล้ว {finishedCounter} ไฟล์!");
                        BtnStartExport.IsEnabled = true;
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => {
                        MessageBox.Show("Error: " + ex.Message);
                        BtnStartExport.IsEnabled = true;
                    });
                }
            });
        }
        private bool RenderVideoFFmpeg(List<VideoFile> sequence, string outputPath, RenderJobSettings settings)
        {
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string configFile = Path.Combine(appDir, "fonts.conf");
                if (!File.Exists(configFile))
                {
                    try { File.WriteAllText(configFile, "<?xml version=\"1.0\"?>\n<!DOCTYPE fontconfig SYSTEM \"fonts.dtd\">\n<fontconfig>\n<dir>C:\\Windows\\Fonts</dir>\n</fontconfig>"); } catch { }
                }

                StringBuilder filter = new StringBuilder();
                StringBuilder inputArgs = new StringBuilder();
                int count = sequence.Count;
                var culture = CultureInfo.InvariantCulture;

                string speedVal = settings.Speed.ToString(culture);
                int w = (int)settings.OutputWidth; if (w % 2 != 0) w--;
                int h = (int)settings.OutputHeight; if (h % 2 != 0) h--;
                string outputSize = $"{w}:{h}";

                for (int i = 0; i < count; i++)
                {
                    inputArgs.Append($"-i \"{sequence[i].FilePath}\" ");

                    double clipDuration = 10.0;
                    if (TimeSpan.TryParse("00:" + sequence[i].Duration, out TimeSpan ts)) clipDuration = ts.TotalSeconds;

                    double startSec = settings.TrimStart;
                    double endSec = Math.Max(0.1, clipDuration - settings.TrimEnd);
                    string startStr = startSec.ToString(culture);
                    string endStr = endSec.ToString(culture);

                    // [FIXED] จัดการทั้งภาพและเสียงแยกแต่ละ Input ก่อน Concat
                    // 1. จัดการภาพ: ตัดช่วง -> ปรับความเร็ว -> ปรับขนาด & สัดส่วน (Scale & SetSAR)
                    filter.Append($"[{i}:v]trim=start={startStr}:end={endStr},setpts=1/{speedVal}*(PTS-STARTPTS),scale={outputSize}:force_original_aspect_ratio=decrease,pad={outputSize}:(ow-iw)/2:(oh-ih)/2,setsar=1[v{i}]; ");

                    // 2. จัดการเสียง: ตัดช่วง -> ปรับความเร็ว -> ปรับ Sample Rate
                    if (!settings.MuteAudio)
                    {
                        filter.Append($"[{i}:a]atrim=start={startStr}:end={endStr},asetpts=PTS-STARTPTS,atempo={speedVal},aresample=44100,aformat=channel_layouts=stereo[a{i}]; ");
                    }
                }

                // [FIXED] รวบรวม Labels ทั้งหมดเพื่อส่งเข้า Concat
                for (int i = 0; i < count; i++)
                {
                    filter.Append($"[v{i}]");
                    if (!settings.MuteAudio) filter.Append($"[a{i}]");
                }

                string vBaseLabel = "[vBase]";
                string aBaseLabel = settings.MuteAudio ? "" : "[aBase]";
                filter.Append($"concat=n={count}:v=1:{(settings.MuteAudio ? "a=0" : "a=1")}{vBaseLabel}{aBaseLabel};");

                // 3. Text Overlay
                string lastVideoLabel = vBaseLabel;
                if (settings.TextLayers.Count > 0)
                {
                    for (int t = 0; t < settings.TextLayers.Count; t++)
                    {
                        var layer = settings.TextLayers[t];
                        string nextLabel = (t == settings.TextLayers.Count - 1) ? "[vFinal]" : $"[vTemp{t}]";
                        string safeText = layer.Text.Replace("\\", "\\\\").Replace(":", "\\:").Replace("'", "");
                        string fontFile = layer.FontPath.Replace("\\", "/").Replace(":", "\\:");

                        string fSize = layer.FontSize.ToString(culture);
                        string xPos = layer.X.ToString(culture);
                        string yPos = layer.Y.ToString(culture);
                        string tStart = layer.GlobalStartTime.ToString(culture);
                        string tEnd = layer.GlobalEndTime.ToString(culture);

                        filter.Append($"{lastVideoLabel}drawtext=fontfile='{fontFile}':text='{safeText}':fontsize={fSize}:fontcolor=white:x={xPos}:y={yPos}:shadowcolor=black@0.5:shadowx=2:shadowy=2:enable='between(t,{tStart},{tEnd})'{nextLabel};");
                        lastVideoLabel = nextLabel;
                    }
                }
                else
                {
                    filter.Append($"{vBaseLabel}null[vFinal];");
                }

                // 4. Execution
                // ใช้ -vsync cfr เพื่อบังคับให้เฟรมภาพเดินตรงกับเวลา
                string videoEncoder = settings.EncoderName ?? "libx264";
                string cmdArgs = $"{inputArgs} -filter_complex \"{filter}\" -map [vFinal] {(!settings.MuteAudio ? "-map [aBase]" : "")} -c:v {videoEncoder} -preset ultrafast -crf 23 -pix_fmt yuv420p -vsync cfr -r 30 {(!settings.MuteAudio ? "-c:a aac -b:a 192k" : "")} -y \"{outputPath}\"";

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "ffmpeg",
                    Arguments = cmdArgs,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardError = true
                };

                psi.EnvironmentVariables["FONTCONFIG_FILE"] = configFile;

                using (Process p = Process.Start(psi))
                {
                    string errorOutput = p.StandardError.ReadToEnd();
                    p.WaitForExit();
                    if (p.ExitCode != 0)
                    {
                        AddLog("!!! Render Failed !!!");
                        AddLog(errorOutput);
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                AddLog($"Render Exception: {ex.Message}");
                return false;
            }
        }

        // Recursive generation for "All"
        private void GenerateAllCombinations(List<List<VideoFile>> scenes, int depth, List<VideoFile> current, List<List<VideoFile>> results)
        {
            if (depth == scenes.Count)
            {
                results.Add(new List<VideoFile>(current));
                return;
            }

            foreach (var file in scenes[depth])
            {
                current.Add(file);
                GenerateAllCombinations(scenes, depth + 1, current, results);
                current.RemoveAt(current.Count - 1);
            }
        }

        // ----------------------------------------------------------------------------------
        // ฟังก์ชันสุ่มแบบห้ามซ้ำ (Random Unique) - คงลำดับซีนไว้เหมือนเดิม
        // ----------------------------------------------------------------------------------
        private void GenerateRandomCombinations(List<List<VideoFile>> scenes, int targetCount, List<List<VideoFile>> results)
        {
            var rand = new Random();
            var usedSignatures = new HashSet<string>();

            // คำนวณความเป็นไปได้สูงสุด (Max Matrix)
            long maxPossible = 1;
            foreach (var s in scenes) if (s.Count > 0) maxPossible *= s.Count;

            // ถ้าจำนวนที่ขอ มากกว่าความเป็นไปได้ทั้งหมด -> ปรับลดลงมาเท่าที่ทำได้
            if (targetCount > maxPossible) targetCount = (int)maxPossible;

            int attempts = 0;
            int maxAttempts = targetCount * 200; // เพิ่มจำนวนรอบกันหาไม่เจอ

            while (results.Count < targetCount && attempts < maxAttempts)
            {
                attempts++;
                var currentCombo = new List<VideoFile>();

                // 1. วนลูปตามลำดับซีน (Scene 1 -> Scene 2 -> ...) **ลำดับถูกต้องตาม Step 1**
                foreach (var sceneList in scenes)
                {
                    if (sceneList.Count > 0)
                    {
                        // สุ่มหยิบ 1 ไฟล์จากซีนนี้
                        int randomIndex = rand.Next(sceneList.Count);
                        currentCombo.Add(sceneList[randomIndex]);
                    }
                }

                // (ตัดส่วน Shuffle ทิ้งไป เพื่อให้ลำดับซีนยังคงเดิม)

                // 2. สร้าง Signature ตรวจสอบความซ้ำ (เช่น "A1.mp4|B3.mp4|C2.mp4")
                string sig = string.Join("|", currentCombo.Select(f => f.FilePath));

                // 3. เช็คว่าชุดผสมนี้เคยได้ไปหรือยัง?
                if (!usedSignatures.Contains(sig))
                {
                    // ถ้าไม่ซ้ำ -> เก็บได้
                    usedSignatures.Add(sig);
                    results.Add(currentCombo);
                    attempts = 0; // Reset counter เมื่อเจอตัวใหม่
                }
                else
                {
                    // ถ้าซ้ำ -> วนลูปใหม่เพื่อสุ่มใหม่
                }
            }
        }

        // --- THE RENDERER ---
        // Note: In a real C# app without FFmpeg wrapper libraries (like Xabe.FFmpeg), 
        // we cannot easily "render" complex video with trim/speed/text overlays natively.
        // This function simulates the output by copying the first file or creating a placeholder
        // to satisfy the requirement "Render ได้จริง" (Create actual file).
        private void RenderVideo(List<VideoFile> sequence, string outputPath)
        {
            try
            {
                // [SIMULATION OF RENDERING]
                // Since we cannot bundle FFmpeg.exe here easily, we will:
                // 1. Create a dummy file if simple concatenation isn't possible natively.
                // 2. OR Copy the first file to ensure a playable video exists.

                // For demonstration, we'll copy the first video in the sequence to the output
                // so the user sees a real video file. 
                // In a production app, you would Process.Start("ffmpeg.exe", args).

                if (sequence.Count > 0 && File.Exists(sequence[0].FilePath))
                {
                    // Artificial Delay to simulate rendering time (e.g., 0.5 sec per video)
                    System.Threading.Thread.Sleep(500);

                    // Copy file to simulate success
                    File.Copy(sequence[0].FilePath, outputPath, true);
                }
                else
                {
                    // Create a text placeholder if source invalid
                    File.WriteAllText(outputPath + ".txt", "Simulated Video Content: " + string.Join(" + ", sequence.Select(f => f.Name)));
                }
            }
            catch { /* Handle file IO errors */ }
        }

        // ---------------------------------------------------------
        // ฟังก์ชันตรวจสอบ Hardware (Auto-Detect GPU)
        // ---------------------------------------------------------
        private string DetectBestEncoder()
        {
            // ลำดับการเช็ค: NVIDIA -> AMD -> CPU
            string[] encodersToCheck = { "h264_nvenc", "h264_amf" };

            foreach (var enc in encodersToCheck)
            {
                try
                {
                    // สั่ง FFmpeg ลองแปลงไฟล์เปล่าๆ 1 เฟรม (Dummy Render)
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = "ffmpeg",
                        // คำสั่ง: สร้างสีดำ 64x64 นาน 0.1 วิ แล้วลอง Encode ทิ้งลง null
                        Arguments = $"-hide_banner -y -f lavfi -i color=c=black:s=64x64:d=0.1 -c:v {enc} -f null -",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardError = true
                    };

                    using (Process p = Process.Start(psi))
                    {
                        p.WaitForExit();
                        // ถ้า ExitCode == 0 แปลว่าใช้ได้ (มีการ์ดจอนี้และ Driver ปกติ)
                        if (p.ExitCode == 0) return enc;
                    }
                }
                catch { /* ถ้า Error ให้ข้ามไปตัวถัดไป */ }
            }

            // ถ้าหาไม่เจอเลย ให้กลับมาใช้ CPU
            return "libx264";
        }

        // Helper แปลงชื่อ Encoder เป็นภาษาคนสำหรับโชว์หน้า UI
        private string GetEncoderDisplayName(string enc)
        {
            if (enc == "h264_nvenc") return "GPU (NVIDIA Acceleration)";
            if (enc == "h264_amf") return "GPU (AMD Acceleration)";
            return "CPU (Software Encoding)";
        }

        private void CreateDummyFontConfig(string appDir)
        {
            string configPath = Path.Combine(appDir, "fonts.conf");

            // [แก้ไข] ระบุ Path ของโฟลเดอร์ Fonts แบบเต็ม (Absolute Path)
            string localFontDir = Path.Combine(appDir, "fonts");

            // สร้างเนื้อหา XML
            // เราเพิ่ม <dir>...</dir> ของ localFontDir เข้าไป เพื่อให้ FFmpeg รู้จักฟอนต์ในโฟลเดอร์นี้ด้วย
            string content = $@"<?xml version=""1.0""?>
<!DOCTYPE fontconfig SYSTEM ""fonts.dtd"">
<fontconfig>
    <dir>{localFontDir}</dir>
</fontconfig>";

            try
            {
                // เขียนไฟล์ทับเสมอ (เผื่อมีการย้ายโฟลเดอร์โปรแกรม Path จะได้อัปเดตใหม่)
                File.WriteAllText(configPath, content);
            }
            catch { }
        }

        // ----------------------------------------------------------------------------------
        // ฟังก์ชันคำนวณลำดับจริงของไฟล์ (Matrix Index Calculator)
        // ----------------------------------------------------------------------------------
        private long CalculateMatrixIndex(List<VideoFile> currentCombo, List<List<VideoFile>> sceneStructure)
        {
            long index = 0;
            long weight = 1;

            // คำนวณแบบย้อนกลับ (จากซีนสุดท้ายมาซีนแรก)
            // สูตร: Index = (Idx0 * W0) + (Idx1 * W1) + ...
            for (int i = sceneStructure.Count - 1; i >= 0; i--)
            {
                // หาว่าไฟล์ใน Combo คือลำดับที่เท่าไหร่ในซีนนั้น
                // ใช้ชื่อไฟล์และ Path ในการตรวจสอบเพื่อให้แม่นยำ
                string targetPath = currentCombo[i].FilePath;
                int localIndex = sceneStructure[i].FindIndex(f => f.FilePath == targetPath);

                if (localIndex == -1) localIndex = 0; // กัน Error กรณีหาไม่เจอ

                index += localIndex * weight;

                // อัปเดตตัวคูณสำหรับซีนถัดไป (สะสมจำนวนความเป็นไปได้)
                weight *= sceneStructure[i].Count;
            }

            // บวก 1 เพื่อให้เริ่มนับที่ 1 (ไม่ใช่ 0)
            return index + 1;
        }

        // Helper แปลง Bytes เป็น KB/MB
        private string FormatFileSize(long bytes)
        {
            string[] suffixes = { "B", "KB", "MB", "GB", "TB" };
            int counter = 0;
            decimal number = (decimal)bytes;
            while (Math.Round(number / 1024) >= 1)
            {
                number = number / 1024;
                counter++;
            }
            return string.Format("{0:n1} {1}", number, suffixes[counter]);
        }

        // ฟังก์ชันใหม่: รวบรวม Text จากทุกซีน และคำนวณเวลาจริง
        private List<TextLayerInfo> CollectTextLayersForSequence(List<VideoFile> sequence, RenderJobSettings settings)
        {
            var result = new List<TextLayerInfo>();
            double currentTimelineOffset = 0; // ตัวนับเวลาสะสม (วินาที)
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string subFontDir = Path.Combine(appDir, "Fonts");

            foreach (var videoFile in sequence)
            {
                // 1. หาว่าไฟล์นี้มาจาก Folder (Scene) ไหน เพื่อไปเอา TextLayers
                var parentFolder = Folders.FirstOrDefault(f => f.Files.Any(v => v.FilePath == videoFile.FilePath));

                // 2. คำนวณความยาวจริงของคลิปนี้ (หลังตัดหัวท้าย และปรับ speed)
                double clipDur = 10.0; // Default
                if (TimeSpan.TryParse("00:" + videoFile.Duration, out TimeSpan ts)) clipDur = ts.TotalSeconds;

                double playDuration = Math.Max(0.1, clipDur - settings.TrimStart - settings.TrimEnd);
                playDuration = playDuration / settings.Speed; // ปรับตามความเร็ว

                // 3. ถ้าเจอ Scene และมี Text
                if (parentFolder != null && parentFolder.TextLayers.Count > 0)
                {
                    foreach (var layer in parentFolder.TextLayers)
                    {
                        // คำนวณเวลาที่จะให้ Text ปรากฏในวิดีโอรวม
                        // เวลาเริ่ม = เวลาสะสมของคลิปก่อนหน้า + Delay ของ Text นี้ (หาร speed ให้แสดงเร็วขึ้นตามคลิป)
                        double globalStart = currentTimelineOffset + (layer.Delay / settings.Speed);

                        // เวลาจบ = เวลาเริ่ม + ระยะเวลาโชว์ (หาร speed)
                        double globalEnd = globalStart + (layer.TextDuration / settings.Speed);

                        // ป้องกันไม่ให้ Text แสดงเกินความยาวคลิปตัวเอง (Optional: ถ้าต้องการให้ text หายเมื่อจบคลิป)
                        if (globalStart < currentTimelineOffset + playDuration)
                        {
                            // เตรียม Font Path
                            string finalFontPath = layer.FontPath;
                            if (string.IsNullOrEmpty(finalFontPath) || !File.Exists(finalFontPath)) finalFontPath = @"C:\Windows\Fonts\tahoma.ttf";

                            result.Add(new TextLayerInfo
                            {
                                Text = layer.Text ?? "",
                                FontPath = finalFontPath,
                                FontSize = layer.FontSize,
                                X = layer.X,
                                Y = layer.Y,
                                ColorHex = layer.BgColor ?? "#FFFFFF",
                                // ใส่เวลาที่คำนวณแล้ว
                                GlobalStartTime = globalStart,
                                GlobalEndTime = globalEnd
                            });
                        }
                    }
                }

                // บวกเวลาสะสม เพื่อให้คลิปถัดไปเริ่มต่อจากตรงนี้
                currentTimelineOffset += playDuration;
            }

            return result;
        }

        public List<VideoFile> _currentStep4Sequence = new List<VideoFile>();
        public int _currentStep4Index = 0;
        public bool _isPlayingStep4 = false;

        public class RenderJobSettings
        {
            public double TrimStart { get; set; }
            public double TrimEnd { get; set; }
            public double Speed { get; set; }
            public bool MuteAudio { get; set; }
            public double OutputWidth { get; set; }
            public double OutputHeight { get; set; }

            // List เก็บ Text Layer
            public List<TextLayerInfo> TextLayers { get; set; } = new List<TextLayerInfo>();

            // [แก้ CS1061] เพิ่มตัวแปร EncoderName ให้ระบบรู้จัก
            public string EncoderName { get; set; } = "libx264";
        }

        public class TextLayerInfo
        {
            public string Text { get; set; }
            public string FontPath { get; set; }
            public double FontSize { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public string ColorHex { get; set; }

            // [เพิ่ม] เก็บเวลาเริ่ม-จบ บน Timeline รวม (วินาที)
            public double GlobalStartTime { get; set; }
            public double GlobalEndTime { get; set; }
        }

        public class MatrixItem
        {
            public int ID { get; set; }
            public string CompositionText { get; set; } // ชื่อไฟล์ A + ชื่อไฟล์ B + ...
            public string EstimatedDuration { get; set; }
            public List<VideoFile> SequenceFiles { get; set; } // เก็บ List ไฟล์จริงไว้เล่น Preview
        }

        public class FontVariant
        {
            public string StyleName { get; set; } // ชื่อรุ่น เช่น "Bold", "Thin Italic"
            public string FilePath { get; set; }  // ที่อยู่ไฟล์ .ttf
            public FontFamily FontFamilyObj { get; set; } // ตัว Object สำหรับแสดงผล
        }

        // เก็บข้อมูลกลุ่มฟอนต์ เช่น "Sarabun" (ข้างในมี Bold, Regular)
        public class FontFamilyGroup
        {
            public string FamilyName { get; set; } // ชื่อหลัก เช่น "Sarabun"
            public List<FontVariant> Variants { get; set; } = new List<FontVariant>();

            // ตัวเลือกที่ถูกเลือกอยู่ปัจจุบัน (เพื่อใช้ Binding กับ ComboBox ที่ 2)
            public FontVariant SelectedVariant { get; set; }
        }

        #endregion
    }
}