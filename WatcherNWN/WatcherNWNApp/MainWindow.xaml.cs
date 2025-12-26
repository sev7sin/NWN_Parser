using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Media;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using Microsoft.Data.Sqlite;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using System.Windows.Threading;
#pragma warning disable CS8629

namespace WatcherNWNApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileMonitor? _fileMonitor;
        private LogProcessingContentProcessor? _logProcessor;
        private readonly LogParser _logParser = new();
        private readonly LogStatsTracker _statsTracker = new();
        private readonly string _defaultCsvPath;
        private string? _csvOutputPath;
        private string? _lastSelectedEntity;
        private readonly List<CombatSession> _combatSessions = new List<CombatSession>();
        private CombatSession? _currentCombat;
        private CombatSession? _selectedCombat;
        private int _combatCounter = 0;
        private readonly PlayerDetector _playerDetector = new PlayerDetector();
        private string? _playerNameOverride;
        private string? _autoPlayerGuess;
        private SocialStore _socialStore;
        private string? _selectedSocialName;
        private readonly ObservableCollection<ReminderItem> _reminders = new ObservableCollection<ReminderItem>();
        private readonly ObservableCollection<string> _recentTalkers = new ObservableCollection<string>();
        private readonly StatsRepository _statsRepo;
        private readonly string _reminderFilePath;
        private readonly DispatcherTimer _reminderSaveDebounce;
        private readonly DispatcherTimer _countdownTimer;
        private string? _selectedEntityForActions;
        private TimeSpan _timerRemaining = TimeSpan.Zero;
        private TimeSpan _timerInitial = TimeSpan.Zero;
        private TimeSpan _timerPreAlert = TimeSpan.Zero;
        private bool _timerPreAlertFired = false;
        private double _beepVolume = 0.5;
        private readonly MediaPlayer _beepPlayer = new MediaPlayer();
        private readonly string _beepSoundPath;
        private string _socialSearchTerm = string.Empty;
        private readonly string _windowSettingsPath;
        private readonly string _socialCorePath;
        private string _socialOverlayPath;
        private bool _loadingSocial = false;
        private readonly DispatcherTimer _socialNotesDebounce;
        private bool _socialNotesDirty = false;
        private bool _socialOverlayLoadedByUser = false;
        private readonly RollDedupeStore _rollDedupeStore;
        private string? _currentLogPath;
        private readonly IngestionStateStore _ingestionStore;
        private string? _lastPlayerNamePreference;



        public MainWindow()
        {
            InitializeComponent();
            
            _defaultCsvPath = Path.Combine(AppContext.BaseDirectory, "AggregatedStats.csv");
            _csvOutputPath = _defaultCsvPath;
            _socialCorePath = Path.Combine(AppContext.BaseDirectory, "SocialNames.json");
            _socialOverlayPath = Path.Combine(AppContext.BaseDirectory, "SocialProfiles.json");
            _socialStore = new SocialStore(_socialOverlayPath, _socialCorePath);
            _socialStore.Load();
            UpdateSocialFileLabel();
            UpdateSocialEditingEnabled();
            _windowSettingsPath = Path.Combine(AppContext.BaseDirectory, "WindowSettings.json");
            _statsRepo = new StatsRepository(Path.Combine(AppContext.BaseDirectory, "Stats.db"));
            _statsRepo.EnsureDatabase();
            _ingestionStore = new IngestionStateStore(Path.Combine(AppContext.BaseDirectory, "Settings", "ingestion_state.json"));
            _reminderFilePath = Path.Combine(AppContext.BaseDirectory, "Reminders.json");
            _reminderSaveDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _reminderSaveDebounce.Tick += ReminderSaveDebounce_Tick;
            LoadWindowSettings();
            LoadReminders();
            UpdateSocialProfileContext();
            _rollDedupeStore = RollDedupeStore.Load(Path.Combine(AppContext.BaseDirectory, "RollDedupe.json"));
            bool loadedFromDb = _statsTracker.LoadFromDatabase(_statsRepo);
            if (!loadedFromDb)
            {
                _statsTracker.LoadFromCsv(_csvOutputPath);
                _statsTracker.SaveToDatabase(_statsRepo);
            }
            ClearRollsOnce();
            txtStatus.Text = "Ready";
            RefreshUiFromStats();
            RefreshCombatUi();
            RefreshCombatTotalsUi();
            RefreshPlayerUi();
            RefreshCombatSelection();
            RefreshSocialUi();
            reminderItems.ItemsSource = _reminders;
            UpdateReminderSummary();
            InitializeTimerOptions();
            recentTalkItems.ItemsSource = _recentTalkers;

            _countdownTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _countdownTimer.Tick += CountdownTimer_Tick;

            _socialNotesDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _socialNotesDebounce.Tick += SocialNotesDebounce_Tick;

            _beepSoundPath = EnsureBeepSoundFile();
            _beepPlayer.Open(new Uri(_beepSoundPath, UriKind.Absolute));
            _beepPlayer.Volume = _beepVolume;
            sliderVolume.Value = _beepVolume;
            txtVolumeValue.Text = $"{(int)(_beepVolume * 100)}%";

            // Initialize the file monitor with default settings
            _fileMonitor = new FileMonitor();
        }

        protected override void OnClosed(EventArgs e)
        {
            // Clean up resources
            FlushSocialNotesIfDirty();
            _fileMonitor?.StopMonitoring();
            _beepPlayer.Close();
            _countdownTimer.Stop();
            _reminderSaveDebounce.Stop();
            SaveReminders();
            _socialNotesDebounce.Stop();
            SaveWindowSettings();
            PersistCurrentDedupeState();
            _rollDedupeStore.Save();
            _ingestionStore.Save();
            base.OnClosed(e);
        }


        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                Title = "Select a text file to monitor"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string selectedFilePath = openFileDialog.FileName;
                StartMonitoring(selectedFilePath);
            }
        }

        private void StartMonitoring(string filePath)
        {
            try
            {
                // Stop any existing monitoring
                _fileMonitor?.StopMonitoring();

                // Reload known stats from the persistent CSV and clear displayed text for the new session
                _statsTracker.Reset();
                _statsTracker.LoadFromDatabase(_statsRepo);
                _currentLogPath = filePath;
                var dedupeState = _rollDedupeStore.GetOrCreate(filePath);
                _statsTracker.SetDedupeState(dedupeState);
                UpdateOverallDiceAverage();
                _combatSessions.Clear();
                _currentCombat = null;
                _selectedCombat = null;
                _combatCounter = 0;
            _playerDetector.Reset();
            _autoPlayerGuess = null;
            txtStatus.Text = "Monitoring...";
            RefreshUiFromStats();
            RefreshCombatUi();
                RefreshCombatTotalsUi();
                RefreshPlayerUi();
                RefreshCombatSelection();

                if (!string.IsNullOrWhiteSpace(_csvOutputPath))
                {
                    try
                    {
                _statsTracker.SaveToDatabase(_statsRepo);
                    }
                    catch (Exception)
                    {
                        // Ignore write errors here; will retry on next parse batch.
                    }
                }

                // Create processor with current paths/state
            _logProcessor = new LogProcessingContentProcessor(_logParser, _statsTracker, () => _statsTracker.SaveToDatabase(_statsRepo), OnStatsUpdated, OnAttackSeen, OnDamageSeen, OnHealSeen, OnAnyEvent, OnRawLine);
                
                // Set up new monitoring with the selected file path
            _fileMonitor?.StartMonitoring(filePath, OnFileChanged, _ingestionStore);
            UpdateSocialProfileContext();
                PersistCurrentDedupeState();
        }
            catch (Exception ex)
            {
                txtStatus.Text = "Error starting monitoring";
                MessageBox.Show($"Error starting monitoring: {ex.Message}");
            }
        }

        private void BtnClearAllDice_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Reset all recorded dice rolls and last-roll timestamp?", "Confirm reset", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            if (MessageBox.Show("This CANNOT be undone. Continue?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            _statsTracker.ClearAllRolls();
            _statsTracker.ClearDedupeState();
            _rollDedupeStore.ClearAll();
            _ingestionStore.ClearAll();
            _statsTracker.SaveToDatabase(_statsRepo);
            UpdateOverallDiceAverage();
            if (!string.IsNullOrWhiteSpace(_lastSelectedEntity))
            {
                ShowEntityDetails(_lastSelectedEntity);
            }
            txtStatus.Text = "Dice stats reset";
        }

        private void OnFileChanged(string content)
        {
            // Process the content using the configured processor
            _logProcessor?.ProcessContent(content);
        }

        private void CmbEntities_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? selected = cmbEntities.SelectedItem as string;
            _selectedEntityForActions = selected;
            UpdateEntityButtonsState();
            ShowEntityDetails(selected);
        }

        private void CmbPlayer_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ApplyPlayerOverrideFromUi();
        }

        private void CmbPlayer_LostKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e)
        {
            ApplyPlayerOverrideFromUi();
        }

        private void CmbSocial_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FlushSocialNotesIfDirty();
            string? selected = cmbSocial.SelectedItem as string;
            ShowSocialProfile(selected);
        }

        private void TxtSocialSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            _socialSearchTerm = txtSocialSearch.Text?.Trim() ?? string.Empty;
            RefreshSocialUi();
        }

        private void SocialCheckChanged(object sender, RoutedEventArgs e)
        {
            if (_loadingSocial)
                return;
            SaveSocialProfile(includeNotes: false);
        }

        private void TxtAlias_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_loadingSocial)
                return;
            SaveSocialProfile(includeNotes: false);
        }

        private void TxtSocialNotes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_loadingSocial)
                return;
            _socialNotesDirty = true;
            _socialNotesDebounce.Stop();
            _socialNotesDebounce.Start();
        }

        private void BtnSaveSocial_Click(object sender, RoutedEventArgs e)
        {
            SaveSocialProfile(includeNotes: true);
        }

        private void SocialNotesDebounce_Tick(object? sender, EventArgs e)
        {
            _socialNotesDebounce.Stop();
            if (_socialNotesDirty)
            {
                SaveSocialProfile(includeNotes: true);
            }
        }

        private void FlushSocialNotesIfDirty()
        {
            if (_socialNotesDirty && !string.IsNullOrWhiteSpace(_selectedSocialName))
            {
                SaveSocialProfile(includeNotes: true);
            }
            else if (_socialNotesDirty)
            {
                _socialNotesDebounce.Stop();
                _socialNotesDirty = false;
            }
        }

        private void BtnLoadSocial_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Social files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Load social data"
            };
            if (ofd.ShowDialog() == true)
            {
                LoadSocialFile(ofd.FileName, userInitiated: true);
            }
        }

        private void BtnNewSocial_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Social files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Create new social data",
                FileName = "SocialProfiles.json"
            };
            if (sfd.ShowDialog() == true)
            {
                LoadSocialFile(sfd.FileName, createIfMissing: true, userInitiated: true);
            }
        }

        private void BtnStartTimer_Click(object sender, RoutedEventArgs e)
        {
            int minutes = GetSelectedTimerMinutes();
            if (minutes <= 0)
            {
                minutes = 1;
            }

            _timerInitial = TimeSpan.FromMinutes(minutes);
            _timerRemaining = _timerInitial;
            _timerPreAlert = GetSelectedPreAlert();
            _timerPreAlertFired = false;
            _countdownTimer.Start();
            UpdateTimerDisplay();
        }

        private void BtnPauseTimer_Click(object sender, RoutedEventArgs e)
        {
            _countdownTimer.Stop();
        }

        private void BtnStopTimer_Click(object sender, RoutedEventArgs e)
        {
            _countdownTimer.Stop();
            _timerRemaining = TimeSpan.Zero;
            _timerInitial = TimeSpan.Zero;
            _timerPreAlert = TimeSpan.Zero;
            _timerPreAlertFired = false;
            UpdateTimerDisplay();
        }

        private void BtnResetTimer_Click(object sender, RoutedEventArgs e)
        {
            _countdownTimer.Stop();
            _timerRemaining = _timerInitial;
            _timerPreAlertFired = false;
            UpdateTimerDisplay();
        }

        private void BtnAddReminder_Click(object sender, RoutedEventArgs e)
        {
            _reminders.Add(new ReminderItem { Enabled = true, Text = string.Empty });
            UpdateReminderSummary();
            ScheduleSaveReminders();
        }

        private void BtnRemoveReminder_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is ReminderItem item)
            {
                _reminders.Remove(item);
                UpdateReminderSummary();
                ScheduleSaveReminders();
            }
        }

        private void ReminderTextChanged(object sender, TextChangedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is ReminderItem item)
            {
                item.ClearMatch();
            }
            UpdateReminderSummary();
            ScheduleSaveReminders();
        }

        private void ReminderCheckChanged(object sender, RoutedEventArgs e)
        {
            UpdateReminderSummary();
            ScheduleSaveReminders();
        }

        private void BtnClearReminders_Click(object sender, RoutedEventArgs e)
        {
            if (_reminders.Count == 0)
                return;

            if (MessageBox.Show("Clear all reminders?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                _reminders.Clear();
                UpdateReminderSummary();
                ScheduleSaveReminders();
            }
        }

        private void BtnResetEntity_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedEntityForActions))
                return;

            string name = _selectedEntityForActions;
            if (MessageBox.Show($"Reset stats for '{name}'?", "Confirm reset", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            _statsTracker.ResetEntity(name);
            _statsTracker.SaveToDatabase(_statsRepo);
            RefreshUiFromStats();
            ShowEntityDetails(name);
        }

        private void BtnDeleteEntity_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedEntityForActions))
                return;

            string name = _selectedEntityForActions;
            if (MessageBox.Show($"Delete '{name}' and all its data?", "Confirm delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            _statsTracker.DeleteEntity(name);
            _statsTracker.SaveToDatabase(_statsRepo);
            _selectedEntityForActions = null;
            _lastSelectedEntity = null;
            RefreshUiFromStats();
            ShowEntityDetails(null);
        }

        private void CmbCombat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string? selectedName = cmbCombat.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selectedName))
                return;

            CombatSession? match = _combatSessions.FirstOrDefault(c => c.Name.Equals(selectedName, StringComparison.OrdinalIgnoreCase));
            if (match != null)
            {
                _selectedCombat = match;
                RefreshCombatTotalsUi();
                RefreshCombatUi();
            }
        }

        private void ApplyPlayerOverrideFromUi()
        {
            string text = cmbPlayer.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                _playerNameOverride = null;
            }
            else
            {
                _playerNameOverride = NameNormalizer.Normalize(text);
            }

            RefreshPlayerUi();
            UpdateSocialProfileContext();
            _lastPlayerNamePreference = _playerNameOverride;
            SaveWindowSettings();
        }

        private void OnStatsUpdated()
        {
            RefreshUiFromStats();
            RefreshCombatUi();
            RefreshCombatTotalsUi();
            RefreshPlayerUi();
            PersistCurrentDedupeState();
        }

        private void RefreshUiFromStats()
        {
            Dispatcher.Invoke(() =>
            {
                var names = _statsTracker.Snapshot()
                    .Select(s => s.Name)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                string? desiredSelection = _lastSelectedEntity;

                cmbEntities.ItemsSource = names;
                UpdateOverallDiceAverage();

                if (!string.IsNullOrWhiteSpace(desiredSelection))
                {
                    string? match = names.FirstOrDefault(n => n.Equals(desiredSelection, StringComparison.OrdinalIgnoreCase));
                    if (match != null)
                    {
                        cmbEntities.SelectedItem = match;
                        ShowEntityDetails(match);
                        return;
                    }
                }

                cmbEntities.SelectedIndex = -1;
                _selectedEntityForActions = null;
                UpdateEntityButtonsState();
                ShowEntityDetails(null);
            });
        }

        private void ShowEntityDetails(string? name)
        {
            _lastSelectedEntity = name;

            if (string.IsNullOrWhiteSpace(name))
            {
                txtEntityName.Text = "Select an entity";
                txtAcStats.Text = "AC: -";
                txtTouchStats.Text = "Touch AC: -";
                txtSr.Text = "Spell Resistance: -";
                txtConcealment.Text = "Concealment: -";
                txtDiceAverage.Text = "Dice Average: -";
                txtFirstSeen.Text = "First seen: -";
                txtLastSeen.Text = "Last seen: -";
                txtSpells.Text = "-";
                return;
            }

            EntityStats? stats = _statsTracker.Get(name);
            if (stats == null)
            {
                txtEntityName.Text = name;
                txtAcStats.Text = "AC: -";
                txtTouchStats.Text = "Touch AC: -";
                txtSr.Text = "Spell Resistance: -";
                txtConcealment.Text = "Concealment: -";
                txtDiceAverage.Text = "Dice Average: -";
                txtFirstSeen.Text = "First seen: -";
                txtLastSeen.Text = "Last seen: -";
                txtSpells.Text = "-";
                return;
            }

            txtEntityName.Text = stats.Name;
            txtAcStats.Text = $"AC: min hit {FormatValue(stats.MinHitAc)} | max miss {FormatValue(stats.MaxMissAc)}";
            txtTouchStats.Text = $"Touch AC: min hit {FormatValue(stats.MinHitTouch)} | max miss {FormatValue(stats.MaxMissTouch)}";

            string srText = stats.HasSpellResistance.HasValue
                ? (stats.HasSpellResistance.Value ? "Yes" : "No")
                : "Unknown";
            txtSr.Text = $"Spell Resistance: {srText}";

            txtConcealment.Text = $"Concealment: {FormatPercent(stats.ConcealmentPercent)}";
            txtDiceAverage.Text = stats.RollCount > 0
                ? $"Dice Average: {FormatAverage(stats.RollSum, stats.RollCount)} (out of {stats.RollCount})"
                : "Dice Average: -";

            if (stats.SpellsCast.Count == 0)
            {
                txtSpells.Text = "-";
            }
            else
            {
                txtSpells.Text = string.Join(Environment.NewLine, stats.SpellsCast.OrderBy(s => s, StringComparer.OrdinalIgnoreCase));
            }

            if (stats.Abilities.Count == 0)
            {
                txtAbilities.Text = "-";
            }
            else
            {
                txtAbilities.Text = string.Join(Environment.NewLine, stats.Abilities.OrderBy(s => s, StringComparer.OrdinalIgnoreCase));
            }

            txtFirstSeen.Text = $"First seen: {FormatDate(stats.FirstSeen)}";
            txtLastSeen.Text = $"Last seen: {FormatDate(stats.LastSeen)}";
        }

        private static string FormatValue(int? value) => value.HasValue ? value.Value.ToString() : "-";

        private static string FormatPercent(int? value) => value.HasValue ? $"{value.Value}%" : "-";

        private static string FormatAverage(long sum, long count)
        {
            if (count <= 0) return "-";
            double avg = (double)sum / (double)count;
            return avg.ToString("0.0");
        }

        private static string FormatDate(DateTime? date) => date.HasValue ? date.Value.ToString("yyyy-MM-dd HH:mm:ss") : "-";

        private bool IsPlayer(string? name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return false;

            string effective = _playerNameOverride ?? _autoPlayerGuess ?? string.Empty;
            return !string.IsNullOrWhiteSpace(effective) && name.Equals(effective, StringComparison.OrdinalIgnoreCase);
        }

        private string? GetEffectivePlayerName()
        {
            return _playerNameOverride ?? _autoPlayerGuess;
        }

        private void UpdateEntityButtonsState()
        {
            bool hasSelection = !string.IsNullOrWhiteSpace(_selectedEntityForActions);
            btnResetEntity.IsEnabled = hasSelection;
            btnDeleteEntity.IsEnabled = hasSelection;
        }

        private void UpdateOverallDiceAverage()
        {
            var snapshot = _statsTracker.Snapshot();
            DiceSummary summary = DiceSummary.FromSnapshot(snapshot, _statsTracker.LastRollTimestamp);
            txtOverallDiceAverage.Text = summary.AverageDisplay;
            txtLastRollTimestamp.Text = summary.LastRollDisplay;
        }

        private void UpdateSocialProfileContext()
        {
            FlushSocialNotesIfDirty();
            string profileKey = GetEffectivePlayerName() ?? "Default";
            if (!string.Equals(profileKey, _socialStore.ActiveProfile, StringComparison.OrdinalIgnoreCase))
            {
                _socialStore.SetActiveProfile(profileKey);
                RefreshSocialUi();
            }
        }

        private void RecordIncomingRoll(string enemyName, int? rawRoll, DateTime? timestamp)
        {
            if (string.IsNullOrWhiteSpace(enemyName) || !rawRoll.HasValue)
                return;

            _statsTracker.RecordIncomingRoll(enemyName, rawRoll.Value, timestamp);
            UpdateOverallDiceAverage();
            PersistCurrentDedupeState();

            if (string.Equals(_lastSelectedEntity, enemyName, StringComparison.OrdinalIgnoreCase))
            {
                ShowEntityDetails(enemyName);
            }
        }

        private void SeedLastRollFromData()
        {
            var snapshot = _statsTracker.Snapshot();
            long totalRolls = snapshot.Sum(s => s.RollCount);
            if (totalRolls <= 0)
                return;

            if (_statsTracker.LastRollTimestamp.HasValue)
                return;

            var candidates = snapshot
                .Select(s => s.LastRollTimestamp ?? s.LastSeen)
                .Where(dt => dt.HasValue)
                .Select(dt => dt.Value)
                .ToList();

            if (candidates.Count == 0)
                return;

            DateTime fallback = candidates.Max();
            _statsTracker.EnsureLastRollTimestamp(fallback);
            PersistCurrentDedupeState();
            UpdateOverallDiceAverage();
        }

        private void ClearRollsOnce()
        {
            try
            {
                string markerPath = Path.Combine(AppContext.BaseDirectory, "RollsCleared.marker");
                if (File.Exists(markerPath))
                    return;

                _statsTracker.ClearAllRolls();
                _statsTracker.SaveToDatabase(_statsRepo);
                _ingestionStore.Save();
                File.WriteAllText(markerPath, DateTime.Now.ToString("o"));
            }
            catch
            {
                // If clearing fails, skip silently to avoid blocking startup
            }
        }

        private void RecordCombatRoll(string opponent, int roll, DateTime? timestamp, bool incoming)
        {
            if (string.IsNullOrWhiteSpace(opponent))
                return;

            DateTime ts = timestamp ?? DateTime.Now;
            EnsureCombat(ts, opponent);
            if (_currentCombat == null)
                return;

            CombatAggregate agg = GetOrCreateAggregate(_currentCombat, opponent);
            AddCombatTarget(_currentCombat, opponent, ts);
            if (incoming)
                agg.RegisterIncomingRoll(roll, ts);
            else
                agg.RegisterOutgoingRoll(roll, ts);

            if (incoming)
                _currentCombat.Overall.RegisterIncomingRoll(roll, ts);
            else
                _currentCombat.Overall.RegisterOutgoingRoll(roll, ts);

            _currentCombat.LastActivity = ts;
            RefreshCombatTotalsUi();
            RefreshCombatUi();
        }

        private void LoadSocialFile(string path, bool createIfMissing = false, bool userInitiated = false)
        {
            try
            {
                FlushSocialNotesIfDirty();
                _socialOverlayPath = path;
                _socialStore = new SocialStore(_socialOverlayPath, _socialCorePath);
                _socialStore.Load();
                if (createIfMissing && !File.Exists(_socialOverlayPath))
                {
                    _socialStore.Save();
                }
                _socialOverlayLoadedByUser = _socialOverlayLoadedByUser || userInitiated;
                UpdateSocialProfileContext();
                RefreshSocialUi();
                ShowSocialProfile(_selectedSocialName);
                UpdateSocialFileLabel();
                UpdateSocialEditingEnabled();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading social data: {ex.Message}");
            }
        }

        private void UpdateSocialFileLabel()
        {
            string name = "-";
            if (!string.IsNullOrWhiteSpace(_socialOverlayPath))
            {
                name = Path.GetFileNameWithoutExtension(_socialOverlayPath);
            }
            txtSocialFileName.Text = name;
        }

        private void UpdateSocialEditingEnabled()
        {
            bool enabled = _socialOverlayLoadedByUser;
            chkHaveMet.IsEnabled = enabled;
            chkKnowName.IsEnabled = enabled;
            txtAlias.IsEnabled = enabled;
            txtSocialNotes.IsEnabled = enabled;
            btnSaveSocial.IsEnabled = enabled;
        }

        private void OnAnyEvent(ParsedEvent parsed)
        {
            if (!string.IsNullOrWhiteSpace(parsed.Actor))
            {
                string? guess = _playerDetector.Register(parsed);
                if (!string.Equals(guess, _autoPlayerGuess, StringComparison.OrdinalIgnoreCase))
                {
                    _autoPlayerGuess = guess;
                    RefreshPlayerUi();
                    UpdateSocialProfileContext();
                }
            }

            bool added = false;
            if (parsed.EventType == "talk")
            {
                DateTime? ts = parsed.Timestamp ?? DateTime.Now;
                if (!string.IsNullOrWhiteSpace(parsed.Actor))
                {
                    added |= _socialStore.Touch(parsed.Actor, ts, parsed.ControllerUser);
                    RegisterRecentTalker(parsed.Actor);
                }
                if (!string.IsNullOrWhiteSpace(parsed.Target))
                {
                    added |= _socialStore.Touch(parsed.Target, ts, parsed.ControllerUser);
                    RegisterRecentTalker(parsed.Target);
                }
            }

            if (added)
            {
                _socialStore.Save();
                RefreshSocialUi();
            }
            else if (!string.IsNullOrWhiteSpace(_selectedSocialName))
            {
                if (string.Equals(_selectedSocialName, parsed.Actor, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(_selectedSocialName, parsed.Target, StringComparison.OrdinalIgnoreCase))
                {
                    ShowSocialProfile(_selectedSocialName);
                }
            }
        }

        private void CountdownTimer_Tick(object? sender, EventArgs e)
        {
            if (_timerRemaining.TotalSeconds <= 0)
            {
                _countdownTimer.Stop();
                _timerRemaining = TimeSpan.Zero;
                UpdateTimerDisplay();
                return;
            }

            _timerRemaining = _timerRemaining - TimeSpan.FromSeconds(1);
            if (_timerPreAlert > TimeSpan.Zero && !_timerPreAlertFired && _timerRemaining <= _timerPreAlert && _timerRemaining > TimeSpan.Zero)
            {
                _timerPreAlertFired = true;
                PlayBeep();
            }

            if (_timerRemaining.TotalSeconds <= 0)
            {
                _timerRemaining = TimeSpan.Zero;
                _countdownTimer.Stop();
                PlayBeep();
            }

            UpdateTimerDisplay();
        }

        private int GetSelectedTimerMinutes()
        {
            if (cmbTimerDuration.SelectedItem is ComboBoxItem item && item.Tag is string tag && int.TryParse(tag, out int m))
            {
                return m;
            }

            if (cmbTimerDuration.SelectedItem is int mi)
                return mi;

            if (cmbTimerDuration.SelectedItem is string s && int.TryParse(s, out int ms))
                return ms;

            return 0;
        }

        private void InitializeTimerOptions()
        {
            cmbTimerDuration.Items.Clear();
            for (int i = 1; i <= 30; i++)
            {
                ComboBoxItem item = new ComboBoxItem { Content = $"{i} min", Tag = i.ToString() };
                cmbTimerDuration.Items.Add(item);
            }
            cmbTimerDuration.SelectedIndex = 4; // default 5 min
            UpdateTimerDisplay();

            cmbTimerPreAlert.Items.Clear();
            cmbTimerPreAlert.Items.Add(new ComboBoxItem { Content = "None", Tag = "0" });
            cmbTimerPreAlert.Items.Add(new ComboBoxItem { Content = "30 sec", Tag = "0.5" });
            cmbTimerPreAlert.Items.Add(new ComboBoxItem { Content = "1 min", Tag = "1" });
            cmbTimerPreAlert.Items.Add(new ComboBoxItem { Content = "2 min", Tag = "2" });
            cmbTimerPreAlert.Items.Add(new ComboBoxItem { Content = "5 min", Tag = "5" });
            cmbTimerPreAlert.SelectedIndex = 0;
        }

        private void UpdateTimerDisplay()
        {
            if (_timerRemaining.TotalSeconds <= 0)
            {
                txtTimerDisplay.Text = "--:--";
            }
            else
            {
                txtTimerDisplay.Text = $"{(int)_timerRemaining.TotalMinutes:00}:{_timerRemaining.Seconds:00}";
            }
        }

        private TimeSpan GetSelectedPreAlert()
        {
            if (cmbTimerPreAlert.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            {
                if (double.TryParse(tag, out double minutes))
                {
                    return TimeSpan.FromMinutes(minutes);
                }
            }
            return TimeSpan.Zero;
        }

        private void RecentTalkButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Content is string name && !string.IsNullOrWhiteSpace(name))
            {
                FlushSocialNotesIfDirty();
                _selectedSocialName = name;
                RefreshSocialUi();
                ShowSocialProfile(name);
            }
        }

        private void RegisterRecentTalker(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            if (IsPlayer(name))
                return;

            string? existing = _recentTalkers.FirstOrDefault(n => n.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(existing))
            {
                _recentTalkers.Remove(existing);
            }

            _recentTalkers.Insert(0, name);
            while (_recentTalkers.Count > 5)
            {
                _recentTalkers.RemoveAt(_recentTalkers.Count - 1);
            }
        }

        private void SliderVolume_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            _beepVolume = sliderVolume.Value;
            _beepPlayer.Volume = _beepVolume;
            txtVolumeValue.Text = $"{(int)(_beepVolume * 100)}%";
        }

        private string EnsureBeepSoundFile()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "beep.wav");
            if (File.Exists(path))
                return path;

            const int sampleRate = 44100;
            const double durationSeconds = 0.2;
            const double frequency = 880.0;
            int samples = (int)(sampleRate * durationSeconds);
            short[] data = new short[samples];
            for (int i = 0; i < samples; i++)
            {
                double t = i / (double)sampleRate;
                data[i] = (short)(Math.Sin(2 * Math.PI * frequency * t) * short.MaxValue * 0.5);
            }

            int byteRate = sampleRate * 2;
            int dataChunkSize = samples * 2;
            int fmtChunkSize = 16;
            int fileSize = 4 + (8 + fmtChunkSize) + (8 + dataChunkSize);

            using (BinaryWriter writer = new BinaryWriter(File.Open(path, FileMode.Create, FileAccess.Write)))
            {
                writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));
                writer.Write(fileSize);
                writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVE"));
                writer.Write(System.Text.Encoding.ASCII.GetBytes("fmt "));
                writer.Write(fmtChunkSize);
                writer.Write((short)1); // PCM
                writer.Write((short)1); // channels
                writer.Write(sampleRate);
                writer.Write(byteRate);
                writer.Write((short)2); // block align
                writer.Write((short)16); // bits per sample
                writer.Write(System.Text.Encoding.ASCII.GetBytes("data"));
                writer.Write(dataChunkSize);
                foreach (short s in data)
                {
                    writer.Write(s);
                }
            }

            return path;
        }

        private void PlayBeep()
        {
            if (_beepVolume <= 0)
                return;
            try
            {
                _beepPlayer.Stop();
                _beepPlayer.Position = TimeSpan.Zero;
                _beepPlayer.Volume = _beepVolume;
                _beepPlayer.Play();
            }
            catch
            {
                // ignore audio failures
            }
        }

        private void SaveWindowSettings()
        {
            try
            {
                WindowSettings settings = new WindowSettings
                {
                    Left = this.Left,
                    Top = this.Top,
                    Width = this.Width,
                    Height = this.Height,
                    State = this.WindowState,
                    LastPlayerName = _lastPlayerNamePreference
                };
                var options = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(_windowSettingsPath, JsonSerializer.Serialize(settings, options));
            }
            catch
            {
                // ignore save errors
            }
        }

        private void LoadWindowSettings()
        {
            try
            {
                if (!File.Exists(_windowSettingsPath))
                    return;

                string json = File.ReadAllText(_windowSettingsPath);
                WindowSettings? settings = JsonSerializer.Deserialize<WindowSettings>(json);
                if (settings == null)
                    return;

                this.Width = settings.Width > 0 ? settings.Width : this.Width;
                this.Height = settings.Height > 0 ? settings.Height : this.Height;
                if (IsFinite(settings.Left) && IsFinite(settings.Top))
                {
                    this.WindowStartupLocation = WindowStartupLocation.Manual;
                    this.Left = settings.Left;
                    this.Top = settings.Top;
                }
                if (Enum.IsDefined(typeof(WindowState), settings.State))
                {
                    this.WindowState = settings.State;
                }

                if (!string.IsNullOrWhiteSpace(settings.LastPlayerName))
                {
                    _lastPlayerNamePreference = settings.LastPlayerName.Trim();
                    _playerDetector.SeedBias(_lastPlayerNamePreference);
                    _autoPlayerGuess = _lastPlayerNamePreference;
                }
            }
            catch
            {
                // ignore load errors
            }
        }

        private static bool IsFinite(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private void PersistCurrentDedupeState()
        {
            if (string.IsNullOrWhiteSpace(_currentLogPath))
                return;

            _rollDedupeStore.Set(_currentLogPath, _statsTracker.SnapshotDedupeState());
            _rollDedupeStore.Save();
        }

        private void LoadReminders()
        {
            try
            {
                if (File.Exists(_reminderFilePath))
                {
                    string json = File.ReadAllText(_reminderFilePath);
                    List<ReminderItem> list = new List<ReminderItem>();
                    try
                    {
                        using var doc = JsonDocument.Parse(json);
                        if (doc.RootElement.ValueKind == JsonValueKind.Array)
                        {
                            list = JsonSerializer.Deserialize<List<ReminderItem>>(json) ?? new List<ReminderItem>();
                        }
                        else if (doc.RootElement.ValueKind == JsonValueKind.Object && doc.RootElement.TryGetProperty("reminders", out var remindersEl) && remindersEl.ValueKind == JsonValueKind.Array)
                        {
                            list = JsonSerializer.Deserialize<List<ReminderItem>>(remindersEl.GetRawText()) ?? new List<ReminderItem>();
                        }
                    }
                    catch
                    {
                        list = JsonSerializer.Deserialize<List<ReminderItem>>(json) ?? new List<ReminderItem>();
                    }

                    _reminders.Clear();
                    foreach (var item in list)
                    {
                        var cleaned = NormalizeReminder(item);
                        if (cleaned != null)
                        {
                            _reminders.Add(cleaned);
                        }
                    }
                }
            }
            catch
            {
                // ignore load errors
            }

            if (_reminders.Count == 0)
            {
                _reminders.Add(new ReminderItem { Enabled = true, Text = "Example reminder" });
            }

            UpdateReminderSummary();
        }

        private ReminderItem? NormalizeReminder(ReminderItem? item)
        {
            if (item == null)
                return null;

            string text = item.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
                return null;

            return new ReminderItem
            {
                Enabled = item.Enabled,
                Text = text,
                CaseSensitive = item.CaseSensitive,
                WholeWord = item.WholeWord
            };
        }

        private void SaveReminders()
        {
            try
            {
                List<ReminderItem> snapshot;
                lock (_reminders)
                {
                    snapshot = _reminders
                        .Select(NormalizeReminder)
                        .Where(r => r != null)
                        .Cast<ReminderItem>()
                        .ToList();
                }

                var dto = new { reminders = snapshot };
                var options = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(_reminderFilePath, JsonSerializer.Serialize(dto, options));
            }
            catch
            {
                // ignore save errors
            }
        }

        private void ScheduleSaveReminders()
        {
            _reminderSaveDebounce.Stop();
            _reminderSaveDebounce.Start();
        }

        private void ReminderSaveDebounce_Tick(object? sender, EventArgs e)
        {
            _reminderSaveDebounce.Stop();
            SaveReminders();
        }

        private void UpdateReminderSummary()
        {
            int active = _reminders.Count(r => r.Enabled && !string.IsNullOrWhiteSpace(r.Text));
            int total = _reminders.Count(r => !string.IsNullOrWhiteSpace(r.Text));
            txtReminderSummary.Text = $"Active: {active} / {total}";
        }

        private void OnRawLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
                return;

            List<ReminderItem> snapshot;
            lock (_reminders)
            {
                snapshot = _reminders.ToList();
            }

            foreach (var reminder in snapshot)
            {
                if (!reminder.Enabled || string.IsNullOrWhiteSpace(reminder.Text))
                    continue;

                if (ReminderMatches(reminder, line, out int matchStart, out int matchLength))
                {
                    Dispatcher.Invoke(() => reminder.UpdateMatch(line, matchStart, matchLength));
                    PlayBeep();
                    break;
                }
            }
        }

        private bool ReminderMatches(ReminderItem reminder, string line, out int matchStart, out int matchLength)
        {
            if (string.IsNullOrWhiteSpace(reminder.Text))
            {
                matchStart = -1;
                matchLength = 0;
                return false;
            }

            string needle = reminder.Text;
            if (reminder.WholeWord)
            {
                string pattern = $@"\b{Regex.Escape(needle)}\b";
                var options = reminder.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                var match = Regex.Match(line, pattern, options);
                if (match.Success)
                {
                    matchStart = match.Index;
                    matchLength = match.Length;
                    return true;
                }
            }
            else
            {
                var comparison = reminder.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                int idx = line.IndexOf(needle, comparison);
                if (idx >= 0)
                {
                    matchStart = idx;
                    matchLength = needle.Length;
                    return true;
                }
            }

            matchStart = -1;
            matchLength = 0;
            return false;
        }

        private void RefreshPlayerUi()
        {
            Dispatcher.Invoke(() =>
            {
                txtPlayerAuto.Text = $"Player - Auto: {(_autoPlayerGuess ?? "-")}";
                cmbPlayer.ItemsSource = _statsTracker.Snapshot()
                    .Select(s => s.Name)
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (string.IsNullOrWhiteSpace(_playerNameOverride))
                {
                    cmbPlayer.Text = _autoPlayerGuess ?? cmbPlayer.Text;
                }
            });
        }

        private void RefreshSocialUi()
        {
            Dispatcher.Invoke(() =>
            {
                var names = _socialStore.Names();
                if (!string.IsNullOrWhiteSpace(_socialSearchTerm))
                {
                    names = names.Where(n =>
                    {
                        if (n.IndexOf(_socialSearchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;

                        SocialProfile? p = _socialStore.Get(n);
                        if (p != null && !string.IsNullOrWhiteSpace(p.Alias) && p.Alias.IndexOf(_socialSearchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;

                        return false;
                    }).ToList();
                }
                cmbSocial.ItemsSource = names;

                if (!string.IsNullOrWhiteSpace(_selectedSocialName) && names.Contains(_selectedSocialName))
                {
                    cmbSocial.SelectedItem = _selectedSocialName;
                }
                else if (names.Count > 0)
                {
                    _selectedSocialName = names[0];
                    cmbSocial.SelectedItem = _selectedSocialName;
                }
                else
                {
                    cmbSocial.SelectedIndex = -1;
                    _selectedSocialName = null;
                    ClearSocialFields();
                }
            });
        }

        private void ShowSocialProfile(string? name)
        {
            _selectedSocialName = name;
            SocialProfile? profile = string.IsNullOrWhiteSpace(name) ? null : _socialStore.Get(name);
            _loadingSocial = true;
            if (profile == null)
            {
                ClearSocialFields();
                _loadingSocial = false;
                return;
            }

            chkHaveMet.IsChecked = profile.HaveMet;
            chkKnowName.IsChecked = profile.KnowName;
            txtAlias.Text = profile.Alias ?? string.Empty;
            txtSocialFirstSeen.Text = profile.FirstSeen.HasValue ? profile.FirstSeen.Value.ToString("yyyy-MM-dd hh:mm tt") : "-";
            txtSocialLastSeen.Text = profile.LastSeen.HasValue ? profile.LastSeen.Value.ToString("yyyy-MM-dd hh:mm tt") : "-";
            txtSocialNotes.Text = profile.Notes ?? string.Empty;
            txtSocialUsers.Text = profile.Users.Count > 0 ? string.Join(", ", profile.Users) : "-";
            var related = _socialStore.RelatedByUsers(profile);
            txtSocialRelated.Text = related.Count > 0 ? string.Join(", ", related) : "-";
            EntityStats? stats = _statsTracker.Get(profile.Name);
            if (stats != null && stats.SpellsCast.Count > 0)
            {
                spellsItems.ItemsSource = stats.SpellsCast.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
            }
            else
            {
                spellsItems.ItemsSource = new List<string>();
            }
            _loadingSocial = false;
        }

        private void ClearSocialFields()
        {
            chkHaveMet.IsChecked = false;
            chkKnowName.IsChecked = false;
            txtAlias.Text = string.Empty;
            txtSocialFirstSeen.Text = "-";
            txtSocialLastSeen.Text = "-";
            txtSocialNotes.Text = string.Empty;
            txtSocialUsers.Text = "-";
            txtSocialRelated.Text = "-";
            spellsItems.ItemsSource = new List<string>();
        }

        private void SaveSocialProfile(bool includeNotes)
        {
            if (string.IsNullOrWhiteSpace(_selectedSocialName))
                return;

            string? alias = string.IsNullOrWhiteSpace(txtAlias.Text) ? null : txtAlias.Text.Trim();
            _socialStore.SetAlias(_selectedSocialName, alias);
            SocialOverlay overlay = _socialStore.GetOrCreateOverlay(_selectedSocialName);
            overlay.HaveMet = chkHaveMet.IsChecked == true;
            overlay.KnowName = chkKnowName.IsChecked == true;
            if (includeNotes)
            {
                overlay.Notes = txtSocialNotes.Text;
                _socialNotesDirty = false;
                _socialNotesDebounce.Stop();
            }
            _socialStore.Save();
            ShowSocialProfile(_selectedSocialName);
        }

        private void RefreshCombatSelection()
        {
            Dispatcher.Invoke(() =>
            {
                cmbCombat.ItemsSource = _combatSessions.Select(s => s.Name).ToList();
                if (_selectedCombat != null)
                {
                    cmbCombat.SelectedItem = _selectedCombat.Name;
                }
                else if (_currentCombat != null)
                {
                    cmbCombat.SelectedItem = _currentCombat.Name;
                    _selectedCombat = _currentCombat;
                }
                else
                {
                    cmbCombat.SelectedIndex = -1;
                }
            });
        }

        private void OnAttackSeen(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Target))
                return;

            bool actorIsPlayer = IsPlayer(parsed.Actor);
            bool targetIsPlayer = IsPlayer(parsed.Target);

            if (!actorIsPlayer && !targetIsPlayer)
                return;

            if (targetIsPlayer && parsed.RawRoll.HasValue && !string.IsNullOrWhiteSpace(parsed.Actor))
            {
                RecordIncomingRoll(parsed.Actor!, parsed.RawRoll, parsed.Timestamp);
                RecordCombatRoll(parsed.Actor!, parsed.RawRoll.Value, parsed.Timestamp, incoming: true);
            }

            if (actorIsPlayer && parsed.RawRoll.HasValue)
            {
                RecordCombatRoll(parsed.Target, parsed.RawRoll.Value, parsed.Timestamp, incoming: false);
                RecordCombatActivity(parsed.Target, parsed.Timestamp, session =>
                {
                    AddCombatTarget(session, parsed.Target, parsed.Timestamp);
                });
            }
        }

        private void OnDamageSeen(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Target) || parsed.Amount == null)
                return;

            bool actorIsPlayer = IsPlayer(parsed.Actor);
            bool targetIsPlayer = IsPlayer(parsed.Target);
            if (!actorIsPlayer && string.Equals(parsed.Actor, "Someone", StringComparison.OrdinalIgnoreCase))
                return; // ignore environment-only events

            if (!actorIsPlayer && !targetIsPlayer)
                return; // ignore non-player combat

            string combatKey = actorIsPlayer ? parsed.Target! : parsed.Actor ?? parsed.Target!;

            RecordCombatActivity(combatKey, parsed.Timestamp, session =>
            {
                if (!targetIsPlayer)
                {
                    AddCombatTarget(session, parsed.Target, parsed.Timestamp);
                    AddDamage(session, parsed.Target, parsed.Amount.Value, parsed.Timestamp);
                }
                else
                {
                    if (!string.IsNullOrWhiteSpace(parsed.Actor))
                    {
                        AddCombatTarget(session, parsed.Actor, parsed.Timestamp);
                    }
                    AddDamageTaken(session, parsed.Actor, parsed.Amount.Value, parsed.Timestamp);
                }
            });

            if (targetIsPlayer && parsed.RawRoll.HasValue && !string.IsNullOrWhiteSpace(parsed.Actor))
            {
                RecordIncomingRoll(parsed.Actor!, parsed.RawRoll, parsed.Timestamp);
                RecordCombatRoll(parsed.Actor!, parsed.RawRoll.Value, parsed.Timestamp, incoming: true);
            }
        }

        private void OnHealSeen(ParsedEvent parsed)
        {
            if (parsed.Amount == null)
                return;

            if (!IsPlayer(parsed.Actor))
                return;

            RecordCombatActivity(parsed.Actor, parsed.Timestamp, session =>
            {
                AddHealing(session, parsed.Amount.Value, parsed.Timestamp);
            });
        }

        private void RecordCombatActivity(string? target, DateTime? timestamp, Action<CombatSession> apply)
        {
            DateTime now = timestamp ?? DateTime.Now;
            EnsureCombat(now, target);
            if (_currentCombat == null)
                return;

            apply(_currentCombat);
            _currentCombat.LastActivity = now;

            if (_selectedCombat == null)
            {
                _selectedCombat = _currentCombat;
                RefreshCombatSelection();
            }

            RefreshCombatTotalsUi();
            RefreshCombatUi();
        }

        private void EnsureCombat(DateTime now, string? firstTarget)
        {
            if (_currentCombat != null && (now - _currentCombat.LastActivity).TotalSeconds <= 40)
                return;

            _combatCounter++;
            string namePart = !string.IsNullOrWhiteSpace(firstTarget) ? firstTarget : "Combat";
            string combatName = $"{namePart} ({_combatCounter})";

            CombatSession session = new CombatSession
            {
                Name = combatName,
                LastActivity = now,
                Start = now
            };

            _combatSessions.Add(session);
            CombatSession? previous = _currentCombat;
            _currentCombat = session;
            if (_selectedCombat == null || _selectedCombat == previous)
            {
                _selectedCombat = session;
            }
            RefreshCombatSelection();
        }

        private CombatAggregate GetOrCreateAggregate(CombatSession session, string name)
        {
            if (!session.Aggregates.TryGetValue(name, out CombatAggregate? agg))
            {
                agg = new CombatAggregate { Name = name };
                session.Aggregates[name] = agg;
            }
            return agg;
        }

        private void AddDamage(CombatSession session, string target, int amount, DateTime? timestamp)
        {
            DateTime ts = timestamp ?? DateTime.Now;
            CombatAggregate targetAgg = GetOrCreateAggregate(session, target);
            targetAgg.RegisterDamage(amount, ts);

            session.Overall.RegisterDamage(amount, ts);
        }

        private void AddDamageTaken(CombatSession session, string? source, int amount, DateTime? timestamp)
        {
            DateTime ts = timestamp ?? DateTime.Now;
            session.Overall.RegisterDamageTaken(amount, ts);
            if (!string.IsNullOrWhiteSpace(source))
            {
                CombatAggregate agg = GetOrCreateAggregate(session, source);
                agg.RegisterDamageTaken(amount, ts);
            }
        }

        private void AddHealing(CombatSession session, int amount, DateTime? timestamp)
        {
            DateTime ts = timestamp ?? DateTime.Now;
            session.Overall.RegisterHeal(amount, ts);
        }

        private void AddCombatTarget(CombatSession session, string target, DateTime? timestamp)
        {
            DateTime seen = timestamp ?? DateTime.Now;
            if (IsPlayer(target))
                return;
            int existingIndex = session.Slots.FindIndex(s => s.Name.Equals(target, StringComparison.OrdinalIgnoreCase));
            if (existingIndex >= 0)
            {
                session.Slots[existingIndex].LastSeen = seen;
                return;
            }

            if (session.Slots.Count < 5)
            {
                session.Slots.Add(new CombatSlot { Name = target, LastSeen = seen });
            }
            else
            {
                int oldestIndex = 0;
                DateTime oldestTime = session.Slots[0].LastSeen;
                for (int i = 1; i < session.Slots.Count; i++)
                {
                    if (session.Slots[i].LastSeen < oldestTime)
                    {
                        oldestTime = session.Slots[i].LastSeen;
                        oldestIndex = i;
                    }
                }
                session.Slots[oldestIndex] = new CombatSlot { Name = target, LastSeen = seen };
            }
        }

        private void RefreshCombatUi()
        {
            Dispatcher.Invoke(() =>
            {
                CombatSession? view = _selectedCombat ?? _currentCombat;
                List<CombatDisplay> display = new List<CombatDisplay>();
                if (view != null)
                {
                    foreach (CombatSlot slot in view.Slots)
                    {
                        EntityStats? stats = _statsTracker.Get(slot.Name);
                        CombatAggregate? agg = view.Aggregates.ContainsKey(slot.Name) ? view.Aggregates[slot.Name] : null;
                        display.Add(new CombatDisplay
                        {
                            Name = slot.Name,
                            DamageText = agg != null ? $"Damage: {agg.Damage:F0}" : "Damage: -",
                            DpsText = agg != null ? $"DPS: {agg.ComputeDps(DateTime.Now):F1}" : "DPS: -",
                            DamageTakenText = FormatDamageTaken(agg, out Brush brush),
                            DamageTakenBrush = brush,
                            Summary = BuildCombatSummary(stats, agg),
                            LastSeen = slot.LastSeen
                        });
                    }
                }

                while (display.Count < 5)
                {
                    display.Add(new CombatDisplay
                    {
                        Name = "(empty)",
                        Summary = "-",
                        LastSeen = DateTime.MinValue
                    });
                }

                combatItems.ItemsSource = display;
            });
        }

        private string BuildCombatSummary(EntityStats? stats, CombatAggregate? agg)
        {
            if (stats == null)
                return "No data";

            string ac = $"AC min hit {FormatValue(stats.MinHitAc)}, max miss {FormatValue(stats.MaxMissAc)}";
            string touch = $"Touch min hit {FormatValue(stats.MinHitTouch)}, max miss {FormatValue(stats.MaxMissTouch)}";
            string sr = stats.HasSpellResistance.HasValue ? (stats.HasSpellResistance.Value ? "SR: Yes" : "SR: No") : "SR: Unknown";
            string conceal = $"Conceal: {FormatPercent(stats.ConcealmentPercent)}";
            string incomingRoll = agg != null && agg.IncomingRollCount > 0 ? $"Rolls vs you: {FormatAverage(agg.IncomingRollSum, agg.IncomingRollCount)}" : "Rolls vs you: -";
            string outgoingRoll = agg != null && agg.OutgoingRollCount > 0 ? $"Your rolls: {FormatAverage(agg.OutgoingRollSum, agg.OutgoingRollCount)}" : "Your rolls: -";
            return $"{ac}\n{touch}\n{sr}\n{conceal}\n{incomingRoll}\n{outgoingRoll}";
        }

        private void RefreshCombatTotalsUi()
        {
            Dispatcher.Invoke(() =>
            {
                CombatSession? view = _selectedCombat ?? _currentCombat;
                if (view == null)
                {
                    txtTotalDamage.Text = "0";
                    txtTotalHealing.Text = "0";
                    txtTotalDps.Text = "0";
                    txtCombatStart.Text = "-";
                    txtCombatDuration.Text = "-";
                    txtTotalDamageTaken.Text = "-";
                }
                else
                {
                    txtTotalDamage.Text = view.Overall.Damage.ToString("F0");
                    txtTotalHealing.Text = view.Overall.Healing.ToString("F0");
                    txtTotalDps.Text = view.Overall.ComputeDps(DateTime.Now).ToString("F1");
                    txtTotalDamageTaken.Text = FormatDamageTaken(view.Overall, out _);
                    txtCombatStart.Text = view.Start.ToString("hh:mm:ss tt");
                    double durationSeconds = Math.Max(0, (view.LastActivity - view.Start).TotalSeconds);
                    txtCombatDuration.Text = $"{durationSeconds:F1}s";
                }
            });
        }

        private string FormatDamageTaken(CombatAggregate? agg, out Brush brush)
        {
            brush = Brushes.White;
            if (agg == null)
                return string.Empty;

            if (agg.DamageTaken <= 0)
            {
                brush = Brushes.Transparent;
                return string.Empty;
            }

            brush = Brushes.Red;
            return $"Damage Taken: {agg.DamageTaken:F0}";
        }
    }

    /// <summary>
    /// Handles file monitoring and polling functionality
    /// </summary>
    public class FileMonitor
    {
        private string? _currentFilePath;
        private DispatcherTimer? _pollingTimer;
        private Action<string>? _contentChangedCallback;
        private FileState? _currentState;
        private readonly Dictionary<string, FileState> _fileStates = new(StringComparer.OrdinalIgnoreCase);
        private IngestionStateStore? _ingestionStore;
        private const int MaxRecentLineHashes = 500;



        public void StartMonitoring(string filePath, Action<string> contentChangedCallback, IngestionStateStore ingestionStore)
        {
            // Stop any existing timer
            StopMonitoring();

            if (string.IsNullOrEmpty(filePath))
                return;

            _currentFilePath = filePath;
            _contentChangedCallback = contentChangedCallback;
            _ingestionStore = ingestionStore;

            if (!_fileStates.TryGetValue(filePath, out var state))
            {
                state = new FileState();
                _fileStates[filePath] = state;
            }

            // Restore persisted state if available
            var persisted = _ingestionStore?.Get(filePath);
            if (persisted != null)
            {
                state.LastPosition = persisted.LastOffset;
                state.RecentLineHashes = new Queue<string>(persisted.RecentLineHashes ?? Enumerable.Empty<string>());
                state.RecentHashSet = new HashSet<string>(state.RecentLineHashes, StringComparer.OrdinalIgnoreCase);
            }
            _currentState = state;

            try
            {
                // Set up polling timer to check for changes every 500ms (adjust as needed)
                _pollingTimer = new DispatcherTimer();
                _pollingTimer.Interval = TimeSpan.FromMilliseconds(500);
                _pollingTimer.Tick += PollingTimer_Tick;
                _pollingTimer.Start();

                // Load any initial content
                ReadNewLines();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error starting monitoring: {ex.Message}");
            }
        }

        public void StopMonitoring()
        {
            _pollingTimer?.Stop();
            _pollingTimer = null;
            _currentFilePath = null;
            _contentChangedCallback = null;
            _currentState = null;
            _ingestionStore = null;
        }

        private void PollingTimer_Tick(object? sender, EventArgs e)
        {
            ReadNewLines();
        }

        private void ReadNewLines()
        {
            if (string.IsNullOrEmpty(_currentFilePath) || _contentChangedCallback == null)
                return;
            if (_currentState == null)
                return;

            try
            {
                using FileStream stream = new FileStream(_currentFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                if (stream.Length < _currentState.LastPosition)
                {
                    _currentState.LastPosition = 0;
                }

                stream.Seek(_currentState.LastPosition, SeekOrigin.Begin);
                using StreamReader reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
                string newText = reader.ReadToEnd();
                _currentState.LastPosition = stream.Position;

                if (string.IsNullOrEmpty(newText) && string.IsNullOrEmpty(_currentState.LineBuffer))
                    return;

                _currentState.LineBuffer += newText;
                string[] lines = _currentState.LineBuffer.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                if (lines.Length > 1)
                {
                    List<string> completeLines = new List<string>();

                    for (int i = 0; i < lines.Length - 1; i++)
                    {
                        string line = lines[i];
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        string hash = LogStatsTracker.ComputeLineHashStatic(line);
                        if (_currentState.RecentHashSet.Contains(hash))
                            continue;

                        _currentState.RecentLineHashes.Enqueue(hash);
                        _currentState.RecentHashSet.Add(hash);
                        if (_currentState.RecentLineHashes.Count > MaxRecentLineHashes)
                        {
                            string old = _currentState.RecentLineHashes.Dequeue();
                            _currentState.RecentHashSet.Remove(old);
                        }

                        completeLines.Add(line);
                    }

                    if (completeLines.Count > 0)
                    {
                        string combined = string.Join(Environment.NewLine, completeLines);
                        _contentChangedCallback?.Invoke(combined);
                    }
                }

                _currentState.LineBuffer = lines[^1];

                // Persist ingestion state
                if (_ingestionStore != null && _currentFilePath != null)
                {
                    _ingestionStore.Set(_currentFilePath, new IngestionState
                    {
                        FilePath = _currentFilePath,
                        LastOffset = _currentState.LastPosition,
                        FileLength = stream.Length,
                        LastWriteTimeUtc = File.GetLastWriteTimeUtc(_currentFilePath),
                        RecentLineHashes = _currentState.RecentLineHashes.ToList()
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading file: {ex.Message}");
            }
        }

        private class FileState
        {
            public long LastPosition { get; set; } = 0;
            public string LineBuffer { get; set; } = string.Empty;
            public Queue<string> RecentLineHashes { get; set; } = new Queue<string>();
            public HashSet<string> RecentHashSet { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// Interface for content processing
    /// </summary>
    public interface IContentProcessor
    {
        void ProcessContent(string content);
    }

    /// <summary>
    /// Default content processor that simply displays text in the UI
    /// </summary>
    public class DefaultContentProcessor : IContentProcessor
    {
        private TextBox _outputTextBox;



        public DefaultContentProcessor(TextBox outputTextBox)
        {
            _outputTextBox = outputTextBox;
        }

        public void ProcessContent(string content)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    _outputTextBox.Text = content;
                    _outputTextBox.ScrollToEnd();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error processing content: {ex.Message}");
                }
            });
        }
    }

    public class LogProcessingContentProcessor : IContentProcessor
    {
        private readonly LogParser _parser;
        private readonly LogStatsTracker _statsTracker;
        private readonly Action? _onUpdated;
        private readonly Action<ParsedEvent>? _onAttack;
        private readonly Action<ParsedEvent>? _onDamage;
        private readonly Action<ParsedEvent>? _onHeal;
        private readonly Action<ParsedEvent>? _onAnyEvent;
        private readonly Action<string>? _onRawLine;
        private readonly Action _persist;

        public LogProcessingContentProcessor(LogParser parser, LogStatsTracker statsTracker, Action persist, Action? onUpdated, Action<ParsedEvent>? onAttack, Action<ParsedEvent>? onDamage, Action<ParsedEvent>? onHeal, Action<ParsedEvent>? onAnyEvent, Action<string>? onRawLine)
        {
            _parser = parser;
            _statsTracker = statsTracker;
            _persist = persist;
            _onUpdated = onUpdated;
            _onAttack = onAttack;
            _onDamage = onDamage;
            _onHeal = onHeal;
            _onAnyEvent = onAnyEvent;
            _onRawLine = onRawLine;
        }

        public void ProcessContent(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
                return;

            string[] lines = content.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string line in lines)
            {
                _onRawLine?.Invoke(line);
                foreach (ParsedEvent parsed in _parser.Parse(line))
                {
                    _onAnyEvent?.Invoke(parsed);
                    _statsTracker.Apply(parsed);
                    if (parsed.EventType == "attack" || parsed.EventType == "touch_attack")
                    {
                        _onAttack?.Invoke(parsed);
                    }
                    else if (parsed.EventType == "damage")
                    {
                        _onDamage?.Invoke(parsed);
                    }
                    else if (parsed.EventType == "heal")
                    {
                        _onHeal?.Invoke(parsed);
                    }
                }
            }

            WriteCsvSnapshot();
            _onUpdated?.Invoke();
        }

        private void WriteCsvSnapshot()
        {
            _persist();
        }
    }

        public class ParsedEvent
        {
            public string EventType { get; set; } = string.Empty;
            public string? Actor { get; set; }
            public string? Target { get; set; }
        // Stores the total attack result (d20 + modifiers) for AC calculations
        public int? RawRoll { get; set; }
        public bool? Hit { get; set; }
        public bool IsTouchAttack { get; set; }
        public bool HasSpellResistance { get; set; }
        public int? ConcealmentPercent { get; set; }
            public string? SpellName { get; set; }
            public string? TimestampText { get; set; }
            public DateTime? Timestamp { get; set; }
            public int? Amount { get; set; }
            public string? ControllerUser { get; set; }
            public string? Ability { get; set; }
            public string Raw { get; set; } = string.Empty;
        }

    public class LogParser
    {
        private static readonly Regex ChatPrefixRegex = new(@"^\[CHAT WINDOW TEXT\]\s*\[(?<ts>.+?)\]\s*(?<rest>.*)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex AttackRegex = new(@"(?<actor>.+?) attempts (?<attackType>.+?) on (?<target>.+?) : \*(?<outcome>[^*]+)\* : \((?<roll>\d+)\s+\+\s+(?<bonus>-?\d+)\s+=\s+(?<total>\d+)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex AttackSimpleRegex = new(@"(?<actor>.+?) attacks (?<target>.+?) : \*(?<outcome>[^*]+)\* : \((?<roll>\d+)\s+\+\s+(?<bonus>-?\d+)\s+=\s+(?<total>\d+)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SpellResistRegex = new(@"(?<target>.+?) attempts to resist spell\s*:\s*(?<result>spell resisted|failure)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex SpellCastRegex = new(@"(?<actor>.+?) cast(?:s|ing)\s+(?<spell>[\p{L}\p{M}][\p{L}\p{M}' ]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ConcealmentRegex = new(@"(?:(?<target>[\p{L}\p{M}][^:]+?)[:]\s*)?(?<percent>\d{1,3})%\s*(?:miss chance|concealment|chance to miss)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex DamageRegex = new(@"(?<actor>.+?) damages (?<target>.+?): (?<amount>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex HealRegex = new(@"(?<actor>.+?) heals (?<target>.+?): (?<amount>\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex TalkRegex = new(@"^(?:\[CHAT WINDOW TEXT\]\s*\[(?<ts>.+?)\]\s*)?(?:(?<user>\[[^\]]+\])\s*)?(?<actor>[^:]+):\s*\[Talk\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public IEnumerable<ParsedEvent> Parse(string line)
        {
            List<ParsedEvent> events = new List<ParsedEvent>();
            string trimmed = line.Trim();
            string payload = trimmed;
            string? timestamp = null;
            string? ability = null;

            Match prefixMatch = ChatPrefixRegex.Match(trimmed);
            if (prefixMatch.Success)
            {
                timestamp = prefixMatch.Groups["ts"].Value;
                payload = prefixMatch.Groups["rest"].Value.Trim();
            }
            DateTime? parsedTimestamp = TryParseTimestamp(timestamp);

            if (payload.StartsWith("Power Attack", StringComparison.OrdinalIgnoreCase))
            {
                ability = "Power Attack";
                int idx = payload.IndexOf(':');
                if (idx >= 0 && idx + 1 < payload.Length)
                {
                    payload = payload.Substring(idx + 1).Trim();
                }
            }

            Match attackMatch = AttackRegex.Match(payload);
            if (attackMatch.Success)
            {
                bool isTouch = attackMatch.Groups["attackType"].Value.IndexOf("touch", StringComparison.OrdinalIgnoreCase) >= 0;
                string outcomeText = attackMatch.Groups["outcome"].Value;
                bool isHit = outcomeText.IndexOf("miss", StringComparison.OrdinalIgnoreCase) < 0;

                if (int.TryParse(attackMatch.Groups["roll"].Value, out int rawRoll))
                {
                    string actor = NameNormalizer.Normalize(attackMatch.Groups["actor"].Value);
                    string target = NameNormalizer.Normalize(attackMatch.Groups["target"].Value);
                    if (!string.IsNullOrWhiteSpace(actor) && !string.IsNullOrWhiteSpace(target))
                    {
                        events.Add(new ParsedEvent
                        {
                            EventType = isTouch ? "touch_attack" : "attack",
                            Actor = actor,
                            Target = target,
                            RawRoll = rawRoll, // d20 roll only
                            Hit = isHit,
                            IsTouchAttack = isTouch,
                            TimestampText = timestamp,
                            Timestamp = parsedTimestamp,
                            Ability = ability,
                            Raw = line
                        });
                    }
                }
            }
            else
            {
                Match attackSimple = AttackSimpleRegex.Match(payload);
                if (attackSimple.Success)
                {
                    bool isTouch = payload.IndexOf("touch", StringComparison.OrdinalIgnoreCase) >= 0;
                    string outcomeText = attackSimple.Groups["outcome"].Value;
                    bool isHit = outcomeText.IndexOf("miss", StringComparison.OrdinalIgnoreCase) < 0;

                    if (int.TryParse(attackSimple.Groups["roll"].Value, out int rawRoll))
                    {
                        string actor = NameNormalizer.Normalize(attackSimple.Groups["actor"].Value);
                        string target = NameNormalizer.Normalize(attackSimple.Groups["target"].Value);
                        if (!string.IsNullOrWhiteSpace(actor) && !string.IsNullOrWhiteSpace(target))
                        {
                            events.Add(new ParsedEvent
                            {
                                EventType = isTouch ? "touch_attack" : "attack",
                                Actor = actor,
                                Target = target,
                                RawRoll = rawRoll,
                                Hit = isHit,
                                IsTouchAttack = isTouch,
                                TimestampText = timestamp,
                                Timestamp = parsedTimestamp,
                                Ability = ability,
                                Raw = line
                            });
                        }
                    }
                }
            }

            Match srMatch = SpellResistRegex.Match(payload);
            if (srMatch.Success)
            {
                string target = NameNormalizer.Normalize(srMatch.Groups["target"].Value);
                if (!string.IsNullOrWhiteSpace(target))
                {
                    events.Add(new ParsedEvent
                    {
                        EventType = "spell_resistance",
                        Target = target,
                        HasSpellResistance = true,
                        TimestampText = timestamp,
                        Timestamp = parsedTimestamp,
                        Raw = line
                    });
                }
            }

            Match spellCastMatch = SpellCastRegex.Match(payload);
            if (spellCastMatch.Success)
            {
                string actor = NameNormalizer.Normalize(spellCastMatch.Groups["actor"].Value);
                if (!string.IsNullOrWhiteSpace(actor))
                {
                    events.Add(new ParsedEvent
                    {
                        EventType = "spell_cast",
                        Actor = actor,
                        SpellName = spellCastMatch.Groups["spell"].Value.Trim(),
                        TimestampText = timestamp,
                        Timestamp = parsedTimestamp,
                        Raw = line
                    });
                }
            }

            Match concealMatch = ConcealmentRegex.Match(payload);
            if (concealMatch.Success && int.TryParse(concealMatch.Groups["percent"].Value, out int percent))
            {
                string target = concealMatch.Groups["target"].Success ? NameNormalizer.Normalize(concealMatch.Groups["target"].Value) : string.Empty;
                events.Add(new ParsedEvent
                {
                    EventType = "concealment",
                    Target = string.IsNullOrWhiteSpace(target) ? null : target,
                    ConcealmentPercent = percent,
                    TimestampText = timestamp,
                    Timestamp = parsedTimestamp,
                    Raw = line
                });
            }

            Match dmgMatch = DamageRegex.Match(payload);
            if (dmgMatch.Success && int.TryParse(dmgMatch.Groups["amount"].Value, out int dmg))
            {
                string actor = NameNormalizer.Normalize(dmgMatch.Groups["actor"].Value);
                string target = NameNormalizer.Normalize(dmgMatch.Groups["target"].Value);
                events.Add(new ParsedEvent
                {
                    EventType = "damage",
                    Actor = actor,
                    Target = target,
                    Amount = dmg,
                    TimestampText = timestamp,
                    Timestamp = parsedTimestamp,
                    Raw = line
                });
            }

            Match healMatch = HealRegex.Match(payload);
            if (healMatch.Success && int.TryParse(healMatch.Groups["amount"].Value, out int heal))
            {
                string actor = NameNormalizer.Normalize(healMatch.Groups["actor"].Value);
                events.Add(new ParsedEvent
                {
                    EventType = "heal",
                    Actor = actor,
                    Amount = heal,
                    TimestampText = timestamp,
                    Timestamp = parsedTimestamp,
                    Raw = line
                });
            }

            Match talkMatch = TalkRegex.Match(trimmed);
            bool talkAdded = false;
            if (talkMatch.Success)
            {
                string actor = NameNormalizer.Normalize(talkMatch.Groups["actor"].Value);
                string? userRaw = talkMatch.Groups["user"].Success ? talkMatch.Groups["user"].Value : null;
                string? controller = null;
                if (!string.IsNullOrWhiteSpace(userRaw))
                {
                    controller = userRaw.Trim('[', ']', ' ');
                }
                events.Add(new ParsedEvent
                {
                    EventType = "talk",
                    Actor = actor,
                    TimestampText = talkMatch.Groups["ts"].Value,
                    Timestamp = TryParseTimestamp(talkMatch.Groups["ts"].Value),
                    ControllerUser = controller,
                    Raw = line
                });
                talkAdded = true;
            }
            else if (prefixMatch.Success && payload.Contains(":"))
            {
                // Fallback chat detection: only if it's not an action line
                string actorPart = payload.Split(':')[0];
                string lowered = payload.ToLowerInvariant();
                bool looksLikeAction = lowered.Contains(" damages ") || lowered.Contains(" attacks ") || lowered.Contains(" attempts ") || lowered.Contains(" heals ") || lowered.Contains(" cast");
                bool hasTalkTag = payload.IndexOf("[talk]", StringComparison.OrdinalIgnoreCase) >= 0;
                string actor = NameNormalizer.Normalize(actorPart);
                bool actorLooksLikeName = !string.IsNullOrWhiteSpace(actor) && Regex.IsMatch(actor, "^[A-Za-z][A-Za-z '\\-]*$");

                if (!talkAdded && !looksLikeAction && hasTalkTag && actorLooksLikeName)
                {
                    events.Add(new ParsedEvent
                    {
                        EventType = "talk",
                        Actor = actor,
                        TimestampText = timestamp,
                        Timestamp = parsedTimestamp,
                        Raw = line
                    });
                    talkAdded = true;
                }
            }

            return events;
        }

        private static DateTime? TryParseTimestamp(string? timestampText)
        {
            if (string.IsNullOrWhiteSpace(timestampText))
                return null;

            // Try common formats found in the logs
            string[] formats = new[]
            {
                "ddd MMM dd HH:mm:ss yyyy",
                "ddd MMM d HH:mm:ss yyyy",
                "ddd MMM dd HH:mm:ss",
                "ddd MMM d HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-ddTHH:mm:ss",
                "yyyy-MM-ddTHH:mm:sszzz",
                "o"
            };

            DateTime parsed;
            if (DateTime.TryParseExact(timestampText, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
            {
                // If no year was provided, assume current year
                if (parsed.Year == 1)
                {
                    parsed = new DateTime(DateTime.Now.Year, parsed.Month, parsed.Day, parsed.Hour, parsed.Minute, parsed.Second, parsed.Kind);
                }
                return parsed;
            }

            if (DateTime.TryParse(timestampText, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
            {
                if (parsed.Year == 1)
                {
                    parsed = new DateTime(DateTime.Now.Year, parsed.Month, parsed.Day, parsed.Hour, parsed.Minute, parsed.Second, parsed.Kind);
                }
                return parsed;
            }

            return null;
        }
    }

    public static class NameNormalizer
    {
        private static readonly Regex BracketedRegex = new(@"\[[^\]]*\]", RegexOptions.Compiled);
        private static readonly Regex AngleRegex = new(@"<[^>]*>", RegexOptions.Compiled);
        private static readonly Regex MultiSpaceRegex = new(@"\s+", RegexOptions.Compiled);

        public static string Normalize(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string cleaned = value;
            cleaned = BracketedRegex.Replace(cleaned, " ");
            cleaned = AngleRegex.Replace(cleaned, " ");

            int colonIndex = cleaned.IndexOf(':');
            if (colonIndex >= 0)
            {
                cleaned = cleaned.Substring(0, colonIndex);
            }

            cleaned = cleaned.Trim();
            cleaned = MultiSpaceRegex.Replace(cleaned, " ");
            return cleaned;
        }
    }

    public class EntityStats
    {
        public EntityStats(string name)
        {
            Name = name;
        }

        public string Name { get; }
        public int? MinHitAc { get; set; }
        public int? MaxMissAc { get; set; }
        public int? MinHitTouch { get; set; }
        public int? MaxMissTouch { get; set; }
        public bool? HasSpellResistance { get; set; }
        public int? ConcealmentPercent { get; set; }
        public HashSet<string> SpellsCast { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> Abilities { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public DateTime? FirstSeen { get; set; }
        public DateTime? LastSeen { get; set; }
        public long RollSum { get; set; }
        public long RollCount { get; set; }
        public DateTime? LastRollTimestamp { get; set; }
    }

    public class LogStatsTracker
    {
        private readonly Dictionary<string, EntityStats> _entities = new Dictionary<string, EntityStats>(StringComparer.OrdinalIgnoreCase);
        private RollDedupeState _dedupeState = RollDedupeState.Normalize(null);

        public DateTime? LastRollTimestamp => _dedupeState.LastRollTimestamp;

        public void Apply(ParsedEvent parsed)
        {
            switch (parsed.EventType)
            {
                case "attack":
                    ApplyAttack(parsed, isTouch: false);
                    break;
                case "touch_attack":
                    ApplyAttack(parsed, isTouch: true);
                    break;
                case "spell_resistance":
                    ApplySpellResistance(parsed);
                    break;
                case "concealment":
                    ApplyConcealment(parsed);
                    break;
                case "spell_cast":
                    ApplySpellCast(parsed);
                    break;
            }
        }

        private void ApplyAttack(ParsedEvent parsed, bool isTouch)
        {
            if (string.IsNullOrWhiteSpace(parsed.Target) || parsed.RawRoll == null || parsed.Hit == null)
                return;

            if (!CheckAndMarkRoll(parsed))
                return;

            EntityStats stats = GetOrCreate(parsed.Target);
            UpdateSeen(stats, parsed.Timestamp);
            string? ability = NormalizeAbility(parsed.Ability);
            if (ability != null && !string.IsNullOrWhiteSpace(parsed.Actor))
            {
                EntityStats abilityOwner = GetOrCreate(parsed.Actor);
                UpdateSeen(abilityOwner, parsed.Timestamp);
                abilityOwner.Abilities.Add(ability);
            }
            if (isTouch)
            {
                if (parsed.Hit.Value)
                {
                    stats.MinHitTouch = stats.MinHitTouch.HasValue ? Math.Min(stats.MinHitTouch.Value, parsed.RawRoll.Value) : parsed.RawRoll.Value;
                }
                else
                {
                    stats.MaxMissTouch = stats.MaxMissTouch.HasValue ? Math.Max(stats.MaxMissTouch.Value, parsed.RawRoll.Value) : parsed.RawRoll.Value;
                }
            }
            else
            {
                if (parsed.Hit.Value)
                {
                    stats.MinHitAc = stats.MinHitAc.HasValue ? Math.Min(stats.MinHitAc.Value, parsed.RawRoll.Value) : parsed.RawRoll.Value;
                }
                else
                {
                    stats.MaxMissAc = stats.MaxMissAc.HasValue ? Math.Max(stats.MaxMissAc.Value, parsed.RawRoll.Value) : parsed.RawRoll.Value;
                }
            }
        }

        private void ApplySpellResistance(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Target))
                return;

            EntityStats stats = GetOrCreate(parsed.Target);
            UpdateSeen(stats, parsed.Timestamp);
            if (parsed.HasSpellResistance)
            {
                stats.HasSpellResistance = true;
            }
        }

        private void ApplyConcealment(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Target))
                return;

            EntityStats stats = GetOrCreate(parsed.Target);
            UpdateSeen(stats, parsed.Timestamp);
            if (parsed.ConcealmentPercent.HasValue)
            {
                stats.ConcealmentPercent = parsed.ConcealmentPercent;
            }
        }

        private void ApplySpellCast(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Actor) || string.IsNullOrWhiteSpace(parsed.SpellName))
                return;

            EntityStats stats = GetOrCreate(parsed.Actor);
            UpdateSeen(stats, parsed.Timestamp);
            stats.SpellsCast.Add(parsed.SpellName);
        }

        private EntityStats GetOrCreate(string name)
        {
            if (!_entities.TryGetValue(name, out EntityStats? stats))
            {
                stats = new EntityStats(name);
                _entities[name] = stats;
            }

            return stats;
        }

        public void Reset()
        {
            _entities.Clear();
            _dedupeState = RollDedupeState.Normalize(null);
        }

        public void ClearAllRolls()
        {
            foreach (var entity in _entities.Values)
            {
                entity.RollSum = 0;
                entity.RollCount = 0;
            }
        }

        public void ClearDedupeState()
        {
            _dedupeState = RollDedupeState.Normalize(null);
        }

        public RollDedupeState SnapshotDedupeState()
        {
            return _dedupeState;
        }

        public void EnsureLastRollTimestamp(DateTime? timestamp)
        {
            if (!_dedupeState.LastRollTimestamp.HasValue && timestamp.HasValue)
            {
                _dedupeState.LastRollTimestamp = timestamp;
            }
        }

        public bool LoadFromDatabase(StatsRepository repo)
        {
            var list = repo.LoadEntities();
            if (list.Count == 0)
                return false;

            _entities.Clear();
            foreach (var e in list)
            {
                _entities[e.Name] = e;
            }
            return true;
        }

        public void SaveToDatabase(StatsRepository repo)
        {
            repo.SaveEntities(Snapshot());
        }

        public void LoadFromCsv(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                var lines = File.ReadLines(path).ToList();
                if (lines.Count == 0)
                    return;

                List<string> headerCells = ParseCsvLine(lines[0]);
                Dictionary<string, int> headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < headerCells.Count; i++)
                {
                    if (!string.IsNullOrWhiteSpace(headerCells[i]))
                    {
                        headerMap[headerCells[i].Trim()] = i;
                    }
                }

                string? GetCell(List<string> row, string columnName, params int[] fallbacks)
                {
                    if (headerMap.TryGetValue(columnName, out int idx) && idx < row.Count)
                        return row[idx];
                    foreach (int fb in fallbacks)
                    {
                        if (fb >= 0 && fb < row.Count)
                            return row[fb];
                    }
                    return null;
                }

                foreach (string line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    List<string> cells = ParseCsvLine(line);
                    if (cells.Count < 2)
                        continue; // need at least name + one stat

                    string name = NameNormalizer.Normalize(cells[0]);
                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    EntityStats stats = GetOrCreate(name);
                    MergeStat(stats, ParseNullableInt(cells[1]), touch: false, hit: true);
                    MergeStat(stats, ParseNullableInt(cells[2]), touch: false, hit: false);
                    MergeStat(stats, ParseNullableInt(cells[3]), touch: true, hit: true);
                    MergeStat(stats, ParseNullableInt(cells[4]), touch: true, hit: false);

                    bool? hasSr = ParseNullableBool(cells[5]);
                    if (hasSr.HasValue && hasSr.Value)
                    {
                        stats.HasSpellResistance = true;
                    }
                    else if (!stats.HasSpellResistance.HasValue)
                    {
                        stats.HasSpellResistance = hasSr;
                    }

                    int? conceal = ParseNullableInt(GetCell(cells, "concealment_percent", 6));
                    if (conceal.HasValue)
                    {
                        stats.ConcealmentPercent = stats.ConcealmentPercent ?? conceal;
                    }

                    string? spellsCell = GetCell(cells, "spells_cast", 7);
                    if (!string.IsNullOrWhiteSpace(spellsCell))
                    {
                        foreach (string spell in spellsCell.Split('|', StringSplitOptions.RemoveEmptyEntries))
                        {
                            stats.SpellsCast.Add(spell.Trim());
                        }
                    }

                    string? abilitiesCell = GetCell(cells, "abilities");
                    if (!string.IsNullOrWhiteSpace(abilitiesCell))
                    {
                        foreach (var ability in abilitiesCell.Split('|', StringSplitOptions.RemoveEmptyEntries))
                        {
                            string? normalized = NormalizeAbility(ability);
                            if (normalized != null)
                            {
                                stats.Abilities.Add(normalized);
                            }
                        }
                    }

                    string? firstCell = GetCell(cells, "first_seen", 9, 8);
                    string? lastCell = GetCell(cells, "last_seen", 10, 9);
                    if (!string.IsNullOrWhiteSpace(firstCell) || !string.IsNullOrWhiteSpace(lastCell))
                    {
                        DateTime? first = ParseNullableDate(firstCell);
                        DateTime? last = ParseNullableDate(lastCell);
                        if (first.HasValue)
                            stats.FirstSeen = stats.FirstSeen.HasValue ? (first < stats.FirstSeen ? first : stats.FirstSeen) : first;
                        if (last.HasValue)
                            stats.LastSeen = stats.LastSeen.HasValue ? (last > stats.LastSeen ? last : stats.LastSeen) : last;
                    }

                    string? rollSumCell = GetCell(cells, "roll_sum", 11, 10);
                    string? rollCountCell = GetCell(cells, "roll_count", 12, 11);
                    if (!string.IsNullOrWhiteSpace(rollSumCell) || !string.IsNullOrWhiteSpace(rollCountCell))
                    {
                        long? rollSum = ParseNullableLong(rollSumCell);
                        long? rollCount = ParseNullableLong(rollCountCell);
                        if (rollSum.HasValue)
                            stats.RollSum = rollSum.Value;
                        if (rollCount.HasValue)
                            stats.RollCount = rollCount.Value;
                    }
                }
            }
            catch (Exception)
            {
                // Ignore load errors to keep monitoring running
            }
        }

        public IReadOnlyCollection<EntityStats> Snapshot()
        {
            return _entities.Values.ToList();
        }

        public void ResetEntity(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            EntityStats stats = GetOrCreate(name);
            stats.MinHitAc = null;
            stats.MaxMissAc = null;
            stats.MinHitTouch = null;
            stats.MaxMissTouch = null;
            stats.HasSpellResistance = null;
            stats.ConcealmentPercent = null;
            stats.SpellsCast.Clear();
            stats.Abilities.Clear();
            stats.FirstSeen = null;
            stats.LastSeen = null;
            stats.RollSum = 0;
            stats.RollCount = 0;
            stats.LastRollTimestamp = null;
        }

        public void DeleteEntity(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            _entities.Remove(name);
        }

        public void SaveToCsv(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            CsvWriter.Write(path, Snapshot());
        }

        public EntityStats? Get(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            _entities.TryGetValue(name, out EntityStats? stats);
            return stats;
        }

        private static void UpdateSeen(EntityStats stats, DateTime? timestamp)
        {
            DateTime seen = timestamp ?? DateTime.Now;
            if (!stats.FirstSeen.HasValue || seen < stats.FirstSeen.Value)
            {
                stats.FirstSeen = seen;
            }

            if (!stats.LastSeen.HasValue || seen > stats.LastSeen.Value)
            {
                stats.LastSeen = seen;
            }
        }

        public void RecordIncomingRoll(string name, int roll, DateTime? timestamp)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            if (roll < 1 || roll > 20)
                return; // only accept d20 rolls

            EntityStats stats = GetOrCreate(name);
            UpdateSeen(stats, timestamp);
            stats.RollSum += roll;
            stats.RollCount += 1;
            stats.LastRollTimestamp = timestamp ?? DateTime.Now;
        }

        private static void MergeStat(EntityStats stats, int? value, bool touch, bool hit)
        {
            if (!value.HasValue)
                return;

            if (touch)
            {
                if (hit)
                {
                    stats.MinHitTouch = stats.MinHitTouch.HasValue ? Math.Min(stats.MinHitTouch.Value, value.Value) : value;
                }
                else
                {
                    stats.MaxMissTouch = stats.MaxMissTouch.HasValue ? Math.Max(stats.MaxMissTouch.Value, value.Value) : value;
                }
            }
            else
            {
                if (hit)
                {
                    stats.MinHitAc = stats.MinHitAc.HasValue ? Math.Min(stats.MinHitAc.Value, value.Value) : value;
                }
                else
                {
                    stats.MaxMissAc = stats.MaxMissAc.HasValue ? Math.Max(stats.MaxMissAc.Value, value.Value) : value;
                }
            }
        }

        private static int? ParseNullableInt(string? value)
        {
            if (int.TryParse(value, out int result))
                return result;
            return null;
        }

        public bool CheckAndMarkRoll(ParsedEvent parsed)
        {
            string rawSignature = parsed.Raw ?? parsed.TimestampText ?? string.Empty;
            string lineHash = ComputeLineHashStatic(rawSignature);
            // Always use recent line hash to avoid double-count even without timestamp
            if (_dedupeState.RecentLineHashes.Contains(lineHash))
            {
                return false;
            }

            if (!parsed.Timestamp.HasValue)
            {
                // No timestamp: dedupe purely on line hash
                _dedupeState.RecentLineHashes.Add(lineHash);
                return true;
            }

            DateTime ts = parsed.Timestamp.Value;
            string rollKey = BuildRollKey(parsed, ts, rawSignature);

            if (_dedupeState.LastRollTimestamp.HasValue)
            {
                if (ts < _dedupeState.LastRollTimestamp.Value)
                    return false; // older than last processed roll

                if (ts == _dedupeState.LastRollTimestamp.Value)
                {
                    if (_dedupeState.RollKeys.Contains(rollKey))
                        return false; // duplicate within same timestamp
                }
                else
                {
                    _dedupeState.RollKeys.Clear();
                }
            }

            _dedupeState.LastRollTimestamp = ts;
            _dedupeState.RollKeys.Add(rollKey);
            _dedupeState.RecentLineHashes.Add(lineHash);
            TrimRecentLineHashes();
            return true;
        }

        private static string BuildRollKey(ParsedEvent parsed, DateTime ts, string? rawSignature = null)
        {
            string rawPart = string.IsNullOrWhiteSpace(rawSignature) ? string.Empty : rawSignature.Trim();
            return $"{ts:O}|{parsed.Actor}|{parsed.Target}|{parsed.RawRoll}|{parsed.Hit}|{parsed.IsTouchAttack}|{rawPart}";
        }

        private static string ComputeLineHash(string raw) => ComputeLineHashStatic(raw);

        public static string ComputeLineHashStatic(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return string.Empty;

            using var sha = SHA256.Create();
            byte[] bytes = Encoding.UTF8.GetBytes(raw);
            byte[] hash = sha.ComputeHash(bytes);
            return Convert.ToHexString(hash);
        }

        private void TrimRecentLineHashes()
        {
            const int maxHashes = 20000;
            if (_dedupeState.RecentLineHashes.Count > maxHashes)
            {
                // Drop oldest by removing an arbitrary chunk (not strictly ordered, but set is fine for our needs)
                _dedupeState.RecentLineHashes = _dedupeState.RecentLineHashes.TakeLast(maxHashes).ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
        }

        public void SetDedupeState(RollDedupeState state)
        {
            _dedupeState = RollDedupeState.Normalize(state);
        }

        internal static string? NormalizeAbility(string? ability)
        {
            if (string.IsNullOrWhiteSpace(ability))
                return null;

            string trimmed = ability.Trim();
            if (DateTime.TryParse(trimmed, out _))
                return null;

            return trimmed;
        }

        private static long? ParseNullableLong(string? value)
        {
            if (long.TryParse(value, out long result))
                return result;
            return null;
        }

        private static bool? ParseNullableBool(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            if (value.Equals("true", StringComparison.OrdinalIgnoreCase))
                return true;
            if (value.Equals("false", StringComparison.OrdinalIgnoreCase))
                return false;

            return null;
        }

        private static DateTime? ParseNullableDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            if (DateTime.TryParseExact(value, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime iso))
                return iso;

            if (DateTime.TryParse(value, out DateTime result))
                return result;

            return null;
        }

        private static List<string> ParseCsvLine(string line)
        {
            List<string> cells = new List<string>();
            StringBuilder current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (inQuotes)
                {
                    if (c == '"' && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else if (c == '"')
                    {
                        inQuotes = false;
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == '"')
                    {
                        inQuotes = true;
                    }
                    else if (c == ',')
                    {
                        cells.Add(current.ToString());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            cells.Add(current.ToString());
            return cells;
        }
    }

    public static class CsvWriter
    {
        public static void Write(string path, IReadOnlyCollection<EntityStats> stats)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("name,min_hit_ac,max_miss_ac,min_hit_touch,max_miss_touch,has_spell_resistance,concealment_percent,spells_cast,abilities,first_seen,last_seen,roll_sum,roll_count");

            foreach (EntityStats entity in stats.OrderBy(e => e.Name))
            {
                string spells = string.Join("|", entity.SpellsCast.OrderBy(s => s));
                string abilities = string.Join("|", entity.Abilities.OrderBy(s => s));
                string firstSeen = entity.FirstSeen.HasValue ? entity.FirstSeen.Value.ToString("o") : string.Empty;
                string lastSeen = entity.LastSeen.HasValue ? entity.LastSeen.Value.ToString("o") : string.Empty;
                builder.AppendLine(string.Join(",",
                    Escape(entity.Name),
                    ValueOrEmpty(entity.MinHitAc),
                    ValueOrEmpty(entity.MaxMissAc),
                    ValueOrEmpty(entity.MinHitTouch),
                    ValueOrEmpty(entity.MaxMissTouch),
                    entity.HasSpellResistance.HasValue ? (entity.HasSpellResistance.Value ? "true" : "false") : string.Empty,
                    ValueOrEmpty(entity.ConcealmentPercent),
                    Escape(spells),
                    Escape(abilities),
                    firstSeen,
                    lastSeen,
                    ValueOrEmpty(entity.RollSum),
                    ValueOrEmpty(entity.RollCount)));
            }

            File.WriteAllText(path, builder.ToString());
        }

        private static string ValueOrEmpty(int? value) => value.HasValue ? value.Value.ToString() : string.Empty;
        private static string ValueOrEmpty(long? value) => value.HasValue ? value.Value.ToString() : string.Empty;

        private static string Escape(string value)
        {
            if (value.Contains(",") || value.Contains("\""))
            {
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }

            return value;
        }
    }

    public class RollDedupeState
    {
        public DateTime? LastRollTimestamp { get; set; }
        public HashSet<string> RollKeys { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> RecentLineHashes { get; set; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public static RollDedupeState Normalize(RollDedupeState? state)
        {
            state ??= new RollDedupeState();
            state.RollKeys ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            state.RecentLineHashes ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return state;
        }
    }

    public class RollDedupeStore
    {
        private readonly string _path;
        private readonly Dictionary<string, RollDedupeState> _states = new Dictionary<string, RollDedupeState>(StringComparer.OrdinalIgnoreCase);
        private const int MaxRecentLineHashes = 20000;

        public RollDedupeStore(string path)
        {
            _path = path;
        }

        public RollDedupeState GetOrCreate(string filePath)
        {
            string key = Path.GetFullPath(filePath);
            if (!_states.TryGetValue(key, out var state))
            {
                state = new RollDedupeState();
                _states[key] = state;
            }
            return RollDedupeState.Normalize(state);
        }

        public void Set(string filePath, RollDedupeState state)
        {
            string key = Path.GetFullPath(filePath);
            RollDedupeState normalized = RollDedupeState.Normalize(state);
            // Trim to keep memory bounded
            if (normalized.RecentLineHashes.Count > MaxRecentLineHashes)
            {
                normalized.RecentLineHashes = normalized.RecentLineHashes.TakeLast(MaxRecentLineHashes).ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
            _states[key] = normalized;
        }

        public void Save()
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(_states, options);
                File.WriteAllText(_path, json);
            }
            catch
            {
                // Ignore persistence errors silently
            }
        }

        public void ClearAll()
        {
            _states.Clear();
            Save();
        }

        public static RollDedupeStore Load(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    var dict = JsonSerializer.Deserialize<Dictionary<string, RollDedupeState>>(json) ?? new Dictionary<string, RollDedupeState>(StringComparer.OrdinalIgnoreCase);
                    var store = new RollDedupeStore(path);
                    foreach (var kvp in dict)
                    {
                        store._states[kvp.Key] = RollDedupeState.Normalize(kvp.Value);
                    }
                    return store;
                }
            }
            catch
            {
                // Fall through to new store on error
            }

            return new RollDedupeStore(path);
        }
    }

    public class DiceSummary
    {
        public string AverageDisplay { get; }
        public string LastRollDisplay { get; }

        private DiceSummary(string avg, string last)
        {
            AverageDisplay = avg;
            LastRollDisplay = last;
        }

        public static DiceSummary FromSnapshot(IEnumerable<EntityStats> snapshot, DateTime? lastRollTimestamp)
        {
            long sum = snapshot.Sum(s => s.RollSum);
            long count = snapshot.Sum(s => s.RollCount);
            string avgText = count > 0 ? $"{FormatAverage(sum, count)} (out of {count})" : "-";
            string lastText = lastRollTimestamp?.ToString("yyyy-MM-dd HH:mm:ss") ?? "-";
            return new DiceSummary(avgText, lastText);
        }

        private static string FormatAverage(long sum, long count)
        {
            if (count <= 0) return "-";
            double avg = (double)sum / (double)count;
            return avg.ToString("0.0");
        }
    }

    public class IngestionState
    {
        public string? FilePath { get; set; }
        public long LastOffset { get; set; }
        public long FileLength { get; set; }
        public DateTime LastWriteTimeUtc { get; set; }
        public List<string> RecentLineHashes { get; set; } = new List<string>();
    }

    public class IngestionStateStore
    {
        private readonly string _path;
        private Dictionary<string, IngestionState> _states = new Dictionary<string, IngestionState>(StringComparer.OrdinalIgnoreCase);
        private const int MaxHashes = 500;

        public IngestionStateStore(string path)
        {
            _path = path;
            Load();
        }

        public IngestionState? Get(string filePath)
        {
            string key = Path.GetFullPath(filePath);
            if (_states.TryGetValue(key, out var state))
            {
                return state;
            }
            return null;
        }

        public void Clear(string filePath)
        {
            string key = Path.GetFullPath(filePath);
            if (_states.Remove(key))
            {
                Save();
            }
        }

        public void ClearAll()
        {
            _states.Clear();
            Save();
        }

        public void Set(string filePath, IngestionState state)
        {
            string key = Path.GetFullPath(filePath);
            state.FilePath = key;
            if (state.RecentLineHashes.Count > MaxHashes)
            {
                state.RecentLineHashes = state.RecentLineHashes.TakeLast(MaxHashes).ToList();
            }
            _states[key] = state;
        }

        public void Save()
        {
            try
            {
                string dir = Path.GetDirectoryName(_path) ?? AppContext.BaseDirectory;
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                string tmp = _path + ".tmp";
                var opts = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(tmp, JsonSerializer.Serialize(_states, opts));
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
                File.Move(tmp, _path);
            }
            catch
            {
                // ignore save errors
            }
        }

        private void Load()
        {
            try
            {
                if (File.Exists(_path))
                {
                    string json = File.ReadAllText(_path);
                    _states = JsonSerializer.Deserialize<Dictionary<string, IngestionState>>(json) ?? new Dictionary<string, IngestionState>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kvp in _states)
                    {
                        if (kvp.Value.RecentLineHashes.Count > MaxHashes)
                        {
                            kvp.Value.RecentLineHashes = kvp.Value.RecentLineHashes.TakeLast(MaxHashes).ToList();
                        }
                    }
                }
            }
            catch
            {
                _states = new Dictionary<string, IngestionState>(StringComparer.OrdinalIgnoreCase);
            }
        }
    }

    public class StatsRepository
    {
        private readonly string _dbPath;

            public StatsRepository(string dbPath)
        {
            _dbPath = dbPath;
        }

        public void EnsureDatabase()
        {
            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            connection.Open();
            using var cmd = connection.CreateCommand();
                cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS entities(
    name TEXT PRIMARY KEY,
    min_hit_ac INTEGER,
    max_miss_ac INTEGER,
    min_hit_touch INTEGER,
    max_miss_touch INTEGER,
    has_spell_resistance INTEGER,
    concealment_percent INTEGER,
    first_seen TEXT,
    last_seen TEXT,
    roll_sum INTEGER,
    roll_count INTEGER,
    last_roll_ts TEXT
);
CREATE TABLE IF NOT EXISTS spells(
    entity_name TEXT NOT NULL,
    spell TEXT NOT NULL,
    PRIMARY KEY(entity_name, spell)
);
CREATE TABLE IF NOT EXISTS abilities(
    entity_name TEXT NOT NULL,
    ability TEXT NOT NULL,
    PRIMARY KEY(entity_name, ability)
);
";
            cmd.ExecuteNonQuery();

            // Migration: add last_roll_ts column if missing
            using var checkCmd = connection.CreateCommand();
            checkCmd.CommandText = "PRAGMA table_info(entities);";
            bool hasLastRoll = false;
            using (var reader = checkCmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string colName = reader.GetString(1);
                    if (colName.Equals("last_roll_ts", StringComparison.OrdinalIgnoreCase))
                    {
                        hasLastRoll = true;
                        break;
                    }
                }
            }

            if (!hasLastRoll)
            {
                using var alterCmd = connection.CreateCommand();
                alterCmd.CommandText = "ALTER TABLE entities ADD COLUMN last_roll_ts TEXT;";
                alterCmd.ExecuteNonQuery();
            }
        }

        public List<EntityStats> LoadEntities()
        {
            List<EntityStats> list = new List<EntityStats>();
            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            connection.Open();

            Dictionary<string, EntityStats> map = new Dictionary<string, EntityStats>(StringComparer.OrdinalIgnoreCase);

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = "SELECT name,min_hit_ac,max_miss_ac,min_hit_touch,max_miss_touch,has_spell_resistance,concealment_percent,first_seen,last_seen,roll_sum,roll_count,last_roll_ts FROM entities";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    EntityStats stats = new EntityStats(name)
                    {
                        MinHitAc = reader.IsDBNull(1) ? null : reader.GetInt32(1),
                        MaxMissAc = reader.IsDBNull(2) ? null : reader.GetInt32(2),
                        MinHitTouch = reader.IsDBNull(3) ? null : reader.GetInt32(3),
                        MaxMissTouch = reader.IsDBNull(4) ? null : reader.GetInt32(4),
                        HasSpellResistance = reader.IsDBNull(5) ? null : reader.GetInt32(5) == 1,
                        ConcealmentPercent = reader.IsDBNull(6) ? null : reader.GetInt32(6),
                        FirstSeen = reader.IsDBNull(7) ? null : ParseNullableDate(reader.GetString(7)),
                        LastSeen = reader.IsDBNull(8) ? null : ParseNullableDate(reader.GetString(8)),
                        RollSum = reader.IsDBNull(9) ? 0 : reader.GetInt64(9),
                        RollCount = reader.IsDBNull(10) ? 0 : reader.GetInt64(10),
                        LastRollTimestamp = reader.FieldCount > 11 && !reader.IsDBNull(11) ? ParseNullableDate(reader.GetString(11)) : null
                    };
                    map[name] = stats;
                    list.Add(stats);
                }
            }

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = "SELECT entity_name, spell FROM spells";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    string spell = reader.GetString(1);
                    if (map.TryGetValue(name, out var stats))
                    {
                        stats.SpellsCast.Add(spell);
                    }
                }
            }

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = "SELECT entity_name, ability FROM abilities";
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string name = reader.GetString(0);
                    string ability = reader.GetString(1);
                    if (map.TryGetValue(name, out var stats))
                    {
                        string? normalized = LogStatsTracker.NormalizeAbility(ability);
                        if (normalized != null)
                        {
                            stats.Abilities.Add(normalized);
                        }
                    }
                }
            }

            return list;
        }

        private static DateTime? ParseNullableDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            if (DateTime.TryParseExact(value, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime iso))
                return iso;

            if (DateTime.TryParse(value, out DateTime result))
                return result;

            return null;
        }

        public void SaveEntities(IReadOnlyCollection<EntityStats> entities)
        {
            using var connection = new SqliteConnection($"Data Source={_dbPath}");
            connection.Open();
            using var tx = connection.BeginTransaction();

            using (var clearCmd = connection.CreateCommand())
            {
                clearCmd.CommandText = "DELETE FROM spells; DELETE FROM abilities; DELETE FROM entities;";
                clearCmd.ExecuteNonQuery();
            }

            using (var insertCmd = connection.CreateCommand())
            {
                insertCmd.CommandText = @"INSERT INTO entities(name,min_hit_ac,max_miss_ac,min_hit_touch,max_miss_touch,has_spell_resistance,concealment_percent,first_seen,last_seen,roll_sum,roll_count,last_roll_ts)
VALUES ($name,$min_hit_ac,$max_miss_ac,$min_hit_touch,$max_miss_touch,$has_sr,$conceal,$first,$last,$roll_sum,$roll_count,$last_roll_ts);";
                var pName = insertCmd.CreateParameter(); pName.ParameterName = "$name"; insertCmd.Parameters.Add(pName);
                var pMinAc = insertCmd.CreateParameter(); pMinAc.ParameterName = "$min_hit_ac"; insertCmd.Parameters.Add(pMinAc);
                var pMaxAc = insertCmd.CreateParameter(); pMaxAc.ParameterName = "$max_miss_ac"; insertCmd.Parameters.Add(pMaxAc);
                var pMinTouch = insertCmd.CreateParameter(); pMinTouch.ParameterName = "$min_hit_touch"; insertCmd.Parameters.Add(pMinTouch);
                var pMaxTouch = insertCmd.CreateParameter(); pMaxTouch.ParameterName = "$max_miss_touch"; insertCmd.Parameters.Add(pMaxTouch);
                var pHasSr = insertCmd.CreateParameter(); pHasSr.ParameterName = "$has_sr"; insertCmd.Parameters.Add(pHasSr);
                var pCon = insertCmd.CreateParameter(); pCon.ParameterName = "$conceal"; insertCmd.Parameters.Add(pCon);
                var pFirst = insertCmd.CreateParameter(); pFirst.ParameterName = "$first"; insertCmd.Parameters.Add(pFirst);
                var pLast = insertCmd.CreateParameter(); pLast.ParameterName = "$last"; insertCmd.Parameters.Add(pLast);
                var pRollSum = insertCmd.CreateParameter(); pRollSum.ParameterName = "$roll_sum"; insertCmd.Parameters.Add(pRollSum);
                var pRollCount = insertCmd.CreateParameter(); pRollCount.ParameterName = "$roll_count"; insertCmd.Parameters.Add(pRollCount);
                var pLastRollTs = insertCmd.CreateParameter(); pLastRollTs.ParameterName = "$last_roll_ts"; insertCmd.Parameters.Add(pLastRollTs);

                foreach (var e in entities)
                {
                    pName.Value = e.Name;
                    pMinAc.Value = (object?)e.MinHitAc ?? DBNull.Value;
                    pMaxAc.Value = (object?)e.MaxMissAc ?? DBNull.Value;
                    pMinTouch.Value = (object?)e.MinHitTouch ?? DBNull.Value;
                    pMaxTouch.Value = (object?)e.MaxMissTouch ?? DBNull.Value;
                    pHasSr.Value = e.HasSpellResistance.HasValue ? (e.HasSpellResistance.Value ? 1 : 0) : (object)DBNull.Value;
                    pCon.Value = (object?)e.ConcealmentPercent ?? DBNull.Value;
                    pFirst.Value = e.FirstSeen.HasValue ? e.FirstSeen.Value.ToString("o") : (object)DBNull.Value;
                    pLast.Value = e.LastSeen.HasValue ? e.LastSeen.Value.ToString("o") : (object)DBNull.Value;
                    pRollSum.Value = e.RollSum;
                    pRollCount.Value = e.RollCount;
                    pLastRollTs.Value = e.LastRollTimestamp.HasValue ? e.LastRollTimestamp.Value.ToString("o") : (object)DBNull.Value;
                    insertCmd.ExecuteNonQuery();
                }
            }

            using (var spellsCmd = connection.CreateCommand())
            {
                spellsCmd.CommandText = "INSERT INTO spells(entity_name, spell) VALUES ($name,$spell);";
                var pName = spellsCmd.CreateParameter(); pName.ParameterName = "$name"; spellsCmd.Parameters.Add(pName);
                var pSpell = spellsCmd.CreateParameter(); pSpell.ParameterName = "$spell"; spellsCmd.Parameters.Add(pSpell);

                foreach (var e in entities)
                {
                    pName.Value = e.Name;
                    foreach (var spell in e.SpellsCast)
                    {
                        pSpell.Value = spell;
                        spellsCmd.ExecuteNonQuery();
                    }
                }
            }

            using (var abilitiesCmd = connection.CreateCommand())
            {
                abilitiesCmd.CommandText = "INSERT INTO abilities(entity_name, ability) VALUES ($name,$ability);";
                var pName = abilitiesCmd.CreateParameter(); pName.ParameterName = "$name"; abilitiesCmd.Parameters.Add(pName);
                var pAbility = abilitiesCmd.CreateParameter(); pAbility.ParameterName = "$ability"; abilitiesCmd.Parameters.Add(pAbility);

                foreach (var e in entities)
                {
                    pName.Value = e.Name;
                    foreach (var ability in e.Abilities)
                    {
                        pAbility.Value = ability;
                        abilitiesCmd.ExecuteNonQuery();
                    }
                }
            }

            tx.Commit();
        }
    }

    public class WindowSettings
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public WindowState State { get; set; }
        public string? LastPlayerName { get; set; }
    }

    public class CombatSession
    {
        public string Name { get; set; } = string.Empty;
        public List<CombatSlot> Slots { get; } = new List<CombatSlot>();
        public Dictionary<string, CombatAggregate> Aggregates { get; } = new Dictionary<string, CombatAggregate>(StringComparer.OrdinalIgnoreCase);
        public CombatAggregate Overall { get; } = new CombatAggregate { Name = "Overall" };
        public DateTime Start { get; set; }
        public DateTime LastActivity { get; set; }
    }

    public class CombatDisplay
    {
        public string Name { get; set; } = string.Empty;
        public string DamageText { get; set; } = string.Empty;
        public string DpsText { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public string DamageTakenText { get; set; } = string.Empty;
        public Brush DamageTakenBrush { get; set; } = Brushes.White;
        public DateTime LastSeen { get; set; }
    }

    public class CombatSlot
    {
        public string Name { get; set; } = string.Empty;
        public DateTime LastSeen { get; set; }
    }

        public class CombatAggregate
        {
            public string Name { get; set; } = string.Empty;
            public double Damage { get; private set; }
            public double DamageTaken { get; private set; }
            public double Healing { get; private set; }
            public DateTime? FirstAction { get; private set; }
            public DateTime? LastAction { get; private set; }
            public long IncomingRollSum { get; private set; }
            public long IncomingRollCount { get; private set; }
            public long OutgoingRollSum { get; private set; }
            public long OutgoingRollCount { get; private set; }

            public void RegisterDamage(double amount, DateTime timestamp)
            {
                Damage += amount;
                Touch(timestamp);
        }

        public void RegisterDamageTaken(double amount, DateTime timestamp)
        {
            DamageTaken += amount;
            Touch(timestamp);
        }

        public void RegisterHeal(double amount, DateTime timestamp)
            {
                Healing += amount;
                Touch(timestamp);
            }

            public void RegisterIncomingRoll(int roll, DateTime timestamp)
            {
                IncomingRollSum += roll;
                IncomingRollCount += 1;
                Touch(timestamp);
            }

            public void RegisterOutgoingRoll(int roll, DateTime timestamp)
            {
                OutgoingRollSum += roll;
                OutgoingRollCount += 1;
                Touch(timestamp);
            }

            private void Touch(DateTime timestamp)
            {
                if (!FirstAction.HasValue || timestamp < FirstAction.Value)
                {
                    FirstAction = timestamp;
            }

            if (!LastAction.HasValue || timestamp > LastAction.Value)
            {
                LastAction = timestamp;
            }
        }

        public double ComputeDps(DateTime now)
        {
            if (!FirstAction.HasValue)
                return 0;

            DateTime end = LastAction ?? now;
            if ((now - end).TotalSeconds > 15)
            {
                // Stale combat, lock DPS to last action time
                now = end;
            }

            double seconds = Math.Max(1, (now - FirstAction.Value).TotalSeconds);
            return Damage / seconds;
        }
    }

    public class PlayerDetector
    {
        private readonly Dictionary<string, int> _scores = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private string? _currentGuess;
        private int _eventCount;

        public string? Register(ParsedEvent parsed)
        {
            if (string.IsNullOrWhiteSpace(parsed.Actor))
                return _currentGuess;

            int weight = parsed.EventType switch
            {
                "attack" => 3,
                "touch_attack" => 3,
                "damage" => 2,
                "heal" => 2,
                "spell_cast" => 1,
                _ => 0
            };

            if (weight == 0)
                return _currentGuess;

            _eventCount++;
            if (_scores.ContainsKey(parsed.Actor))
                _scores[parsed.Actor] += weight;
            else
                _scores[parsed.Actor] = weight;

            if (_eventCount >= 5)
            {
                string best = _scores.OrderByDescending(kvp => kvp.Value).ThenBy(kvp => kvp.Key).First().Key;
                if (!string.Equals(best, _currentGuess, StringComparison.OrdinalIgnoreCase))
                {
                    _currentGuess = best;
                }
            }

            return _currentGuess;
        }

        public void SeedBias(string name, int weight = 20)
        {
            if (string.IsNullOrWhiteSpace(name))
                return;

            string normalized = NameNormalizer.Normalize(name);
            _scores[normalized] = weight;
            _currentGuess = normalized;
        }

        public void Reset()
        {
            _scores.Clear();
            _currentGuess = null;
            _eventCount = 0;
        }
    }

    public class SocialProfile
    {
        public string Name { get; set; } = string.Empty;
        public bool HaveMet { get; set; }
        public bool KnowName { get; set; }
        public string? Alias { get; set; }
        public DateTime? FirstSeen { get; set; }
        public DateTime? LastSeen { get; set; }
        public string? Notes { get; set; }
        public List<string> Users { get; set; } = new List<string>();
    }

    // Core data that is shared across all player profiles
    public class SocialCore
    {
        public string Name { get; set; } = string.Empty;
        public DateTime? FirstSeen { get; set; }
        public DateTime? LastSeen { get; set; }
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Alias { get; set; }
        public List<string> Users { get; set; } = new List<string>();
    }

    // Per-player overlay data
    public class SocialOverlay
    {
        public string Name { get; set; } = string.Empty;
        public bool HaveMet { get; set; }
        public bool KnowName { get; set; }
        public string? Notes { get; set; }
    }

    public class ReminderItem : INotifyPropertyChanged
    {
        private bool _enabled;
        private string _text = string.Empty;
        private bool _caseSensitive;
        private bool _wholeWord;
        private string _lastMatchPrefix = string.Empty;
        private string _lastMatchMatch = string.Empty;
        private string _lastMatchSuffix = string.Empty;
        private bool _hasMatch = false;

        public bool Enabled
        {
            get => _enabled;
            set => SetField(ref _enabled, value);
        }

        public string Text
        {
            get => _text;
            set
            {
                if (SetField(ref _text, value))
                {
                    ClearMatch();
                }
            }
        }

        public bool CaseSensitive
        {
            get => _caseSensitive;
            set => SetField(ref _caseSensitive, value);
        }

        public bool WholeWord
        {
            get => _wholeWord;
            set => SetField(ref _wholeWord, value);
        }

        public string LastMatchPrefix
        {
            get => _lastMatchPrefix;
            private set => SetField(ref _lastMatchPrefix, value);
        }

        public string LastMatchMatch
        {
            get => _lastMatchMatch;
            private set => SetField(ref _lastMatchMatch, value);
        }

        public string LastMatchSuffix
        {
            get => _lastMatchSuffix;
            private set => SetField(ref _lastMatchSuffix, value);
        }

        public bool HasMatch
        {
            get => _hasMatch;
            private set => SetField(ref _hasMatch, value);
        }

        public void UpdateMatch(string line, int start, int length)
        {
            if (start < 0 || length <= 0 || start >= line.Length)
            {
                ClearMatch();
                return;
            }

            string prefix = start > 0 ? line.Substring(0, start) : string.Empty;
            string match = line.Substring(start, Math.Min(length, line.Length - start));
            string suffix = start + length < line.Length ? line.Substring(start + length) : string.Empty;

            LastMatchPrefix = prefix;
            LastMatchMatch = match;
            LastMatchSuffix = suffix;
            HasMatch = !string.IsNullOrEmpty(match);
        }

        public void ClearMatch()
        {
            LastMatchPrefix = string.Empty;
            LastMatchMatch = string.Empty;
            LastMatchSuffix = string.Empty;
            HasMatch = false;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;
            field = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            return true;
        }
    }

    public class SocialStore
    {
        private readonly string _overlayPath;
        private readonly string _corePath;
        private readonly Dictionary<string, SocialCore> _cores = new Dictionary<string, SocialCore>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Dictionary<string, SocialOverlay>> _overlaysByProfile = new Dictionary<string, Dictionary<string, SocialOverlay>>(StringComparer.OrdinalIgnoreCase);
        private string _activeProfile = "Default";
        private static readonly Regex NameFilter = new Regex("^[\\p{L}\\p{M}][\\p{L}\\p{M} '\\-]*$", RegexOptions.Compiled);

        public SocialStore(string overlayPath, string corePath)
        {
            _overlayPath = overlayPath;
            _corePath = corePath;
        }

        public string ActiveProfile => _activeProfile;

        public void SetActiveProfile(string profileKey)
        {
            if (string.IsNullOrWhiteSpace(profileKey))
                profileKey = "Default";
            _activeProfile = profileKey.Trim();
            EnsureOverlayMap(_activeProfile);
        }

        private Dictionary<string, SocialOverlay> CurrentOverlays => EnsureOverlayMap(_activeProfile);

        public void Load()
        {
            _cores.Clear();
            _overlaysByProfile.Clear();
            _activeProfile = "Default";

            bool coresChanged = LoadCoresFromFile(_corePath);
            coresChanged |= LoadOverlaysFromFile(_overlayPath, importCores: true);

            EnsureOverlayMap(_activeProfile);

            if (coresChanged)
            {
                SaveCores();
            }
        }

        public void Save()
        {
            SaveCores();
            SaveOverlays();
        }

        private void SaveCores()
        {
            try
            {
                EnsureDirectory(_corePath);
                var options = new JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };
                var dto = new
                {
                    cores = _cores
                };
                File.WriteAllText(_corePath, JsonSerializer.Serialize(dto, options));
            }
            catch
            {
                // ignore save errors
            }
        }

        private void SaveOverlays()
        {
            try
            {
                EnsureDirectory(_overlayPath);
                var options = new JsonSerializerOptions { WriteIndented = true, DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull };
                var dto = new
                {
                    overlays = _overlaysByProfile,
                    activeProfile = _activeProfile
                };
                File.WriteAllText(_overlayPath, JsonSerializer.Serialize(dto, options));
            }
            catch
            {
                // ignore save errors
            }
        }

        private static void EnsureDirectory(string path)
        {
            string? dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        private static bool TryGetPropertyCaseInsensitive(JsonElement element, string propertyName, out JsonElement value)
        {
            if (element.ValueKind == JsonValueKind.Object)
            {
                if (element.TryGetProperty(propertyName, out value))
                {
                    return true;
                }

                foreach (var prop in element.EnumerateObject())
                {
                    if (string.Equals(prop.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        value = prop.Value;
                        return true;
                    }
                }
            }

            value = default;
            return false;
        }

        private bool LoadCoresFromFile(string path)
        {
            if (!File.Exists(path))
                return false;

            bool changed = false;
            try
            {
                string json = File.ReadAllText(path);
                using var doc = JsonDocument.Parse(json);
                if (TryGetPropertyCaseInsensitive(doc.RootElement, "cores", out var coresElement))
                {
                    changed |= ImportCoresFromElement(coresElement);
                }
                else
                {
                    var legacy = JsonSerializer.Deserialize<Dictionary<string, SocialProfile>>(json);
                    if (legacy != null)
                    {
                        foreach (var kvp in legacy)
                        {
                            changed |= ImportCoreFromProfile(kvp.Key, kvp.Value);
                        }
                    }
                }
            }
            catch
            {
                // ignore load errors
            }

            return changed;
        }

        private bool LoadOverlaysFromFile(string path, bool importCores)
        {
            if (!File.Exists(path))
                return false;

            bool coresChanged = false;
            try
            {
                string json = File.ReadAllText(path);
                using var doc = JsonDocument.Parse(json);
                if (TryGetPropertyCaseInsensitive(doc.RootElement, "overlays", out var overlaysElement))
                {
                    foreach (var profileProp in overlaysElement.EnumerateObject())
                    {
                        var map = JsonSerializer.Deserialize<Dictionary<string, SocialOverlay>>(profileProp.Value.GetRawText());
                        if (map != null)
                        {
                            var normalized = NormalizeOverlayMap(map);
                            _overlaysByProfile[profileProp.Name] = normalized;

                            if (importCores)
                            {
                                foreach (var name in normalized.Keys)
                                {
                                    coresChanged |= UpsertCore(new SocialCore { Name = name });
                                }
                            }
                        }

                        if (importCores)
                        {
                            coresChanged |= ImportAliasFromOverlay(profileProp.Value);
                        }
                    }

                    if (TryGetPropertyCaseInsensitive(doc.RootElement, "activeProfile", out var activeProp) && activeProp.ValueKind == JsonValueKind.String)
                    {
                        _activeProfile = activeProp.GetString() ?? "Default";
                    }
                    else
                    {
                        _activeProfile = "Default";
                    }

                    if (importCores && TryGetPropertyCaseInsensitive(doc.RootElement, "cores", out var coresElement))
                    {
                        coresChanged |= ImportCoresFromElement(coresElement);
                    }
                }
                else
                {
                    var legacy = JsonSerializer.Deserialize<Dictionary<string, SocialProfile>>(json);
                    if (legacy != null)
                    {
                        var overlayMap = new Dictionary<string, SocialOverlay>(StringComparer.OrdinalIgnoreCase);
                        foreach (var kvp in legacy)
                        {
                            if (!IsValidName(kvp.Key))
                                continue;

                            overlayMap[kvp.Key] = new SocialOverlay
                            {
                                Name = kvp.Key,
                                HaveMet = kvp.Value?.HaveMet ?? false,
                                KnowName = kvp.Value?.KnowName ?? false,
                                Notes = kvp.Value?.Notes
                            };

                            if (importCores)
                            {
                                coresChanged |= ImportCoreFromProfile(kvp.Key, kvp.Value);
                            }
                        }

                        _overlaysByProfile["Default"] = overlayMap;
                        _activeProfile = "Default";
                    }
                }
            }
            catch
            {
                // ignore load errors
            }

            return coresChanged;
        }

        private bool ImportCoresFromElement(JsonElement coresElement)
        {
            bool changed = false;
            foreach (var coreProp in coresElement.EnumerateObject())
            {
                var core = JsonSerializer.Deserialize<SocialCore>(coreProp.Value.GetRawText());
                if (core != null && IsValidName(coreProp.Name))
                {
                    core.Name = coreProp.Name;
                    core = NormalizeCore(core);
                    changed |= UpsertCore(core);
                }
            }

            return changed;
        }

        private bool ImportAliasFromOverlay(JsonElement overlayMapElement)
        {
            bool changed = false;
            foreach (var overlayProp in overlayMapElement.EnumerateObject())
            {
                if (!IsValidName(overlayProp.Name) || overlayProp.Value.ValueKind != JsonValueKind.Object)
                    continue;

                JsonElement aliasProp;
                if (!overlayProp.Value.TryGetProperty("Alias", out aliasProp) && !overlayProp.Value.TryGetProperty("alias", out aliasProp))
                    continue;

                if (aliasProp.ValueKind == JsonValueKind.String)
                {
                    string? alias = aliasProp.GetString();
                    if (!string.IsNullOrWhiteSpace(alias))
                    {
                        SocialCore core = EnsureCore(overlayProp.Name);
                        if (string.IsNullOrWhiteSpace(core.Alias))
                        {
                            core.Alias = alias.Trim();
                            changed = true;
                        }
                    }
                }
            }

            return changed;
        }

        private bool ImportCoreFromProfile(string name, SocialProfile? profile)
        {
            if (!IsValidName(name))
                return false;

            SocialCore core = new SocialCore
            {
                Name = name,
                FirstSeen = profile?.FirstSeen,
                LastSeen = profile?.LastSeen,
                Users = profile?.Users ?? new List<string>(),
                Alias = profile?.Alias
            };

            return UpsertCore(core);
        }

        private bool UpsertCore(SocialCore incoming)
        {
            incoming = NormalizeCore(incoming);
            if (!_cores.TryGetValue(incoming.Name, out SocialCore? existing))
            {
                _cores[incoming.Name] = incoming;
                return true;
            }

            bool changed = false;
            if (incoming.FirstSeen.HasValue && (!existing.FirstSeen.HasValue || incoming.FirstSeen.Value < existing.FirstSeen.Value))
            {
                existing.FirstSeen = incoming.FirstSeen;
                changed = true;
            }
            if (incoming.LastSeen.HasValue && (!existing.LastSeen.HasValue || incoming.LastSeen.Value > existing.LastSeen.Value))
            {
                existing.LastSeen = incoming.LastSeen;
                changed = true;
            }

            foreach (var user in incoming.Users)
            {
                if (!existing.Users.Any(u => u.Equals(user, StringComparison.OrdinalIgnoreCase)))
                {
                    existing.Users.Add(user);
                    changed = true;
                }
            }

            if (string.IsNullOrWhiteSpace(existing.Alias) && !string.IsNullOrWhiteSpace(incoming.Alias))
            {
                existing.Alias = incoming.Alias;
                changed = true;
            }

            return changed;
        }

        private static SocialCore NormalizeCore(SocialCore core)
        {
            core.Users ??= new List<string>();
            core.Alias = string.IsNullOrWhiteSpace(core.Alias) ? null : core.Alias.Trim();
            return core;
        }

        public List<string> Names()
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in _cores.Keys)
                set.Add(name);
            foreach (var overlayMap in _overlaysByProfile.Values)
            {
                foreach (var name in overlayMap.Keys)
                    set.Add(name);
            }
            return set.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        public List<string> RelatedByUsers(SocialProfile profile)
        {
            if (profile.Users.Count == 0)
                return new List<string>();

            return _cores.Values
                .Where(p => !p.Name.Equals(profile.Name, StringComparison.OrdinalIgnoreCase)
                            && p.Users.Any(u => profile.Users.Any(u2 => u2.Equals(u, StringComparison.OrdinalIgnoreCase))))
                .Select(p => p.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public SocialProfile? Get(string? name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;
            _cores.TryGetValue(name, out var core);
            var overlay = CurrentOverlays.TryGetValue(name, out var ov) ? ov : null;
            if (core == null && overlay == null)
                return null;

            return Merge(core, overlay);
        }

        public SocialProfile GetOrCreate(string name)
        {
            SocialCore core = EnsureCore(name);
            SocialOverlay overlay = GetOrCreateOverlay(name);
            return Merge(core, overlay);
        }

        public SocialOverlay GetOrCreateOverlay(string name)
        {
            var map = CurrentOverlays;
            if (!map.TryGetValue(name, out var overlay))
            {
                overlay = new SocialOverlay { Name = name };
                map[name] = overlay;
            }
            return overlay;
        }

        public void SetAlias(string name, string? alias)
        {
            if (!IsValidName(name))
                return;

            SocialCore core = EnsureCore(name);
            string? normalized = string.IsNullOrWhiteSpace(alias) ? null : alias.Trim();
            if (!string.Equals(core.Alias, normalized, StringComparison.Ordinal))
            {
                core.Alias = normalized;
            }
        }

        public bool Touch(string name, DateTime? timestamp, string? user = null)
        {
            if (!IsValidName(name))
                return false;

            bool changed = false;
            bool isNew = !_cores.ContainsKey(name);
            SocialCore core = EnsureCore(name);
            changed |= isNew;

            DateTime ts = timestamp ?? DateTime.Now;
            if (!core.FirstSeen.HasValue || ts < core.FirstSeen.Value)
            {
                core.FirstSeen = ts;
                changed = true;
            }
            if (!core.LastSeen.HasValue || ts > core.LastSeen.Value)
            {
                core.LastSeen = ts;
                changed = true;
            }

            if (!string.IsNullOrWhiteSpace(user))
            {
                if (!core.Users.Any(u => u.Equals(user, StringComparison.OrdinalIgnoreCase)))
                {
                    core.Users.Add(user);
                    changed = true;
                }
            }
            return changed;
        }

        private SocialCore EnsureCore(string name)
        {
            if (!_cores.TryGetValue(name, out SocialCore? core))
            {
                core = new SocialCore { Name = name };
                _cores[name] = core;
            }

            return core;
        }

        private static bool IsValidName(string name)
        {
            return !string.IsNullOrWhiteSpace(name) && NameFilter.IsMatch(name.Trim());
        }

        private Dictionary<string, SocialOverlay> EnsureOverlayMap(string key)
        {
            if (!_overlaysByProfile.TryGetValue(key, out var map))
            {
                map = new Dictionary<string, SocialOverlay>(StringComparer.OrdinalIgnoreCase);
                _overlaysByProfile[key] = map;
            }
            return map;
        }

        private static Dictionary<string, SocialOverlay> NormalizeOverlayMap(Dictionary<string, SocialOverlay> input)
        {
            var map = new Dictionary<string, SocialOverlay>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in input)
            {
                if (IsValidName(kvp.Key))
                {
                    SocialOverlay overlay = kvp.Value ?? new SocialOverlay { Name = kvp.Key };
                    overlay.Name ??= kvp.Key;
                    map[kvp.Key] = overlay;
                }
            }
            return map;
        }

        private SocialProfile Merge(SocialCore? core, SocialOverlay? overlay)
        {
            SocialProfile result = new SocialProfile();
            result.Name = core?.Name ?? overlay?.Name ?? string.Empty;
            result.FirstSeen = core?.FirstSeen;
            result.LastSeen = core?.LastSeen;
            result.Users = core?.Users ?? new List<string>();
            result.HaveMet = overlay?.HaveMet ?? false;
            result.KnowName = overlay?.KnowName ?? false;
            result.Alias = core?.Alias;
            result.Notes = overlay?.Notes;
            return result;
        }
    }

    public class SocialStoreData
    {
        public Dictionary<string, SocialCore> Cores { get; set; } = new Dictionary<string, SocialCore>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Dictionary<string, SocialOverlay>> Overlays { get; set; } = new Dictionary<string, Dictionary<string, SocialOverlay>>(StringComparer.OrdinalIgnoreCase);
        public string? ActiveProfile { get; set; }
    }
}
