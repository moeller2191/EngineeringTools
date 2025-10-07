using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace EngineeringTools.UI
{
    public partial class LoadingWindow : Window
    {
        public LoadingWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Update the loading progress and status
        /// </summary>
        public void UpdateProgress(int progressPercent, string status, string detailedStatus = "")
        {
            Dispatcher.Invoke(() =>
            {
                LoadingProgressBar.Value = progressPercent;
                LoadingStatusText.Text = status;
                
                if (!string.IsNullOrEmpty(detailedStatus))
                {
                    DetailedStatusText.Text = detailedStatus;
                }
            });
        }

        /// <summary>
        /// Update a specific loading step
        /// </summary>
        public void UpdateStep(int stepNumber, string status, bool completed = false)
        {
            Dispatcher.Invoke(() =>
            {
                var stepText = stepNumber switch
                {
                    1 => Step1Text,
                    2 => Step2Text,
                    3 => Step3Text,
                    4 => Step4Text,
                    5 => Step5Text,
                    _ => null
                };

                if (stepText != null)
                {
                    stepText.Text = completed ? $"✓ {status}" : $"• {status}";
                    stepText.Foreground = completed ? Brushes.Green : Brushes.Orange;
                }
            });
        }

        /// <summary>
        /// Show error state
        /// </summary>
        public void ShowError(string errorMessage)
        {
            Dispatcher.Invoke(() =>
            {
                LoadingStatusText.Text = "Error during initialization";
                LoadingStatusText.Foreground = Brushes.Red;
                DetailedStatusText.Text = errorMessage;
                DetailedStatusText.Foreground = Brushes.Red;
                LoadingProgressBar.Foreground = Brushes.Red;
            });
        }

        /// <summary>
        /// Show completion state
        /// </summary>
        public void ShowCompleted()
        {
            Dispatcher.Invoke(() =>
            {
                LoadingProgressBar.Value = 100;
                LoadingStatusText.Text = "Loading complete!";
                LoadingStatusText.Foreground = Brushes.Green;
                DetailedStatusText.Text = "Opening Engineering Tools...";
                
                // Update all steps as completed
                UpdateStep(1, "Database connection", true);
                UpdateStep(2, "Excel data loaded", true);
                UpdateStep(3, "Sales Order system ready", true);
                UpdateStep(4, "Programming Check system ready", true);
                UpdateStep(5, "User interface ready", true);
            });
        }

        /// <summary>
        /// Close the loading window with fade effect
        /// </summary>
        public async Task CloseWithDelay(int delayMs = 1000)
        {
            await Task.Delay(delayMs);
            Dispatcher.Invoke(() => Close());
        }
    }
}