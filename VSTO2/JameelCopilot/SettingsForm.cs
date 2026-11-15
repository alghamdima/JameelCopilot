using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using JameelCopilot.Services;

namespace JameelCopilot
{
    /// <summary>
    /// The About dialog: what Jameel Copilot is set to, and a connection test. It is
    /// read-only. Settings belong to an administrator and live in HKLM (see
    /// <see cref="MachineSettings"/>), so there is nothing here to edit and no
    /// password to guard it.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly CustomApiChatService _chat;
        private readonly CancellationTokenSource _closing = new CancellationTokenSource();
        private Button _btnTest;
        private Label _lblStatus;

        // The ribbon button opens it through here, so it always shows what the
        // registry holds now.
        public static void Open(CustomApiChatService chat)
        {
            Config.Load();
            // When a modal dialog ends, WinForms replaces the thread's context
            // with a plain one, and an await on the UI thread then resumes on the
            // thread pool: an open pane stops answering. Put ours back.
            var context = System.Threading.SynchronizationContext.Current;
            try
            {
                using (var form = new SettingsForm(chat))
                {
                    form.ShowDialog();
                }
            }
            finally
            {
                System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            }
        }

        public SettingsForm(CustomApiChatService chat)
        {
            _chat = chat;

            Text = "About Jameel Copilot";
            ClientSize = new Size(380, 262);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(250, 249, 248);
            Font = new Font("Segoe UI", 9F);

            AddMark();
            Controls.Add(new Label
            {
                Location = new Point(68, 14),
                Size = new Size(296, 152),
                Text = string.Join(Environment.NewLine, DescribeSettings()),
            });

            _lblStatus = new Label
            {
                Location = new Point(16, 172),
                Size = new Size(348, 34),
                ForeColor = Color.DarkSlateGray,
            };
            Controls.Add(_lblStatus);

            _btnTest = NewButton("Test connection", new Point(16, 218), 130);
            _btnTest.Click += BtnTest_Click;
            Controls.Add(_btnTest);

            var close = NewButton("Close", new Point(284, 218), 80);
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(close);
            CancelButton = close;
        }

        private void AddMark()
        {
            try
            {
                var image = LogoImages.Sized(80);
                Controls.Add(new PictureBox
                {
                    Location = new Point(16, 14),
                    Size = new Size(40, 40),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Image = image,
                });
                Disposed += (sender, e) => image.Dispose();
            }
            catch (Exception)
            {
                // The dialog is still useful without its picture.
            }
        }

        // As it is sent: with a dot, whatever the PC's regional settings.
        private static string Number(double value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }

        /// <summary>The lines the dialog shows. Never the key, and only the host of the address.</summary>
        internal static string[] DescribeSettings()
        {
            var tools = Config.EnabledWriteTools.OrderBy(name => name, StringComparer.Ordinal).ToArray();
            var lines = new List<string>
            {
                "Version: " + typeof(SettingsForm).Assembly.GetName().Version.ToString(3),
                "AI service: " + (Config.ApiBaseUrl != null ? Config.ApiBaseUrl.Host : "not set"),
                "Model: " + (Config.Model ?? "not set"),
                "Temperature: " + (Config.Temperature.HasValue
                    ? Number(Config.Temperature.Value) + " (creative up to " + Number(Config.TemperatureMax) + ")"
                    : "not sent"),
            };
            if (Config.MaxOutputTokens.HasValue)
            {
                lines.Add("Answer length limit: " + Config.MaxOutputTokens.Value.ToString(CultureInfo.InvariantCulture) + " tokens");
            }
            lines.Add("Write tools: " + (tools.Length == 0 ? "none" : string.Join(", ", tools)));
            if (!Config.IsConfigured)
            {
                lines.Add("");
                lines.Add("Not configured: " + Config.Problem);
            }
            return lines.ToArray();
        }

        // Explicit colors: Windows Server 2025 renders default-styled buttons blank.
        private static Button NewButton(string text, Point location, int width)
        {
            return new Button
            {
                Text = text,
                Location = location,
                Size = new Size(width, 28),
                ForeColor = Color.Black,
                BackColor = Color.FromArgb(225, 225, 225),
                UseVisualStyleBackColor = false,
            };
        }

        private async void BtnTest_Click(object sender, EventArgs e)
        {
            _btnTest.Enabled = false;
            _lblStatus.ForeColor = Color.DarkSlateGray;
            _lblStatus.Text = "Testing...";
            string problem;
            try
            {
                problem = _chat == null
                    ? "The chat service is unavailable. Restart Outlook."
                    : await _chat.TestConnectionAsync(_closing.Token);
            }
            catch (Exception ex)
            {
                problem = ApiErrors.ForException(ex);
            }
            if (IsDisposed || _closing.IsCancellationRequested) return;
            _lblStatus.ForeColor = problem == null ? Color.DarkGreen : Color.DarkRed;
            _lblStatus.Text = problem ?? "Connection OK.";
            _btnTest.Enabled = true;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _closing.Cancel();
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _closing.Dispose();
            base.Dispose(disposing);
        }
    }
}
