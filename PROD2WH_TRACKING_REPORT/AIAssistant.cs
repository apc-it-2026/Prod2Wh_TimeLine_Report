using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROD2WH_TRACKING_REPORT
{
    // ======================== AI PARSER CLASS ========================
    public class AIQueryParser
    {
        public AIQueryResult ParseUserQuery(string userQuery)
        {
            var result = new AIQueryResult
            {
                OriginalQuery = userQuery,
                Success = true,
                Parameters = new SearchParameters()
            };

            try
            {
                userQuery = userQuery.ToLower().Trim();

                // ===== Check for standalone plant code first (e.g., just "ap1") =====
                if (IsStandalonePlantCode(userQuery))
                {
                    result.Intent = "SearchByPlant";
                    ExtractPlantParameters(userQuery, result.Parameters);
                    result.Parameters.vCheckCRD = true;
                    var (startDate, endDate) = GetCurrentWeekRange();
                    result.Parameters.vBeginDate = startDate;
                    result.Parameters.vEndDate = endDate;
                    result.NaturalResponse = $"🏭 Showing {result.Parameters.plant} data for this week (CRD: {startDate:yyyy/MM/dd} to {endDate:yyyy/MM/dd})";
                    result.SuggestedQuestions = GetSuggestionsForPlant();
                }
                else if (IsPlantFilter(userQuery))
                {
                    result.Intent = "SearchByPlant";
                    ExtractPlantParameters(userQuery, result.Parameters);
                    if (!userQuery.Contains("crd") && !userQuery.Contains("week") && !userQuery.Contains("month") && !userQuery.Contains("day"))
                    {
                        result.Parameters.vCheckCRD = true;
                        var (startDate, endDate) = GetCurrentWeekRange();
                        result.Parameters.vBeginDate = startDate;
                        result.Parameters.vEndDate = endDate;
                        result.NaturalResponse = $"🏭 Filtering by plant: {result.Parameters.plant} for this week";
                    }
                    else
                    {
                        result.NaturalResponse = $"🏭 Filtering by plant: {result.Parameters.plant}";
                    }
                    result.SuggestedQuestions = GetSuggestionsForPlant();
                }
                else if (IsSOSearch(userQuery))
                {
                    result.Intent = "SearchBySO";
                    ExtractSOParameters(userQuery, result.Parameters);
                    result.NaturalResponse = $"🔍 Searching for Sales Order: {result.Parameters.vSeId}";
                    result.SuggestedQuestions = GetSuggestionsForSO();
                }
                else if (IsCRDSearch(userQuery))
                {
                    result.Intent = "SearchByCRD";
                    ExtractCRDParameters(userQuery, result.Parameters);
                    result.NaturalResponse = $"📅 Searching CRD from {result.Parameters.vBeginDate:yyyy/MM/dd} to {result.Parameters.vEndDate:yyyy/MM/dd}";
                    result.SuggestedQuestions = GetSuggestionsForCRD();
                }
                else if (IsBulkSearch(userQuery))
                {
                    result.Intent = "SearchByBulk";
                    result.Parameters.SeIdList = "BULK_MODE";
                    result.NaturalResponse = "📋 Bulk search mode activated. Please paste your SO list in the bulk text area.";
                    result.SuggestedQuestions = GetSuggestionsForBulk();
                }
                else if (IsStatusFilter(userQuery))
                {
                    result.Intent = "SearchByStatus";
                    ExtractStatusParameters(userQuery, result.Parameters);
                    result.NaturalResponse = $"🚚 Filtering by shipping status: {result.Parameters.vShipStatus}";
                    result.SuggestedQuestions = GetSuggestionsForStatus();
                }
                else if (IsSummaryRequest(userQuery))
                {
                    result.Intent = "ShowSummary";
                    result.SuggestedTab = 0;
                    result.NaturalResponse = "📊 Loading overall summary dashboard...";
                    result.SuggestedQuestions = GetSuggestionsForSummary();
                }
                else if (IsTrendRequest(userQuery))
                {
                    result.Intent = "ShowTrend";
                    result.SuggestedTab = 2;
                    result.NaturalResponse = "📈 Generating trend charts and analysis...";
                    result.SuggestedQuestions = GetSuggestionsForTrend();
                }
                else if (IsDelayRequest(userQuery))
                {
                    result.Intent = "ShowDelay";
                    result.SuggestedTab = 4;
                    result.NaturalResponse = "⏰ Preparing process delay report...";
                    result.SuggestedQuestions = GetSuggestionsForDelay();
                }
                else if (IsExportRequest(userQuery))
                {
                    result.Intent = "ExportData";
                    result.ShouldExport = true;
                    result.NaturalResponse = "💾 Preparing data for export...";
                    result.SuggestedQuestions = GetSuggestionsForExport();
                }
                else if (IsSOPRequest(userQuery))
                {
                    result.Intent = "ShowSOP";
                    result.SuggestedTab = 1;
                    result.NaturalResponse = "📖 Loading SOP documentation...";
                    result.SuggestedQuestions = GetSuggestionsForSOP();
                }
                else if (IsClearRequest(userQuery))
                {
                    result.Intent = "ClearFilters";
                    result.ShouldClear = true;
                    result.NaturalResponse = "🧹 Clearing all filters...";
                    result.SuggestedQuestions = GetDefaultSuggestions();
                }
                else
                {
                    result.Intent = "ShowSummary";
                    result.SuggestedTab = 0;
                    result.NaturalResponse = "📊 Loading overall summary. Try: '1000303095', 'AP1', 'CRD this week', or 'delay report'";
                    result.SuggestedQuestions = GetDefaultSuggestions();
                }

                if (result.Parameters.vCheckCRD)
                {
                    var monthDiff = ((result.Parameters.vEndDate.Year - result.Parameters.vBeginDate.Year) * 12) + result.Parameters.vEndDate.Month - result.Parameters.vBeginDate.Month;
                    if (monthDiff > 3)
                    {
                        result.Success = false;
                        result.NaturalResponse = "⚠️ CRD date range must not exceed 3 months.";
                    }
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.NaturalResponse = "❌ I couldn't understand. Try: '1000303095', 'AP1', 'CRD this week', or 'delay report'";
            }

            return result;
        }

        #region Intent Detection Methods

        private bool IsStandalonePlantCode(string q)
        {
            string trimmedQuery = q.Trim();
            return Regex.IsMatch(trimmedQuery, @"^(AP\d{1,2}|API|MK1)$", RegexOptions.IgnoreCase);
        }

        private bool IsSOSearch(string q) => Regex.IsMatch(q, @"\b(so|sales order|order|show|find|search).{0,10}\d{3,}\b") || Regex.IsMatch(q, @"\b\d{3,}\b");
        private bool IsCRDSearch(string q) => q.Contains("crd") || q.Contains("due date") || q.Contains("delivery date") || q.Contains("commit date");
        private bool IsBulkSearch(string q) => q.Contains("bulk") || q.Contains("multiple") || q.Contains("list") || q.Contains("several");
        private bool IsStatusFilter(string q) => new[] { "status", "shipping status", "shipped", "pending", "delayed", "transit", "cancelled" }.Any(k => q.Contains(k));
        private bool IsSummaryRequest(string q) => q.Contains("summary") || q.Contains("overall") || q.Contains("total") || q.Contains("dashboard");
        private bool IsTrendRequest(string q) => q.Contains("trend") || q.Contains("chart") || q.Contains("graph") || q.Contains("analysis");
        private bool IsDelayRequest(string q) => q.Contains("delay") || q.Contains("delayed") || q.Contains("late") || q.Contains("overdue");
        private bool IsExportRequest(string q) => q.Contains("export") || q.Contains("download") || q.Contains("excel");
        private bool IsSOPRequest(string q) => q.Contains("sop") || q.Contains("procedure") || q.Contains("guideline");
        private bool IsClearRequest(string q) => q.Contains("clear") || q.Contains("reset") || q.Contains("remove all");

        private bool IsPlantFilter(string q)
        {
            bool hasPlantKeyword = q.Contains("plant") || q.Contains("factory") || q.Contains("site") || q.Contains("filter");
            bool hasPlantCode = Regex.IsMatch(q, @"\b(AP\d{1,2}|API|MK1)\b", RegexOptions.IgnoreCase);
            return hasPlantKeyword && hasPlantCode;
        }

        #endregion

        #region Parameter Extraction Methods

        private void ExtractSOParameters(string q, SearchParameters p)
        {
            var match = Regex.Match(q, @"\b(?:so[-_]?)?(\d{3,})\b", RegexOptions.IgnoreCase);
            if (match.Success) p.vSeId = match.Groups[1].Value;
            p.vCheckCRD = false;
        }

        private void ExtractCRDParameters(string q, SearchParameters p)
        {
            p.vCheckCRD = true;
            if (q.Contains("this week")) { var (s, e) = GetCurrentWeekRange(); p.vBeginDate = s; p.vEndDate = e; }
            else if (q.Contains("this month")) { var (s, e) = GetCurrentMonthRange(); p.vBeginDate = s; p.vEndDate = e; }
            else if (q.Contains("last 7 days") || q.Contains("past week")) { p.vBeginDate = DateTime.Now.AddDays(-7); p.vEndDate = DateTime.Now; }
            else if (q.Contains("last 30 days") || q.Contains("past month")) { p.vBeginDate = DateTime.Now.AddDays(-30); p.vEndDate = DateTime.Now; }
            else if (q.Contains("today")) { p.vBeginDate = DateTime.Now; p.vEndDate = DateTime.Now; }
            else
            {
                var dates = Regex.Matches(q, @"(\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}[/-]\d{4})");
                if (dates.Count >= 2) { p.vBeginDate = DateTime.Parse(dates[0].Value); p.vEndDate = DateTime.Parse(dates[1].Value); }
                else if (dates.Count == 1) { p.vBeginDate = DateTime.Parse(dates[0].Value); p.vEndDate = p.vBeginDate; }
                else { p.vBeginDate = DateTime.Now.AddDays(-30); p.vEndDate = DateTime.Now; }
            }
        }

        private void ExtractStatusParameters(string q, SearchParameters p)
        {
            var statuses = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"shipped","Shipped"},{"on time","Shipped"},{"pending","Pending"},{"delayed","Delayed"},
                {"late","Delayed"},{"in transit","In Transit"},{"transit","In Transit"},{"cancelled","Cancelled"}
            };
            foreach (var s in statuses) if (q.Contains(s.Key)) { p.vShipStatus = s.Value; break; }
            p.vCheckCRD = false;
        }

        private void ExtractPlantParameters(string q, SearchParameters p)
        {
            p.plant = "";
            Match match = Regex.Match(q, @"\b(AP\d{1,2}|API|MK1)\b", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                p.plant = match.Value.ToUpper();
            }
            p.vCheckCRD = false;
        }

        #endregion

        #region Date Helper Methods

        private (DateTime, DateTime) GetCurrentWeekRange()
        {
            var today = DateTime.Now;
            int diff = (7 - (int)today.DayOfWeek + (int)DayOfWeek.Monday) % 7;
            var end = today.AddDays(diff);
            return (end.AddDays(-6), end);
        }

        private (DateTime, DateTime) GetCurrentMonthRange()
        {
            var start = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            return (start, start.AddMonths(1).AddDays(-1));
        }

        #endregion

        #region Suggestion Methods

        private List<string> GetDefaultSuggestions() => new List<string> { "1000303095", "AP1", "CRD this week", "Show delay report", "Shipped" };
        private List<string> GetSuggestionsForSO() => new List<string> { "CRD this week", "Show delay report", "Export to Excel", "AP1" };
        private List<string> GetSuggestionsForCRD() => new List<string> { "Show overall summary", "Shipped", "Show trend chart", "Export to Excel" };
        private List<string> GetSuggestionsForBulk() => new List<string> { "Show summary", "CRD last 30 days", "Show delay report" };
        private List<string> GetSuggestionsForStatus() => new List<string> { "Show delay report", "AP1", "Show overall summary", "CRD this month" };
        private List<string> GetSuggestionsForPlant() => new List<string> { "AP1", "AP2", "MK1", "API", "CRD this week" };
        private List<string> GetSuggestionsForSummary() => new List<string> { "Show trend chart", "Show delay report", "CRD last 30 days", "Export to Excel" };
        private List<string> GetSuggestionsForTrend() => new List<string> { "Show overall summary", "Show delay report", "Export to Excel" };
        private List<string> GetSuggestionsForDelay() => new List<string> { "Show overall summary", "Show trend chart", "AP1", "Export to Excel" };
        private List<string> GetSuggestionsForExport() => new List<string> { "Show overall summary", "CRD this week", "Show delay report" };
        private List<string> GetSuggestionsForSOP() => new List<string> { "Show overall summary", "CRD this month", "Show delay report" };

        #endregion
    }

    public class AIQueryResult
    {
        public bool Success { get; set; } = true;
        public string Intent { get; set; } = "";
        public string OriginalQuery { get; set; } = "";
        public string NaturalResponse { get; set; } = "";
        public SearchParameters Parameters { get; set; } = new SearchParameters();
        public int SuggestedTab { get; set; } = -1;
        public bool ShouldExport { get; set; } = false;
        public bool ShouldClear { get; set; } = false;
        public List<string> SuggestedQuestions { get; set; } = new List<string>();
    }

    public class SearchParameters
    {
        public string vSeId { get; set; } = "";
        public string SeIdList { get; set; } = "";
        public bool vCheckCRD { get; set; } = false;
        public DateTime vBeginDate { get; set; } = DateTime.Now.AddDays(-30);
        public DateTime vEndDate { get; set; } = DateTime.Now;
        public string vShipStatus { get; set; } = "";
        public string plant { get; set; } = "";
    }

    // ======================== AI ASSISTANT FLOATING PANEL ========================
    public class AIAssistantFloatPanel : Form
    {
        private readonly AIQueryParser _parser;
        private RichTextBox txtChatHistory;
        private TextBox txtQuery;
        private FlowLayoutPanel suggestionPanel;
        private Button btnSend;
        private PROD2WH_Tracking_List _mainForm;
        private bool _isExpanded = true;

        public AIAssistantFloatPanel(PROD2WH_Tracking_List mainForm)
        {
            _mainForm = mainForm;
            _parser = new AIQueryParser();
            SetupForm();
            CreateUI();
            LoadInitialMessages();
        }

        private void SetupForm()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.White;
            this.Size = new Size(380, 520);
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
        }

        private void CreateUI()
        {
            // Title Bar - Teal color
            var titleBar = new Panel { Height = 40, Dock = DockStyle.Top, BackColor = Color.Teal };
            var lblTitle = new Label { Text = "🤖 AI Assistant", ForeColor = Color.White, Location = new Point(12, 10), AutoSize = true, Font = new Font("Times New Roman", 12, FontStyle.Bold) };
            var btnMinimize = new Button { Text = "−", FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.Transparent, Location = new Point(this.Width - 65, 8), Size = new Size(25, 25), Font = new Font("Times New Roman", 12, FontStyle.Bold) };
            btnMinimize.Click += (s, e) => ToggleExpand();
            var btnClose = new Button { Text = "✕", FlatStyle = FlatStyle.Flat, ForeColor = Color.White, BackColor = Color.Transparent, Location = new Point(this.Width - 35, 8), Size = new Size(25, 25), Font = new Font("Times New Roman", 10) };
            btnClose.Click += (s, e) => this.Hide();
            titleBar.Controls.AddRange(new Control[] { lblTitle, btnMinimize, btnClose });

            // Chat History
            txtChatHistory = new RichTextBox
            {
                Location = new Point(10, 50),
                Size = new Size(360, 340),
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Times New Roman", 10),
                BackColor = Color.FromArgb(248, 248, 248)
            };

            // Suggestions Panel
            suggestionPanel = new FlowLayoutPanel { Location = new Point(10, 400), Size = new Size(360, 45), AutoScroll = true, WrapContents = true };

            // Input Area
            var inputPanel = new Panel { Location = new Point(10, 450), Size = new Size(360, 55) };
            txtQuery = new TextBox
            {
                Width = 280,
                Height = 35,
                Location = new Point(0, 0),
                Font = new Font("Times New Roman", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtQuery.Text = "Ask me anything...";
            txtQuery.ForeColor = Color.Gray;
            txtQuery.Enter += (s, e) => { if (txtQuery.Text == "Ask me anything...") { txtQuery.Text = ""; txtQuery.ForeColor = Color.Black; } };
            txtQuery.Leave += (s, e) => { if (string.IsNullOrWhiteSpace(txtQuery.Text)) { txtQuery.Text = "Ask me anything..."; txtQuery.ForeColor = Color.Gray; } };
            txtQuery.KeyPress += async (s, e) => { if (e.KeyChar == (char)Keys.Enter) await ProcessQuery(txtQuery.Text); };

            btnSend = new Button
            {
                Text = "Send",
                Location = new Point(285, 0),
                Size = new Size(70, 35),
                BackColor = Color.Teal,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Times New Roman", 10, FontStyle.Bold)
            };
            btnSend.Click += async (s, e) => await ProcessQuery(txtQuery.Text);
            inputPanel.Controls.AddRange(new Control[] { txtQuery, btnSend });

            this.Controls.AddRange(new Control[] { titleBar, txtChatHistory, suggestionPanel, inputPanel });

            // Make draggable
            bool dragging = false; Point dragPoint = Point.Empty;
            titleBar.MouseDown += (s, e) => { dragging = true; dragPoint = new Point(e.X, e.Y); };
            titleBar.MouseMove += (s, e) => { if (dragging) this.Location = new Point(this.Location.X + e.X - dragPoint.X, this.Location.Y + e.Y - dragPoint.Y); };
            titleBar.MouseUp += (s, e) => { dragging = false; };
        }

        private void ToggleExpand()
        {
            _isExpanded = !_isExpanded;
            if (_isExpanded) { this.Size = new Size(380, 520); txtChatHistory.Visible = true; suggestionPanel.Visible = true; }
            else { this.Size = new Size(380, 40); txtChatHistory.Visible = false; suggestionPanel.Visible = false; }
        }

        private async Task ProcessQuery(string query)
        {
            if (string.IsNullOrWhiteSpace(query) || query == "Ask me anything...") return;

            AppendMessage($"You: {query}", Color.FromArgb(0, 102, 204));
            AppendMessage("AI Assistant: thinking...", Color.Gray);
            await Task.Delay(50);

            var result = _parser.ParseUserQuery(query);
            RemoveLastMessage();

            if (result.Success)
            {
                _mainForm.ApplyAIParameters(result.Parameters);

                if (result.ShouldClear)
                    _mainForm.ClearAllFilters();

                if (HasValidParameters(result.Parameters) && !result.ShouldClear)
                    _mainForm.ExecuteSearchFromAI();

                if (result.SuggestedTab >= 0)
                    _mainForm.SwitchToTab(result.SuggestedTab);

                if (result.ShouldExport)
                    _mainForm.TriggerExport();

                AppendMessage($"🤖 {result.NaturalResponse}", Color.FromArgb(0, 150, 0));

                if (result.SuggestedQuestions?.Any() == true)
                {
                    AppendMessage("", Color.Black);
                    AppendMessage("💡 Try these:", Color.FromArgb(255, 140, 0));
                    foreach (var s in result.SuggestedQuestions.Take(4))
                        AppendMessage($"   • {s}", Color.FromArgb(255, 140, 0));
                }
                UpdateSuggestionChips(result.SuggestedQuestions);
            }
            else
            {
                AppendMessage($"🤖 {result.NaturalResponse}", Color.Red);
            }
            txtQuery.Text = "Ask me anything...";
            txtQuery.ForeColor = Color.Gray;
        }

        private bool HasValidParameters(SearchParameters p) => !string.IsNullOrEmpty(p.vSeId) || !string.IsNullOrEmpty(p.SeIdList) || p.vCheckCRD;

        private void AppendMessage(string msg, Color clr)
        {
            if (txtChatHistory.InvokeRequired) { txtChatHistory.Invoke(new Action(() => AppendMessage(msg, clr))); return; }
            if (!string.IsNullOrEmpty(msg))
            {
                txtChatHistory.SelectionStart = txtChatHistory.TextLength;
                txtChatHistory.SelectionColor = clr;
                txtChatHistory.AppendText(msg + Environment.NewLine);
                txtChatHistory.ScrollToCaret();
            }
            else txtChatHistory.AppendText(Environment.NewLine);
        }

        private void RemoveLastMessage()
        {
            if (txtChatHistory.InvokeRequired) { txtChatHistory.Invoke(new Action(RemoveLastMessage)); return; }
            var lines = txtChatHistory.Lines;
            if (lines.Length > 0)
            {
                var idx = txtChatHistory.GetFirstCharIndexFromLine(lines.Length - 1);
                if (idx >= 0) txtChatHistory.Text = txtChatHistory.Text.Substring(0, idx);
            }
        }

        private void UpdateSuggestionChips(List<string> suggestions)
        {
            if (suggestionPanel.InvokeRequired)
            {
                suggestionPanel.Invoke(new Action(() => UpdateSuggestionChips(suggestions)));
                return;
            }

            suggestionPanel.Controls.Clear();
            if (suggestions == null) return;

            foreach (var suggestion in suggestions.Take(4))
            {
                var chip = new Button
                {
                    Text = suggestion.Length > 25 ? suggestion.Substring(0, 22) + "..." : suggestion,
                    BackColor = Color.FromArgb(240, 240, 240),
                    FlatStyle = FlatStyle.Flat,
                    Margin = new Padding(3),
                    Padding = new Padding(10, 5, 10, 5),
                    AutoSize = true,
                    Font = new Font("Times New Roman", 9),
                    Cursor = Cursors.Hand
                };
                chip.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);

                string suggestionText = suggestion;
                chip.Click += async (sender, e) => await ProcessQuery(suggestionText);

                suggestionPanel.Controls.Add(chip);
            }
        }

        private async void LoadInitialMessages()
        {
            await Task.Delay(300);
            AppendMessage("🤖 AI Assistant: Hello! I'm your PROD2WH tracking assistant.", Color.Teal);
            AppendMessage("", Color.Black);
            AppendMessage("I can help you with:", Color.Teal);
            AppendMessage("   • Search by Sales Order (e.g., '1000303095')", Color.Black);
            AppendMessage("   • Filter by Plant (e.g., 'AP1' or 'Plant AP1')", Color.Black);
            AppendMessage("   • Filter by CRD date (e.g., 'CRD this week')", Color.Black);
            AppendMessage("   • View reports (e.g., 'Show delay report')", Color.Black);
            AppendMessage("   • Filter by status (e.g., 'Shipped')", Color.Black);
            AppendMessage("   • Export data (e.g., 'Export to Excel')", Color.Black);
            AppendMessage("", Color.Black);
            AppendMessage("Try one of these:", Color.FromArgb(255, 140, 0));

            var suggestions = new List<string> { "1000303095", "AP1", "CRD this week", "Show delay report", "Shipped" };
            foreach (var s in suggestions) AppendMessage($"   • {s}", Color.FromArgb(255, 140, 0));
            UpdateSuggestionChips(suggestions);
        }
    }

    // ======================== EXTENSION METHODS FOR MAIN FORM ========================

    public static class FormExtensions
    {
        public static void ApplyAIParameters(this PROD2WH_Tracking_List form, SearchParameters parameters)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.ApplyAIParameters(parameters))); return; }

            var textBox_SeId = form.Controls.Find("textBox_SeId", true).FirstOrDefault() as TextBox;
            var richTextBox1 = form.Controls.Find("richTextBox1", true).FirstOrDefault() as RichTextBox;
            var checkBox_CRD = form.Controls.Find("checkBox_CRD", true).FirstOrDefault() as CheckBox;
            var dateTimePicker1 = form.Controls.Find("dateTimePicker1", true).FirstOrDefault() as DateTimePicker;
            var dateTimePicker2 = form.Controls.Find("dateTimePicker2", true).FirstOrDefault() as DateTimePicker;
            var Shipstatuscombo = form.Controls.Find("Shipstatuscombo", true).FirstOrDefault() as ComboBox;
            var plantcombo = form.Controls.Find("plantcombo", true).FirstOrDefault() as ComboBox;

            if (textBox_SeId != null) textBox_SeId.Text = parameters.vSeId;
            if (richTextBox1 != null) richTextBox1.Text = parameters.SeIdList;
            if (checkBox_CRD != null) checkBox_CRD.Checked = parameters.vCheckCRD;
            if (dateTimePicker1 != null) dateTimePicker1.Value = parameters.vBeginDate;
            if (dateTimePicker2 != null) dateTimePicker2.Value = parameters.vEndDate;

            if (!string.IsNullOrEmpty(parameters.vShipStatus) && Shipstatuscombo != null)
            {
                if (Shipstatuscombo.Items.Contains(parameters.vShipStatus))
                    Shipstatuscombo.SelectedItem = parameters.vShipStatus;
                else
                    Shipstatuscombo.Text = parameters.vShipStatus;
            }

            if (!string.IsNullOrEmpty(parameters.plant) && plantcombo != null)
            {
                bool itemExists = false;
                foreach (var item in plantcombo.Items)
                {
                    if (item.ToString().Equals(parameters.plant, StringComparison.OrdinalIgnoreCase))
                    {
                        plantcombo.SelectedItem = item;
                        itemExists = true;
                        break;
                    }
                }
                if (!itemExists)
                {
                    plantcombo.Text = parameters.plant;
                }
            }
        }
        private static Button FindButtonRecursive(Control parent, string buttonName)
        {
            foreach (Control child in parent.Controls)
            {
                if (child is Button btn && btn.Name == buttonName)
                    return btn;

                Button found = FindButtonRecursive(child, buttonName);
                if (found != null)
                    return found;
            }
            return null;
        }

        public static void ClearAllFilters(this PROD2WH_Tracking_List form)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.ClearAllFilters())); return; }

            var textBox_SeId = form.Controls.Find("textBox_SeId", true).FirstOrDefault() as TextBox;
            var richTextBox1 = form.Controls.Find("richTextBox1", true).FirstOrDefault() as RichTextBox;
            var checkBox_CRD = form.Controls.Find("checkBox_CRD", true).FirstOrDefault() as CheckBox;
            var Shipstatuscombo = form.Controls.Find("Shipstatuscombo", true).FirstOrDefault() as ComboBox;
            var plantcombo = form.Controls.Find("plantcombo", true).FirstOrDefault() as ComboBox;
            var dateTimePicker1 = form.Controls.Find("dateTimePicker1", true).FirstOrDefault() as DateTimePicker;
            var dateTimePicker2 = form.Controls.Find("dateTimePicker2", true).FirstOrDefault() as DateTimePicker;

            if (textBox_SeId != null) textBox_SeId.Text = "";
            if (richTextBox1 != null) richTextBox1.Text = "";
            if (checkBox_CRD != null) checkBox_CRD.Checked = false;
            if (Shipstatuscombo != null) { Shipstatuscombo.SelectedIndex = -1; Shipstatuscombo.Text = ""; }
            if (plantcombo != null) { plantcombo.SelectedIndex = -1; plantcombo.Text = ""; }
            if (dateTimePicker1 != null) dateTimePicker1.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            if (dateTimePicker2 != null) dateTimePicker2.Value = DateTime.Now;
        }

        public static void ExecuteSearchFromAI(this PROD2WH_Tracking_List form)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.ExecuteSearchFromAI())); return; }
            var btnSelect = form.Controls.Find("btnSelect", true).FirstOrDefault() as Button;
            if (btnSelect != null)
            {
                btnSelect.PerformClick();
            }
        }

        public static void SwitchToTab(this PROD2WH_Tracking_List form, int tabIndex)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.SwitchToTab(tabIndex))); return; }
            var tabControl = form.Controls.Find("tabControl1", true).FirstOrDefault() as TabControl;
            if (tabControl != null && tabIndex >= 0 && tabIndex < tabControl.TabCount)
                tabControl.SelectedIndex = tabIndex;
        }

        public static void TriggerExport(this PROD2WH_Tracking_List form)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.TriggerExport())); return; }
            var btnExport = form.Controls.Find("Export", true).FirstOrDefault() as Button;
            if (btnExport != null)
                btnExport.PerformClick();
        }

        public static void InitializeAIAssistant(this PROD2WH_Tracking_List form)
        {
            if (form.InvokeRequired) { form.Invoke(new Action(() => form.InitializeAIAssistant())); return; }

            var assistant = new AIAssistantFloatPanel(form);
            assistant.Show();
            assistant.Location = new Point(form.Location.X + form.Width - assistant.Width - 10, form.Location.Y + form.Height - assistant.Height - 10);

            form.Move += (s, e) => { if (assistant.Visible) assistant.Location = new Point(form.Location.X + form.Width - assistant.Width - 10, form.Location.Y + form.Height - assistant.Height - 10); };

            // ===== FIND YOUR EXISTING button2 =====
            Button existingButton2 = FindButtonRecursive(form, "button2");

            // ===== AI TOGGLE BUTTON - Place it BESIDE (to the right of) button2 =====
            var btnToggle = new Button
            {
                Text = "🤖 AI",
                Size = new Size(50, 32),
                BackColor = Color.Navy,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Times New Roman", 20, FontStyle.Bold)
            };

            if (existingButton2 != null)
            {
                // Position it to the RIGHT of button2
                btnToggle.Location = new Point(existingButton2.Right + 10, existingButton2.Top);

                // Add to the same parent as button2
                if (existingButton2.Parent != form)
                {
                    existingButton2.Parent.Controls.Add(btnToggle);
                }
                else
                {
                    form.Controls.Add(btnToggle);
                }
            }
            else
            {
                // Fallback: Add to form at top right corner
                btnToggle.Location = new Point(form.ClientSize.Width - btnToggle.Width - 10, 5);
                btnToggle.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                form.Controls.Add(btnToggle);
            }

            btnToggle.BringToFront();
            btnToggle.Click += (s, e) => { assistant.Visible = !assistant.Visible; if (assistant.Visible) assistant.BringToFront(); };
        }

       
        
    }
}