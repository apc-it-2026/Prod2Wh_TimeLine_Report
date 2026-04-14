using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using TExcel = Microsoft.Office.Interop.Excel;




namespace PROD2WH_TRACKING_REPORT
{
    public partial class PROD2WH_Tracking_List : Form
    {
        DataTable dtJson;
        // Global variables
        private int totalOnTime;
        private int totalDelayed;
        private int totalOrders;
        private int totalShippedLate;

        private Dictionary<string, double> avgDelays;

        private int messageIndex = 0;
        private List<string> messages;
        private Timer rollTimer;
        // Declare once at form level
        ToolTip customToolTip = new ToolTip();
        public PROD2WH_Tracking_List()
        {
            //InitializeComponent();
            //loadbl.Visible = false;
            //progressBar1.Visible = false;
            //ucRollText1.Text = string.Empty;

            //this.dateTimePicker1.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            //this.dateTimePicker2.Value = DateTime.Now.AddDays(0);

            InitializeComponent();
            loadbl.Visible = false;
            progressBar1.Visible = false;
            ucRollText1.Text = string.Empty;

            this.dateTimePicker1.Value = DateTime.Now.AddDays(1 - DateTime.Now.Day);
            this.dateTimePicker2.Value = DateTime.Now.AddDays(0);

            // Add this line ONLY
            this.Load += (s, e) => this.InitializeAIAssistant();


        }

        private async void btnSelect_Click(object sender, EventArgs e)
        {
           
            if (string.IsNullOrEmpty(textBox_SeId.Text) &&
                !checkBox_CRD.Checked &&
                string.IsNullOrEmpty(richTextBox1.Text))
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this,
                    "Please Select Any One Condition: SO Or CRD Or Bulk SO List !!");
                return;
            }

            if (!Validate_CRD_Date())
            {
                SJeMES_Control_Library.MessageHelper.ShowErr(this,
                    "CRD Range Must Not Exceed 3 Months!");
                return;
            }

    
            if (tabControl1.SelectedIndex != 0 &&
                tabControl1.SelectedIndex != 1 &&
                tabControl1.SelectedIndex != 2)
                return;

            try
            {
              
                dataGridView2.DataSource = null;
                loadbl.Visible = true;
                loaderPictureBox.Image = Image.FromFile("Images/SandyLoading.gif");
                loaderPictureBox.SizeMode = PictureBoxSizeMode.CenterImage;
                loaderPictureBox.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.Visible = true;
              

            
                Dictionary<string, object> p = new Dictionary<string, object>()
        {
            { "vSeId", textBox_SeId.Text },
            { "SeIdList", richTextBox1.Text.Trim() },
            { "vCheckCRD", checkBox_CRD.Checked },
            { "vBeginDate", dateTimePicker1.Value.ToShortDateString() },
            { "vEndDate", dateTimePicker2.Value.ToShortDateString() },
        };

                string shipStatus = Shipstatuscombo.SelectedItem?.ToString();
                p.Add("ShipStatus",
                    string.IsNullOrWhiteSpace(shipStatus) || shipStatus.Equals("All", StringComparison.OrdinalIgnoreCase)
                        ? null
                        : shipStatus
                );
                string plantValue = string.Empty;

                bool isSelected = plantcombo.SelectedItem != null;
                bool isTyped = !string.IsNullOrWhiteSpace(plantcombo.Text);


                if (isSelected)
                {
                    plantValue = plantcombo.SelectedItem.ToString();
                }
                else
                {
                    plantValue = plantcombo.Text.Trim().ToUpper();
                }

                p.Add("plant", plantValue);

                // 5️⃣ Make async WebAPI call
                string postData = Newtonsoft.Json.JsonConvert.SerializeObject(p);
                string ret = await Task.Run(() =>
                    SJeMES_Framework.WebAPI.WebAPIHelper.Post(
                        Program.client.APIURL,
                        "KZ_QCO",
                        "KZ_QCO.Controllers.MESUpdateServer",
                        "GetProd2WHDataByCrd",
                        Program.client.UserToken,
                        postData
                    )
                );
                var retDict = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(ret);
                if (!Convert.ToBoolean(retDict["IsSuccess"]))
                {
                    SJeMES_Control_Library.MessageHelper.ShowErr(this, retDict["ErrMsg"].ToString());
                    return;
                }

                string json = retDict["RetData"].ToString();
                dtJson = SJeMES_Framework.Common.JsonHelper.GetDataTableByJson(json);

                if (dtJson == null || dtJson.Rows.Count == 0)
                {
                    SJeMES_Control_Library.MessageHelper.ShowErr(this,
                        "No Data Found!!");
                    return;
                }

                string[] dateColumns = { "CRD_DATE" };
                foreach (string col in dateColumns)
                {
                    if (!dtJson.Columns.Contains(col)) continue;

                    foreach (DataRow row in dtJson.Rows)
                    {
                        if (row[col] != DBNull.Value)
                        {
                            DateTime dt = Convert.ToDateTime(row[col]);
                            row[col] = dt.ToString("yyyy-MM-dd");
                        }
                    }
                    dtJson.Columns[col].DataType = typeof(string);
                }

                dataGridView2.DataSource = dtJson;

               

                ApplyDataGridViewStyles(dataGridView2);
                ApplyShippingStatusColors(dataGridView2);
                if (dtJson.Rows.Count > 0)
                {
                    await Task.Run(() => BuildProd2WHCharts(dtJson));
                }
                #region This Block For UCRollText 

                messages = BuildTickerMessages();
                // Reset ticker display
                messageIndex = 0;               
                ucRollText1.Text = messages[0];

                ucRollText1.AutoSize = false; // prevent shrinking
                ucRollText1.Size = new Size(600, 100); // make it tall enough

                ucRollText1.Left = (panel2.ClientSize.Width - ucRollText1.Width) / 2;
                ucRollText1.Top = (panel2.ClientSize.Height - ucRollText1.Height) / 2;

                #endregion


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                loadbl.Visible = false;
                loaderPictureBox.Visible = false;
                progressBar1.Visible = false;
                
            }
        }
        private void ApplyDataGridViewStyles(DataGridView dgv)
        {
            if (dgv == null) return;

            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 10, 0, 10);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 12, FontStyle.Bold);

            dgv.DefaultCellStyle.Font = new Font("Times New Roman", 11, FontStyle.Regular);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.GridColor = Color.Teal;

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.AutoResizeColumns();
        }
        private void ApplyShippingStatusColors(DataGridView dgv)
        {
            // Make sure the column exists
            if (!dgv.Columns.Contains("SHIPPING_STATUS"))
                return;

            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.Cells["SHIPPING_STATUS"].Value == null)
                    continue;

                string status = row.Cells["SHIPPING_STATUS"].Value.ToString().Trim().ToUpper();

                if (status == "SHIPPED ON TIME")
                {
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else if (status == "DELAYED")
                {
                    row.DefaultCellStyle.BackColor = Color.IndianRed;
                    row.DefaultCellStyle.ForeColor = Color.White;
                }
                else
                {
                    // Reset for other statuses
                    row.DefaultCellStyle.BackColor = Color.White;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
            }
        }




        private bool Validate_CRD_Date()
        {
            DateTime fromDate = dateTimePicker1.Value.Date;
            DateTime toDate = dateTimePicker2.Value.Date;

            DateTime maxAllowedDate = fromDate.AddMonths(3); 

            if (toDate > maxAllowedDate)
            {
                return false;
            }
            return true;
        }

        private void Export_Click(object sender, EventArgs e)
        {
            string a = "Prod2WHExport.xls";

            if (tabControl1.SelectedIndex == 0)
            {
                ExportExcels(a, dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 2)
            {
                ExportExcels(a, dataGridView3);
            }
            else
            {

                SJeMES_Control_Library.MessageHelper.ShowErr(this,
                    "No exportable tab selected.!!");
               
            }
        }

        private void ExportExcels(string fileName, DataGridView myDGV)
        {
            if (myDGV.Rows.Count == 0)
            {

                SJeMES_Control_Library.MessageHelper.ShowErr(this,
                    "No data to export.!!");
                return;
            }

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xlsx";
            saveDialog.Filter = "Excel Files|*.xlsx";
            saveDialog.FileName = fileName;

            if (saveDialog.ShowDialog() != DialogResult.OK)
                return;

            string saveFileName = saveDialog.FileName;

            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = xlApp.Workbooks.Add();
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                int columnCount = myDGV.Columns.Count;
                int rowCount = myDGV.AllowUserToAddRows
                    ? myDGV.Rows.Count - 1
                    : myDGV.Rows.Count;

                // Write headers
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = myDGV.Columns[i].HeaderText;
                }

                // Prepare data array
                object[,] objData = new object[rowCount, columnCount];

                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        objData[i, j] = myDGV.Rows[i].Cells[j].Value;
                    }
                }

                // Set range
                range = worksheet.Range[
                    worksheet.Cells[2, 1],
                    worksheet.Cells[rowCount + 1, columnCount]
                ];

                range.Value2 = objData;

                // Auto fit columns
                worksheet.Columns.AutoFit();

                // Save properly (IMPORTANT)
                workbook.SaveAs(saveFileName);
                workbook.Close();
                xlApp.Quit();

                MessageBox.Show("Successfully saved.",
                                "Message",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed:\n" + ex.Message);
            }
            finally
            {
                // Release COM objects properly
                if (range != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (xlApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            if(tabControl1.SelectedTab == tabPage5)
            {
                Load_Data(dtJson);

                dataGridView3.ShowCellToolTips = false;
                // Configure custom tooltip
                customToolTip.OwnerDraw = true;
                customToolTip.Draw += customToolTip_Draw;

            }
            if (tabControl1.SelectedTab == tabPage1)
            {
               
                DataTable sopTable = BuildSopDataTable(dateTimePicker1.Value);
                dataGridView1.DataSource = sopTable;

                dataGridView1.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 10, 0, 10);
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 12, FontStyle.Bold);

                dataGridView1.DefaultCellStyle.Font = new Font("Times New Roman", 11, FontStyle.Regular);
                dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dataGridView1.GridColor = Color.Teal;

                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;


            }

        }
        private void BuildProd2WHCharts(DataTable dtJson)
        {
            

            var crdGroups = dtJson.AsEnumerable()
    .GroupBy(r => Convert.ToDateTime(r["CRD_DATE"]).Date)
    .Select(g => new
    {
        CRD = g.Key,

        OnTimeSE = g.Where(r => r["SHIPPING_STATUS"].ToString().Contains("ON TIME") ||
                                r["SHIPPING_STATUS"].ToString().Contains("SHIPPED ON TIME"))
                    .Select(r => r["SALES_ORDER"].ToString()).Distinct().ToList(),

        DelayedSE = g.Where(r => r["SHIPPING_STATUS"].ToString().Contains("DELAYED"))
                    .Select(r => r["SALES_ORDER"].ToString()).Distinct().ToList(),

        ShippedLateSE = g.Where(r => r["SHIPPING_STATUS"].ToString().Contains("SHIPPED LATE"))
                    .Select(r => r["SALES_ORDER"].ToString()).Distinct().ToList(),

        AvgDelayDays = g.Average(r => Convert.ToDouble(r["SHIPPING_DELAY_DAYS"]))
    })
    .OrderBy(x => x.CRD)
    .ToList();

            // Totals
            totalOnTime = crdGroups.Sum(x => x.OnTimeSE.Count);
            totalDelayed = crdGroups.Sum(x => x.DelayedSE.Count);
            totalShippedLate = crdGroups.Sum(x => x.ShippedLateSE.Count);
            totalOrders = totalOnTime + totalDelayed + totalShippedLate;


            // --- Helper Panel Creator ---
            Panel CreateChartPanel(Chart chart, string titleText, string summaryText = null)
            {
                Panel panel = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(10),
                    BorderStyle = BorderStyle.FixedSingle
                };

                Label title = new Label
                {
                    Text = titleText,
                    Font = new Font("Times New Roman", 16, FontStyle.Bold),
                    Dock = DockStyle.Top,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Height = 35
                };



                panel.Controls.Add(chart);
                panel.Controls.Add(title);
                chart.Dock = DockStyle.Fill;

                if (!string.IsNullOrEmpty(summaryText))
                {
                    Label summary = new Label
                    {
                        Text = summaryText,
                        Font = new Font("Times New Roman", 12, FontStyle.Bold),
                        Dock = DockStyle.Bottom,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Height = 30
                    };
                    panel.Controls.Add(summary);
                }

                return panel;
            }

            // ================= PIE CHART =================
          

            Chart chartPie = new Chart();
            chartPie.ChartAreas.Add(new ChartArea());
            chartPie.Legends.Add(new Legend());

            Series pieSeries = new Series("Shipping")
            {
                ChartType = SeriesChartType.Pie,
                IsValueShownAsLabel = true
            };

            // Add three slices
            pieSeries.Points.AddXY("On Time", totalOnTime);
            pieSeries.Points.AddXY("Delayed", totalDelayed);
            pieSeries.Points.AddXY("Shipped Late", totalShippedLate);

            // Colors
            pieSeries.Points[0].Color = Color.Green;
            pieSeries.Points[1].Color = Color.Red;
            pieSeries.Points[2].Color = Color.Orange;

            // Distinct SE_ID lists
            var onTimeList = crdGroups
                .SelectMany(x => x.OnTimeSE)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            var delayedList = crdGroups
                .SelectMany(x => x.DelayedSE)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            var shippedLateList = crdGroups
                .SelectMany(x => x.ShippedLateSE)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            // Format nicely (numbered list)
            string onTimeFormatted = string.Join("\n",
                onTimeList.Select((id, index) => $"{index + 1}. {id}"));

            string delayedFormatted = string.Join("\n",
                delayedList.Select((id, index) => $"{index + 1}. {id}"));

            string shippedLateFormatted = string.Join("\n",
                shippedLateList.Select((id, index) => $"{index + 1}. {id}"));

            // Tooltips
            pieSeries.Points[0].ToolTip =
                $"ON TIME ORDERS: {totalOnTime}\n\nSE_ID LIST:\n{onTimeFormatted}";

            pieSeries.Points[1].ToolTip =
                $"DELAYED ORDERS: {totalDelayed}\n\nSE_ID LIST:\n{delayedFormatted}";

            pieSeries.Points[2].ToolTip =
                $"SHIPPED LATE ORDERS: {totalShippedLate}\n\nSE_ID LIST:\n{shippedLateFormatted}";

            chartPie.Series.Add(pieSeries);

            // Mouse click popup
            chartPie.MouseClick += (s, e) =>
            {
                HitTestResult result = chartPie.HitTest(e.X, e.Y);

                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    int pointIndex = result.PointIndex;
                    string content = chartPie.Series[0].Points[pointIndex].ToolTip;

                    Point location = chartPie.PointToScreen(new Point(e.X + 10, e.Y + 10));
                    ShowCopyableTooltip(content, location);
                }
            };

            // Summary text
            string pieSummaryText = $"Total Orders: {totalOrders} | On Time: {totalOnTime} | Delayed: {totalDelayed} | ShippedLate: {totalShippedLate}";
            Panel piePanel = CreateChartPanel(chartPie, "On-Time vs Delayed vs Shipped Late Orders", pieSummaryText);

            // ================= BAR CHART =================
            avgDelays = new Dictionary<string, double>
    {
        { "PC Split", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["PC_SPLIT_DELAY_DAYS"])) },
        { "ERP Pick", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["ERP_PICK_DELAY_DAYS"])) },
        { "WH Issue", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["WH_ISSUE_DELAY_DAYS"])) },

        { "Outsourcing", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["OUT_SRC_DELAY_DAYS"])) },

        { "Cut", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["CUT_DELAY_DAYS"])) },
        { "Stitch", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["STITCH_DELAY_DAYS"])) },
        { "Assemble", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["ASSY_DELAY_DAYS"])) },
        { "Pack", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["PACK_DELAY_DAYS"])) },
        { "FG", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["FG_WH_DELAY_DAYS"])) },
        { "Ship", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["SHIPPING_DELAY_DAYS"])) }
    };

         
            Chart chartBar = new Chart();
            ChartArea chartArea = new ChartArea();

            // Force all labels to display
            chartArea.AxisX.Interval = 1;
            chartArea.AxisX.LabelStyle.Interval = 1;

            chartBar.ChartAreas.Add(chartArea);
            chartBar.Legends.Add(new Legend());

            Series barSeries = new Series("Delays")
            {
                ChartType = SeriesChartType.Bar,
                IsValueShownAsLabel = true,
                Color = Color.Red
            };

            foreach (var stage in avgDelays)
            {
                int idx = barSeries.Points.AddXY(stage.Key, stage.Value);
                barSeries.Points[idx].Color = stage.Value > 0 ? Color.Red : Color.Green;
                barSeries.Points[idx].Label = stage.Value.ToString("F1");
                barSeries.Points[idx].ToolTip = $"{stage.Key} Avg Delay: {stage.Value:F1} Days";
            }

            chartBar.Series.Add(barSeries);

            // Attach double-click handler
            chartBar.DoubleClick += ChartBar_DoubleClick;
            Panel barPanel = CreateChartPanelWithInstruction(chartBar, "Average Delay Days by Stage");

            // ================= LINE CHART =================

            Chart chartLine = new Chart();
            ChartArea area = new ChartArea("MainArea");
            area.AxisX.Interval = 1;
            area.AxisX.LabelStyle.Angle = -45;
            area.AxisY.Title = "Order Count";
            chartLine.ChartAreas.Add(area);
            chartLine.Legends.Add(new Legend());

            // Delayed series
            Series delayedSeries = new Series("Delayed")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Red,
                BorderWidth = 3,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 8,
                IsValueShownAsLabel = true
            };

            // On Time series
            Series onTimeSeries = new Series("On Time")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Green,
                BorderWidth = 3,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 8,
                IsValueShownAsLabel = true
            };

            // Shipped Late series
            Series shippedLateSeries = new Series("Shipped Late")
            {
                ChartType = SeriesChartType.Line,
                Color = Color.Orange,
                BorderWidth = 3,
                MarkerStyle = MarkerStyle.Circle,
                MarkerSize = 8,
                IsValueShownAsLabel = true
            };

            foreach (var point in crdGroups)
            {
                string crdLabel = point.CRD.ToString("yyyy/MM/dd");

                int dIdx = delayedSeries.Points.AddXY(crdLabel, point.DelayedSE.Count);
                delayedSeries.Points[dIdx].ToolTip = $"{crdLabel} - Delayed: {point.DelayedSE.Count}";

                int oIdx = onTimeSeries.Points.AddXY(crdLabel, point.OnTimeSE.Count);
                onTimeSeries.Points[oIdx].ToolTip = $"{crdLabel} - On Time: {point.OnTimeSE.Count}";

                int lIdx = shippedLateSeries.Points.AddXY(crdLabel, point.ShippedLateSE.Count);
                shippedLateSeries.Points[lIdx].ToolTip = $"{crdLabel} - Shipped Late: {point.ShippedLateSE.Count}";
            }

            chartLine.Series.Add(delayedSeries);
            chartLine.Series.Add(onTimeSeries);
            chartLine.Series.Add(shippedLateSeries);

            Panel linePanel = CreateChartPanel(chartLine, "Shipping On-Time vs Delayed vs Shipped Late Trend by CRD Date");


            // ================= GROUPED COLUMN CHART =================
            Chart chartGrouped = new Chart();
            ChartArea groupedArea = new ChartArea("GroupedArea");
            groupedArea.AxisX.Interval = 1;
            groupedArea.AxisX.LabelStyle.Angle = -45;
            chartGrouped.ChartAreas.Add(groupedArea);
            chartGrouped.Legends.Add(new Legend());

            Series totalSeries = new Series("Total")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.Blue,
                IsValueShownAsLabel = true
            };

            Series onTimeSeriesGrouped = new Series("On Time")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.Green,
                IsValueShownAsLabel = true
            };

            Series delayedSeriesGrouped = new Series("Delayed")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.Red,
                IsValueShownAsLabel = true
            };

            Series shippedLateSeriesGrouped = new Series("Shipped Late")
            {
                ChartType = SeriesChartType.Column,
                Color = Color.Orange,
                IsValueShownAsLabel = true
            };

            foreach (var point in crdGroups)
            {
                string crdLabel = point.CRD.ToString("yyyy/MM/dd");
                int total = point.OnTimeSE.Count + point.DelayedSE.Count + point.ShippedLateSE.Count;

                int totalIdx = totalSeries.Points.AddXY(crdLabel, total);
                totalSeries.Points[totalIdx].ToolTip = $"{crdLabel} - Total: {total}";

                int onTimeIdx = onTimeSeriesGrouped.Points.AddXY(crdLabel, point.OnTimeSE.Count);
                onTimeSeriesGrouped.Points[onTimeIdx].ToolTip =
                    $"{crdLabel} - On Time: {point.OnTimeSE.Count}";

                int delayedIdx = delayedSeriesGrouped.Points.AddXY(crdLabel, point.DelayedSE.Count);
                delayedSeriesGrouped.Points[delayedIdx].ToolTip =
                    $"{crdLabel} - Delayed: {point.DelayedSE.Count}";

                int lateIdx = shippedLateSeriesGrouped.Points.AddXY(crdLabel, point.ShippedLateSE.Count);
                shippedLateSeriesGrouped.Points[lateIdx].ToolTip =
                    $"{crdLabel} - Shipped Late: {point.ShippedLateSE.Count}";
            }

            chartGrouped.Series.Add(totalSeries);
            chartGrouped.Series.Add(onTimeSeriesGrouped);
            chartGrouped.Series.Add(delayedSeriesGrouped);
            chartGrouped.Series.Add(shippedLateSeriesGrouped);

            string groupedSummary = $"Total SE_IDs: {totalOrders} | On-Time: {totalOnTime} | Delayed: {totalDelayed} | Shipped Late: {totalShippedLate}";
            Panel groupedPanel = CreateChartPanel(chartGrouped, "On-Time vs Delayed vs Shipped Late SE_IDs per CRD", groupedSummary);


            // ================= LAYOUT =================

            if (tabPage3.InvokeRequired)
            {
                tabPage3.Invoke(new Action(() =>
                {
                    tabPage3.Controls.Clear();

                    TableLayoutPanel topLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        Height = 350,
                        ColumnCount = 2
                    };
                    topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    topLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    topLayout.Controls.Add(piePanel, 0, 0);
                    topLayout.Controls.Add(barPanel, 1, 0);

                    TableLayoutPanel bottomLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 2
                    };
                    bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                    bottomLayout.Controls.Add(linePanel, 0, 0);
                    bottomLayout.Controls.Add(groupedPanel, 1, 0);

                    tabPage3.Controls.Add(bottomLayout);
                    tabPage3.Controls.Add(topLayout);
                }));
            }
            else
            {

                tabPage3.Controls.Clear();

            }



        }
        private Form tooltipForm = null;

        private void ShowCopyableTooltip(string content, Point location)
        {
            // Close old popup
            if (tooltipForm != null)
                tooltipForm.Close();

            tooltipForm = new Form();
            tooltipForm.FormBorderStyle = FormBorderStyle.FixedSingle;
            tooltipForm.StartPosition = FormStartPosition.Manual;
            tooltipForm.ShowInTaskbar = false;
            tooltipForm.TopMost = true;
            tooltipForm.Size = new Size(350, 300);
            tooltipForm.Location = location;

            TextBox txt = new TextBox();
            txt.Multiline = true;
            txt.ReadOnly = true;
            txt.ScrollBars = ScrollBars.Vertical;
            txt.Dock = DockStyle.Fill;
            txt.Text = content;

            tooltipForm.Controls.Add(txt);
            tooltipForm.Show();
        }

        private void ChartBar_DoubleClick(object sender, EventArgs e)
        {
            Chart chartBar = sender as Chart;
            if (chartBar == null) return;

            Panel parentPanel = chartBar.Parent as Panel;
            if (parentPanel == null) return;

            // Clear existing chart panel
            parentPanel.Controls.Clear();

            // Apply special rule
            foreach (DataRow row in dtJson.Rows)
            {
                bool cutStitchAssemZero =
                    Convert.ToInt32(row["CUT_QTY"]) == 0 &&
                    Convert.ToInt32(row["STITCH_QTY"]) == 0 &&
                    Convert.ToInt32(row["ASSEM_QTY"]) == 0;

                bool packFgMatch =
                    Convert.ToInt32(row["PACK_QTY"]) == Convert.ToInt32(row["FG_QTY"]);

                bool erpIssueInvalid =
                    string.IsNullOrEmpty(row["ERP_PICK_DATE"].ToString()) ||
                    string.IsNullOrEmpty(row["ISSUE_DATE"].ToString());

                if (cutStitchAssemZero && (packFgMatch || erpIssueInvalid))
                {
                    row["PC_SPLIT_DELAY_DAYS"] = 0;
                    row["ERP_PICK_DELAY_DAYS"] = 0;
                    row["WH_ISSUE_DELAY_DAYS"] = 0;
                }
            }

            // Recalculate averages
            var AVG = new Dictionary<string, double>
    {
        { "PC Split", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["PC_SPLIT_DELAY_DAYS"])) },
        { "ERP Pick", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["ERP_PICK_DELAY_DAYS"])) },
        { "WH Issue", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["WH_ISSUE_DELAY_DAYS"])) },
        { "Outsourcing", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["OUT_SRC_DELAY_DAYS"])) },
        { "Cut", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["CUT_DELAY_DAYS"])) },
        { "Stitch", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["STITCH_DELAY_DAYS"])) },
        { "Assemble", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["ASSY_DELAY_DAYS"])) },
        { "Pack", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["PACK_DELAY_DAYS"])) },
        { "FG", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["FG_WH_DELAY_DAYS"])) },
        { "Ship", dtJson.AsEnumerable().Average(r => Convert.ToDouble(r["SHIPPING_DELAY_DAYS"])) }
    };

            // Get Top 3 delays
            var top3Delays = AVG
                .OrderByDescending(x => x.Value)
                .Take(3)
                .ToList();

            // Build new bar chart
            Chart newChartBar = new Chart();
            newChartBar.Dock = DockStyle.Fill;

            ChartArea chartArea = new ChartArea();

            // 🔥 IMPORTANT FIX
            chartArea.AxisX.Interval = 1;
            chartArea.AxisX.LabelStyle.Interval = 1;

            // Optional (better visibility)
            chartArea.AxisX.LabelStyle.Angle = -45;
            chartArea.AxisX.LabelStyle.Font = new Font("Arial", 8);

            newChartBar.ChartAreas.Add(chartArea);
            newChartBar.Legends.Add(new Legend());

            Series barSeries = new Series("Delays")
            {
                ChartType = SeriesChartType.Bar,
                IsValueShownAsLabel = true,
                Color = Color.Red
            };

            foreach (var stage in AVG)
            {
                int idx = barSeries.Points.AddXY(stage.Key, stage.Value);
                barSeries.Points[idx].Color = stage.Value > 0 ? Color.Red : Color.Green;
                barSeries.Points[idx].Label = stage.Value.ToString("F1");

                string top3Summary = string.Join("\n", top3Delays.Select(d => $"{d.Key}: {d.Value:F1} days"));
                barSeries.Points[idx].ToolTip =
                    $"{stage.Key} Avg Delay: {stage.Value:F1} Days\n\nTop 3 Delays (Updated):\n{top3Summary}";
            }

            newChartBar.Series.Add(barSeries);

            // Wrap in a fresh panel with title
            Panel barPanel = CreateChartPanel2(newChartBar, "Average Delay Days by Stage");

            parentPanel.Controls.Add(barPanel);

            // Reset ticker
            var messages = new List<string>();
            if (totalOrders > 0)
            {
                var line1 = $"Total Sales Orders: {totalOrders} | On-Time: {totalOnTime} | Delayed: {totalDelayed} | ShippedLate {totalShippedLate}";
                var line2 = AVG != null && AVG.Any()
                    ? "Top 3 Delay Processes: " + string.Join(", ", AVG.OrderByDescending(x => x.Value).Take(3).Select(x => x.Key))
                    : "No delay data available";

                // Combine into one message with newline
                messages.Add(line1 + "\r\n" + line2);
            }
            else
            {
                messages.Add("No orders available");
            }
            messageIndex = 0;
            ucRollText1.Text = messages[0];
            ucRollText1.AutoSize = false;
            ucRollText1.Size = new Size(600, 100);
            ucRollText1.Left = (panel2.ClientSize.Width - ucRollText1.Width) / 2;
            ucRollText1.Top = (panel2.ClientSize.Height - ucRollText1.Height) / 2;
        }




        private void button2_Click(object sender, EventArgs e)
        {
            textBox_SeId.Text = "";
            Shipstatuscombo.Text = "";
            checkBox_CRD.Checked = false;
        }
        public void Load_Data(DataTable dt)
        {
            loadbl.Visible = true;

            dataGridView3.DataSource = dt;

            string[] columnsToHide =
            {
        "PRODUCTION_ORDER",
        "PC_SPLIT_DATE",
        "ERP_PICK_DATE",
        "ISSUE_DATE",
        "PC_SPLIT_STATUS",
        "ERP_PICK_STATUS",
        "WH_ISSUE_STATUS",
        "CUT_STATUS",
        "STITCH_STATUS",
        "ASSY_STATUS",
        "PACK_STATUS",
        "FG_WH_STATUS",
        "SHIPPING_STATUS",
         "OUTSRC_TARGET_QTY",
        "OUTSRC_IN_QTY",
        "OUTSRC_OUT_QTY",
        "OUTSRC_IN_DATE",
        "OUT_SRC_STATUS",
        "OUTSRC_OUT_DATE"

    };

            foreach (string col in columnsToHide)
            {
                if (dataGridView3.Columns.Contains(col))
                {
                    dataGridView3.Columns[col].Visible = false;
                }
            }

            // ===== HEADER STYLE =====
            dataGridView3.EnableHeadersVisualStyles = false;
            ApplyDataGridViewStyles(dataGridView3);
          //  ApplyShippingStatusColors(dataGridView3);

            loadbl.Visible = false;
        }


       
        private DataTable BuildSopDataTable(DateTime crd)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("CRD Date", typeof(string));
            dt.Columns.Add("Process", typeof(string));
            dt.Columns.Add("Scheduled Date", typeof(string));
            dt.Columns.Add("Days Before", typeof(string));

            // Add rows based on your SQL-like logic
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Planning", crd.AddDays(-14).ToString("yyyy/MM/dd"),"14");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"PC Split & Release", crd.AddDays(-13).ToString("yyyy/MM/dd"),"13");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"ERP Picking", crd.AddDays(-13).ToString("yyyy/MM/dd"), "13");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"WH Issue", crd.AddDays(-9).ToString("yyyy/MM/dd"),"9");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Cutting", crd.AddDays(-6).ToString("yyyy/MM/dd"),"6");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Stitching", crd.AddDays(-4).ToString("yyyy/MM/dd"),"4");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Assembly", crd.AddDays(-3).ToString("yyyy/MM/dd"), "3");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Packing", crd.AddDays(-3).ToString("yyyy/MM/dd"),"3");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"FG Warehouse", crd.AddDays(-2).ToString("yyyy/MM/dd"),"2");
            dt.Rows.Add(crd.ToString("yyyy/MM/dd"),"Shipment", crd.ToString("yyyy/MM/dd"),"<=1");
            // Trigger PictureBox redraw
            Soppicturebox.Invalidate();
            return dt;
        }
        private List<string> BuildTickerMessages()
        {
            var messages = new List<string>();
            if (totalOrders > 0)
            {
                var line1 = $"Total Sales Orders: {totalOrders} | On-Time: {totalOnTime} | Delayed: {totalDelayed} | Shippedlate: {totalShippedLate}";
                var line2 = avgDelays != null && avgDelays.Any()
                    ? "Top 3 Delay Processes: " + string.Join(", ", avgDelays.OrderByDescending(x => x.Value).Take(3).Select(x => x.Key))
                    : "No delay data available";

                // Combine into one message with newline
                messages.Add(line1 + "\r\n" + line2);
            }
            else
            {
                messages.Add("No orders available");
            }


            return messages;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

            DataTable sopTable = BuildSopDataTable(dateTimePicker1.Value);
            dataGridView1.DataSource = sopTable;

            dataGridView1.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 10, 0, 10);
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Teal;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 12, FontStyle.Bold);

            dataGridView1.DefaultCellStyle.Font = new Font("Times New Roman", 11, FontStyle.Regular);
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.GridColor = Color.Teal;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void dataGridView3_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var row = dataGridView3.Rows[e.RowIndex];

            string[] delayColumns =
            {

            "PC_SPLIT_DELAY_DAYS",
            "ERP_PICK_DELAY_DAYS",
            "WH_ISSUE_DELAY_DAYS",
            "CUT_DELAY_DAYS",
            "STITCH_DELAY_DAYS",
            "ASSY_DELAY_DAYS",
            "PACK_DELAY_DAYS",
            "FG_WH_DELAY_DAYS",
            "SHIPPING_DELAY_DAYS"
      
            };

            bool allZero = true;

            foreach (string col in delayColumns)
            {
                if (row.Cells[col].Value != null && int.TryParse(row.Cells[col].Value.ToString(), out int delayValue))
                {
                    if (delayValue != 0)
                    {
                        allZero = false;
                        break;
                    }
                }
            }

            // If all delay columns are 0, mark the row green
            if (allZero)
            {
                row.DefaultCellStyle.BackColor = Color.LightGreen;
                row.DefaultCellStyle.ForeColor = Color.Black;
            }
        }

        private void customToolTip_Draw(object sender, DrawToolTipEventArgs e)
        {
            e.Graphics.FillRectangle(Brushes.White, e.Bounds);

            using (Brush b = new SolidBrush(Color.Blue))
            {
                e.Graphics.DrawString(e.ToolTipText, new Font("Segoe UI", 9, FontStyle.Regular), b, e.Bounds);
            }
        }

        private void dataGridView3_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var row = dataGridView3.Rows[e.RowIndex];

                if (row.DefaultCellStyle.BackColor == Color.LightGreen)
                {
                    string message = "✅ All processes completed on time. No delays.";

                    var cellRect = dataGridView3.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    customToolTip.Show(message, dataGridView3, cellRect.Location.X + 20, cellRect.Location.Y + 20, 2000);
                }
            }
        }
        Panel CreateChartPanel2(Chart chart, string titleText, string summaryText = null)
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label title = new Label
            {
                Text = titleText,
                Font = new Font("Times New Roman", 16, FontStyle.Bold),
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 35
            };

            panel.Controls.Add(chart);
            panel.Controls.Add(title);
            chart.Dock = DockStyle.Fill;

            if (!string.IsNullOrEmpty(summaryText))
            {
                Label summary = new Label
                {
                    Text = summaryText,
                    Font = new Font("Times New Roman", 12, FontStyle.Bold),
                    Dock = DockStyle.Bottom,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Height = 30
                };
                panel.Controls.Add(summary);
            }

            return panel;
        }
        private Panel CreateChartPanelWithInstruction(Chart chart, string titleText)
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Title label
            Label title = new Label
            {
                Text = titleText,
                Font = new Font("Times New Roman", 16, FontStyle.Bold),
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                Height = 35
            };

            // Instruction label
            Label instruction = new Label
            {
                Text = "Double‑click the Graph",
                Font = new Font("Segoe UI", 11, FontStyle.Italic),
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.Teal,
                Height = 25
            };

            // Add controls
            panel.Controls.Add(chart);
            panel.Controls.Add(instruction);
            panel.Controls.Add(title);

            chart.Dock = DockStyle.Fill;

            return panel;
        }

        private void Soppicturebox_Paint(object sender, PaintEventArgs e)
        {
            using (Font font = new Font("Segoe UI", 10, FontStyle.Bold))
            using (Brush brush = new SolidBrush(Color.Black))
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)   // <-- Grid rows
                {
                    if (!row.IsNewRow)
                    {
                        string process = row.Cells["Process"].Value?.ToString();
                        string scheduledDate = row.Cells["Scheduled Date"].Value?.ToString();

                        if (process != null && processPositions.ContainsKey(process))
                        {
                            PointF pos = processPositions[process];
                            e.Graphics.DrawString(scheduledDate, font, brush, pos);
                        }
                    }
                }
            }
        }


        Dictionary<string, PointF> processPositions = new Dictionary<string, PointF>
{
    { "Planning", new PointF(50, 100) },
    { "PC Split & Release", new PointF(200, 150) },
    { "ERP Picking", new PointF(350, 200) },
    { "WH Issue", new PointF(500, 250) },
    { "Cutting", new PointF(650, 300) },
    { "Stitching", new PointF(800, 350) },
    { "Assembly", new PointF(950, 400) },
    { "Packing", new PointF(1100, 450) },
    { "FG Warehouse", new PointF(1250, 500) },
    { "Shipment", new PointF(1400, 550) }
};







    }
}
