using LinqToExcel;
using Microsoft.Win32;
using Milestone_Wpf.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace Milestone_Wpf
{
    public partial class MainWindow : Window
    {
        private static readonly string ChartName = "MilestoneChart";
        private static readonly List<string> ScriptFileNames = new List<string>() { "bootstrap.min.css", "plotly-2.3.0.min.js" };
        private static readonly string WorkSheetName = "Raw Data Milestones";
        public MainWindow()
        {
            InitializeComponent();
        }

        private static string CopyScripts(string milestoneDirName)
        {
            string scriptsDirName = $@"{milestoneDirName}\Scripts";
            if (!Directory.Exists(scriptsDirName))
            {
                _ = Directory.CreateDirectory(scriptsDirName);
            }
            ScriptFileNames.ForEach(scriptFileName =>
            {
                if (!File.Exists($@"{scriptsDirName}\{scriptFileName}"))
                {
                    File.Copy($@"scripts\{scriptFileName}", $@"{scriptsDirName}\{scriptFileName}", true);
                }
            });
            return scriptsDirName;
        }

        private static List<JSONData> CreateChartData(IOrderedEnumerable<GroupedData> groupedDatas)
        {
            List<JSONData> jSONDatas = new List<JSONData>();

            foreach (GroupedData dtf in groupedDatas)
            {
                string hoverText = string.Empty;
                int milestoneNumber = default;
                double sumRTO = default;
                double sumWorkloadmthGPID = default;

                dtf.Data.ForEach(dt =>
                {
                    if ((dt.Milestone != null) && (dt.GPIDDescription != null) && (dt.GPID_SubGroup != null))
                    {
                        hoverText += $"{dt.GPIDDescription}-{dt.Milestone}-{dt.GPID_SubGroup} <br>";
                        milestoneNumber += 1;
                    }
                    sumRTO += dt.RTO;
                    sumWorkloadmthGPID += dt.WorkloadmthGPID;
                });

                jSONDatas.Add(new JSONData
                {
                    Y = dtf.Year,
                    M = dtf.Month,
                    Ecc = dtf.EngineeringCenterCode,
                    Gpid = dtf.GPID,
                    Cd = new ChartData
                    {
                        Mt = hoverText != string.Empty ? hoverText : string.Empty,
                        Mn = milestoneNumber != default ? milestoneNumber : default,
                        Y1 = sumRTO != default ? sumRTO : default,
                        Y2 = sumWorkloadmthGPID != default ? sumWorkloadmthGPID : default
                    }
                });
            }

            return jSONDatas;
        }

        private static string CreateHTML(string fileName, string json)
        {
            StringBuilder html = new StringBuilder();
            _ = html.AppendLine(@"<!DOCTYPE html>");
            _ = html.AppendLine(@"<html lang=""en"">");
            _ = html.AppendLine(@"
<head>
    <meta charset=""utf-8"" />
    <title>Milestones</title>
    <link rel=""stylesheet"" href=""./Scripts/bootstrap.min.css"">
    <script src=""./Scripts/plotly-2.3.0.min.js""></script>
</head>");
            _ = html.AppendLine($@"
<body>
    <nav class=""navbar navbar-dark bg-dark"">
        <div class=""container"">
            <span class=""navbar-brand mb-0"">{fileName}</span>
        </div>
    </nav>
    <div class=""container"">
        <div class=""row justify-content-md-center align-items-center"">
            <div class=""col-md-auto"">
                <label class=""fs-6 fw-bold"">Year</label>
                <select id= ""Year"" class=""form-select"" multiple></select>
            </div>
            <div class=""col-md-auto"">
                <label class=""fs-6 fw-bold"">Engineering Center Code</label>
                <select id=""EngineeringCenterCode"" class=""form-select"" multiple></select>
            </div>
            <div class=""col-md-auto"">
                <label class=""fs-6 fw-bold"">GPID</label>
                <select id=""GPID"" class=""form-select"" multiple></select>
            </div>
            <div class=""col-md-auto"">
                <button id=""filterBtn"" type=""submit"" class=""btn btn-primary"">Filter</button>
                <button id=""resetBtn"" type=""reset"" class=""btn btn-secondary"">Reset</button>
            </div>
        </div>
    </div>
    <div id=""plot""></div>
    <div class=""container"">
        <div>
            <button id=""copyTableBtn"" type=""button"" class=""btn btn-outline-secondary btn-sm"" title=""Copy to Clipboard"" style = ""float: right;"">Copy Table</button>
        </div>
        <table id=""milestone_table"" class=""table table-hover table-sm table-bordered border-dark align-middle"">
            <caption>List of Milestones</caption>
            <thead>
                <tr class=""align-items-center"">
                    <th scope=""col"">Date</th>
                    <th scope=""col"">Milestones</th>
                </tr>
            </thead>
            <tbody id= ""milestone_tableBody"">
            </tbody>
        </table>
    </div>");
            _ = html.AppendLine($@"
    <script>
        const json = '{json}';");
            _ = html.AppendLine(@"
        var keys = [];
        var Years = [];
        var Months = [];
        var EngineeringCenterCodes = [];
        var GPIDs = [];
        var minRangeValue = '';
        var maxRangeValue = '';
        var milestone_x = [];
        var milestone_y = [];
        var milestone_text = [];
        var milestone_number = [];
        var rto_x = [];
        var rto_y = [];
        var WorkloadmthGPID_x = [];
        var WorkloadmthGPID_y = [];
        var filteredYears = [];
        var filteredEngineeringCenterCodes = [];
        var filteredGPIDs = [];
        var filterBtn = document.getElementById('filterBtn');
        filterBtn.addEventListener('click', filterKeys);
        function filterKeys()
        {
            filteredYears = [];
            filteredEngineeringCenterCodes = [];
            filteredGPIDs = [];
            var year = document.getElementById('Year');
            var selectedYears = year.selectedOptions;
            for (let i = 0; i < selectedYears.length; i++)
            {
                filteredYears.push(Number(selectedYears[i].label));
            }
            var ecc = document.getElementById('EngineeringCenterCode');
            var selectedEccs = ecc.selectedOptions;
            for (let i = 0; i < selectedEccs.length; i++)
            {
                filteredEngineeringCenterCodes.push(selectedEccs[i].label);
            }
            var gpid = document.getElementById('GPID');
            var selectedGPIDs = gpid.selectedOptions;
            for (let i = 0; i < selectedGPIDs.length; i++)
            {
                filteredGPIDs.push(selectedGPIDs[i].label);
            }
            drawPlot();
            populateTable();
        };
        var resetBtn = document.getElementById('resetBtn');
        resetBtn.addEventListener('click', resetKeys);
        function resetKeys()
        {
            var year = document.getElementById('Year');
            year.selectedIndex = -1;
            var ecc = document.getElementById('EngineeringCenterCode');
            ecc.selectedIndex = -1;
            var gpid = document.getElementById('GPID');
            gpid.selectedIndex = -1;
            filteredYears = [];
            filteredEngineeringCenterCodes = [];
            filteredGPIDs = [];
            drawPlot();
            populateTable();
        };
        var copyTableBtn = document.getElementById('copyTableBtn');
        copyTableBtn.addEventListener('click', copyTableToClipboard);
        function copyTableToClipboard()
        {
            var urlField = document.querySelector('table');
            var range = document.createRange();
            range.selectNode(urlField);
            window.getSelection().addRange(range);
            document.execCommand('copy');
        };
        function getData()
        {
            keys = JSON.parse(json);
            keys.shift();
            keys.forEach(key => {
                if (!(Years.includes(key.Y))) {
                    Years.push(key.Y);
                }
                if (!(Months.includes(key.M))) {
                    Months.push(key.M);
                }
                if (!(EngineeringCenterCodes.includes(key.Ecc))) {
                    EngineeringCenterCodes.push(key.Ecc);
                }
                if (!(GPIDs.includes(key.Gpid))) {
                    GPIDs.push(key.Gpid);
                }
            });
            var yearddb = document.getElementById('Year');
            Years.forEach(year => {
                var el = document.createElement('option');
                el.textContent = year;
                el.value = year;
                yearddb.appendChild(el);
            });
            var eccddb = document.getElementById('EngineeringCenterCode');
            EngineeringCenterCodes.forEach(ecc => {
                var el = document.createElement('option');
                el.textContent = ecc;
                el.value = ecc;
                eccddb.appendChild(el);
            });
            var gpidddb = document.getElementById('GPID');
            GPIDs.forEach(gpid => {
                var el = document.createElement('option');
                el.textContent = gpid;
                el.value = gpid;
                gpidddb.appendChild(el);
            });
        };
        function getYearKeys()
        {
            return ((filteredYears.length === 0) ? Years : filteredYears);
        };
        function getEngineeringCenterCodeKeys()
        {
            return ((filteredEngineeringCenterCodes.length === 0) ? EngineeringCenterCodes : filteredEngineeringCenterCodes);
        };
        function getGPIDKeys()
        {
            return ((filteredGPIDs.length === 0) ? GPIDs : filteredGPIDs);
        };
        function drawPlot()
        {
            milestone_x = [];
            milestone_y = [];
            milestone_text = [];
            milestone_number = [];
            rto_x = [];
            rto_y = [];
            WorkloadmthGPID_x = [];
            WorkloadmthGPID_y = [];
            var Years = getYearKeys();
            var EngineeringCenterCodes = getEngineeringCenterCodeKeys();
            var GPIDs = getGPIDKeys();
            minRangeValue = `${ Math.min(...Years)}-${ Math.min(...Months)}`;
            maxRangeValue = `${ Math.max(...Years)}-${ Math.max(...Months)}`;
            Years.forEach(year => {
                Months.forEach(month => {
                    var hoverText = '';
                    var milestoneNumber = 0;
                    var sumRTO = 0.0;
                    var sumWorkloadmthGPID = 0.0;
                    keys
                        .filter(key =>
                            (JSON.stringify(key.Y) === JSON.stringify(year)) &&
                            (JSON.stringify(key.M) === JSON.stringify(month)) &&
                            (EngineeringCenterCodes.includes(key.Ecc)) &&
                            (GPIDs.includes(key.Gpid)))
                        .forEach(dtf => {
                            if (dtf.Cd.Mt !== """") {
                                hoverText += dtf.Cd.Mt;
                                milestoneNumber += dtf.Cd.Mn;
                            }
                            if (dtf.Cd.Y1 !== 0.0) {
                                sumRTO += dtf.Cd.Y1;
                            }
                            if (dtf.Cd.Y2 !== 0.0) {
                                sumWorkloadmthGPID += dtf.Cd.Y2;
                            }
                        });
                    if (hoverText !== """") {
                        milestone_x.push(`${JSON.stringify(year)}-${JSON.stringify(month)}`);
                        milestone_y.push(0);
                        milestone_text.push(hoverText);
                        milestone_number.push(milestoneNumber);
                    }
                    if (sumRTO !== 0.0) {
                        rto_x.push(`${JSON.stringify(year)}-${JSON.stringify(month)}`);
                        rto_y.push(sumRTO);
                    }
                    if (sumWorkloadmthGPID !== 0.0) {
                        WorkloadmthGPID_x.push(`${JSON.stringify(year)}-${JSON.stringify(month)}`);
                        WorkloadmthGPID_y.push(sumWorkloadmthGPID);
                    }
                })
            });
            var milestoneData =
            {
                name: 'Milestones',
                type: 'scatter',
                mode: 'markers',
                x: milestone_x,
                y: milestone_y,
                text: milestone_text,
                hovertemplate: '<b>%{text}</b>',
                line: { color: 'rgb(128, 0, 128)' }
            };
            var milestoneNumber =
            {
                name: 'No of Milestones',
                x: milestone_x,
                y: milestone_y,
                mode: 'markers+text',
                text: milestone_number,
                textposition: 'bottom',
                textfont:
                    {
                        color: 'rgb(3, 70, 250)'
                    },
                type: 'scatter'
            };
            var rtoData =
            {
                name: 'Sum of RTO',
                type: 'scatter',
                mode: 'lines+markers',
                x: rto_x,
                y: rto_y
            };
            var WorkloadmthGPIDData =
            {
                name: 'Sum of WorkloadmthGPID',
                type: 'scatter',
                mode: 'lines+markers',
                x: WorkloadmthGPID_x,
                y: WorkloadmthGPID_y
            };
            var data = [milestoneNumber, milestoneData, rtoData, WorkloadmthGPIDData];
            var layout =
            {
                title: 'Milestones vs RTO vs WorkloadmthGPID',
                showlegend: true,
                xaxis:  {
                    autorange: true,
                    range:[minRangeValue, maxRangeValue],
                    rangeselector:  {
                        buttons:    [
                            {
                                count: 1,
                                label: '1m',
                                step: 'month',
                                stepmode: 'backward'
                            },
                            {
                                count: 6,
                                label: '6m',
                                step: 'month',
                                stepmode: 'backward'
                            },
                            { step: 'all' }
                        ]
                    },
                    rangeslider: { range:[minRangeValue, maxRangeValue] },
                    type: 'date'
                },
                yaxis:  {
                    autorange: true,
                    type: 'linear'
                }
            };
            var config = {
                toImageButtonOptions: {
                    format: 'svg', // one of png, svg, jpeg, webp
                    filename: `MilestoneCapture-${new Date().toLocaleDateString().replaceAll('/', '_') + '_' + new Date().toLocaleTimeString().replaceAll(' ', '').replaceAll(':', '_')}`,
                    scale: 1 // Multiply title/legend/axis/canvas sizes by this factor
                },
                scrollZoom: true,
                displayModeBar: true,
                displaylogo: false,
                responsive: true
            };
            Plotly.newPlot('plot', data, layout, config);
        };
        function populateTable()
        {
            var rowCount = milestone_table.rows.length;
            for (var i = rowCount - 1; i > 0; i--)
            {
                milestone_table.deleteRow(i);
            }
            var tableBodyRef = document.getElementById('milestone_tableBody');
            for (let i = 0; i < milestone_x.length; i++)
            {
                var newRow = tableBodyRef.insertRow(tableBodyRef.rows.length);
                newRow.innerHTML = `<th scope=""row"">${milestone_x[i]}<br>(${milestone_number[i]})</th><td>${milestone_text[i]}</td>`;
            }
        };
        getData();
        drawPlot();
        populateTable();
    </script>
</body>");
            _ = html.AppendLine(@"</html>");
            return html.ToString();
        }

        private static void CreateHTMLFile(string html, string milestoneDirName, string htmlFileName) => File.WriteAllText($@"{milestoneDirName}\{htmlFileName}", html);

        private static string CreateJSON(List<JSONData> jSONDatas) => JsonConvert.SerializeObject(jSONDatas);

        private static string CreateMilestoneDirectory(string directory)
        {
            string milestoneDirName = $@"{directory}\MileStone";
            if (!Directory.Exists(milestoneDirName))
            {
                _ = Directory.CreateDirectory(milestoneDirName);
            }
            return milestoneDirName;
        }

        private static string CreateOutputDirectoryAndFile(string html, string directory)
        {
            string milestoneDirName = CreateMilestoneDirectory(directory);
            _ = CopyScripts(milestoneDirName);
            string htmlFileName = $@"{ChartName}_{DateTime.Now:dd_MM_yy_HH_mm}.html";
            CreateHTMLFile(html, milestoneDirName, htmlFileName);
            return $@"{milestoneDirName}\{htmlFileName}";
        }

        private static (string fileName, string directory) GetFileInfo(string filePath) => (Path.GetFileName(filePath), Path.GetDirectoryName(filePath));

        private static IOrderedEnumerable<GroupedData> LoadExcel(string filePath)
        {
            List<RawData> rawData = new ExcelQueryFactory($@"{filePath}").Worksheet<RawData>(WorkSheetName).ToList();
            return rawData
                .GroupBy(rd => new { rd.Year, rd.Month, rd.EngineeringCenterCode, rd.GPID })
                .Where(x => x.Key.Year != 0)
                .Select(gd => new GroupedData()
                {
                    Year = gd.Key.Year,
                    Month = gd.Key.Month,
                    EngineeringCenterCode = gd.Key.EngineeringCenterCode,
                    GPID = gd.Key.GPID,
                    Data = gd.Select(x => new MilestoneData()
                    {
                        GPIDDescription = x.GPIDDescription,
                        RTO = x.RTO,
                        WorkloadmthGPID = x.WorkloadmthGPID,
                        Milestone = x.Milestone,
                        GPID_SubGroup = x.GPID_SubGroup
                    }).ToList()
                })
               .OrderBy(x => x?.Year)
               .ThenBy(x => x?.Month)
               .ThenBy(x => x.EngineeringCenterCode)
               .ThenBy(x => x.GPID)
               .ThenByDescending(x => x?.Data.Count);
        }

        private static void OpenHTML(string htmlFilePath) => _ = Process.Start($@"{htmlFilePath}");

        private void BtnOpenFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Binary Workbooks (.xlsb)| *.xlsb",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filePath in openFileDialog.FileNames)
                {
                    ProcessFile(filePath);
                }
            }
        }

        private void Log(string message) => _ = logList.Items.Add(message);

        private void ProcessFile(string filePath)
        {
            try
            {
                (string fileName, string directory) = GetFileInfo(filePath);
                Log($"Loading {fileName}");
                List<JSONData> jSONDatas = CreateChartData(LoadExcel(filePath));
                string json = CreateJSON(jSONDatas);
                //File.WriteAllText($@"{directory}\data_{DateTime.Now:dd_MM_yy_HH_mm}.json", json);
                string html = CreateHTML(fileName, json);
                string htmlFilePath = CreateOutputDirectoryAndFile(html, directory);
                Log($@"Opening {htmlFilePath}");
                OpenHTML(htmlFilePath);
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
        }
    }
}