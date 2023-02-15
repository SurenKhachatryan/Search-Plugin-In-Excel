using ExcelFindMatchRows.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFindMatchRows
{
    public partial class Ribbon1
    {
        public CancellationTokenSource CancelTokenSource;
        public CancellationToken CancelToken;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            CancelTokenSource = new CancellationTokenSource();
            CancelToken = CancelTokenSource.Token;
        }

        private async void Search_Button_Click(object sender, RibbonControlEventArgs e)
        {
            await Task.Run(async () =>
            {
                try
                {
                    var searchResultTab = "Search Result";

                    buttonSearch.Visible = false;
                    buttonCencel.Visible = true;
                    ProgressLabel.Visible = true;

                    if (string.IsNullOrWhiteSpace(searchEditBox.Text))
                    {
                        MessageBox.Show("Please enter text to search");
                        Restart();
                        return;
                    }

                    var application = Globals.ThisAddIn.GetApplication();

                    foreach (Excel.Worksheet workSheet in application.Worksheets)
                    {
                        if (workSheet.Name == searchResultTab)
                        {
                            try
                            {
                                workSheet.Delete();
                            }
                            catch
                            {
                                MessageBox.Show("Please Press Esc Then Press Search");
                                Restart();
                                return;
                            }
                            break;
                        }
                    }

                    var results = new List<ResultModel>();

                    foreach (Excel.Worksheet sheet in (Excel.Sheets)application.Worksheets)
                    {
                        if (sheet.Name == searchResultTab)
                            continue;

                        results.Add(new ResultModel()
                        {
                            TableName = sheet.Name,
                            Rows = await FindAll(sheet, searchEditBox.Text)
                        });
                    }

                    if (results.SelectMany(x => x.Rows).Count() == 0)
                    {
                        MessageBox.Show("Nothing found for your request");
                        Restart();
                        return;
                    }

                    if (!TryCreateWorkSheet(application, searchResultTab, out var resultWorkSheet))
                    {
                        MessageBox.Show($"{searchResultTab} failed to create the problem may be related to the document mode or other problems.");
                        Restart();
                        return;
                    }

                    InsertResultData(searchResultTab, results, resultWorkSheet);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }, CancelToken);

            if (CancelTokenSource.IsCancellationRequested)
            {
                CancelTokenSource = new CancellationTokenSource();
                CancelToken = CancelTokenSource.Token;
            }

            Restart();
        }

        private void InsertResultData(string searchResultTab, List<ResultModel> results, Worksheet resultWorkSheet)
        {
            ((Excel.Range)resultWorkSheet.Cells[3, 1]).Value = searchEditBox.Text;

            var startWithRow = 6;

            foreach (var result in results.Where(x => x.Rows.Any() && x.TableName != searchResultTab)
                                          .OrderBy(x => x.TableName)
                                          .ToList())
            {
                if (CancelToken.IsCancellationRequested)
                {
                    CancelToken.ThrowIfCancellationRequested();
                }

                for (int i = startWithRow; i <= result.Rows.Count + startWithRow - 1; i++)
                {
                    if (CancelToken.IsCancellationRequested)
                    {
                        CancelToken.ThrowIfCancellationRequested();
                    }

                    var row = result.Rows[i - startWithRow];

                    row.Copy(((Excel.Range)resultWorkSheet.Range[resultWorkSheet.Cells[i, 2], resultWorkSheet.Cells[i, row.Columns.Count + 1]]));

                    ((Excel.Range)resultWorkSheet.Cells[i, 1]).Value = result.TableName;
                }

                startWithRow += result.Rows.Count;
            }
        }

        private void buttonCencel_Click(object sender, RibbonControlEventArgs e)
        {
            buttonCencel.Visible = true;
            CancelTokenSource.Cancel();
        }

        private bool TryCreateWorkSheet(Excel.Application application, string name, out Excel.Worksheet worksheet)
        {
            try
            {
                worksheet = (Excel.Worksheet)application.Worksheets.Add();
                worksheet.Name = name;
                return true;
            }
            catch
            {
                worksheet = null;
                return false;
            }
        }

        private async Task<List<Excel.Range>> FindAll(Excel.Worksheet sheet, string search)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            var range = sheet.Range[((Excel.Range)sheet.Cells[1, 1]), ((Excel.Range)sheet.Cells[sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count])];

            var response = new List<Excel.Range>();

            currentFind = range.Find(search, LookIn: Excel.XlFindLookIn.xlValues);

            while (currentFind != null)
            {
                var fisrtCells = (Excel.Range)sheet.Cells[currentFind.Row, 1];
                var lastCells = (Excel.Range)sheet.Cells[currentFind.Row, sheet.UsedRange.Columns.Count];

                response.Add(sheet.get_Range(fisrtCells.Address, lastCells.Address));

                if (CancelToken.IsCancellationRequested)
                    CancelToken.ThrowIfCancellationRequested();

                if (firstFind == null)
                {
                    firstFind = currentFind;
                }
                else if (currentFind.Address == firstFind.Address)
                {
                    break;
                }

                currentFind = range.FindNext(lastCells);
            }

            return await Task.FromResult(response);
        }

        private void Restart()
        {
            buttonSearch.Visible = true;
            buttonCencel.Visible = false;
            ProgressLabel.Visible = false;
        }
    }
}
