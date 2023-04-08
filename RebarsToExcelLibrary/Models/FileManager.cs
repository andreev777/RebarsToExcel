using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RebarsToExcel.Models
{
    public static class FileManager
    {
        private static readonly Color firstRowCommonСolor = Color.FromArgb(191, 228, 204);
        private static readonly Color firstRowRebarClassСolor = Color.FromArgb(172, 184, 219);
        private static readonly Color rebarClassDarkСolor = Color.FromArgb(235, 238, 246);

        public static void Save(List<RebarAssembly> rebarAssemblies, string fileName)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            Workbook workbook = app.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets["Лист1"];

            SetDataToWorksheet(worksheet, rebarAssemblies);

            FormatDataWorksheet(worksheet);

            SaveFile(app, workbook, fileName);
        }

        private static void SetDataToWorksheet(Worksheet worksheet, List<RebarAssembly> rebarAssemblies)
        {
            worksheet.Range["B1"].Value = "Поз.";
            worksheet.Range["C1"].Value = "Наименование";
            worksheet.Range["D1"].Value = "Кол.";
            worksheet.Range["E1"].Value = "Масса ед., кг";
            worksheet.Range["F1"].Value = "Арматура изделия";

            #region Получаем список всех уникальных классов и диаметров арматуры

            var allRebarClasses = rebarAssemblies
                .SelectMany(x => x.Rebars
                .Select(rebar => rebar))
                .OrderBy(rebar => rebar.Class)
                .ThenBy(rebar => rebar.Diameter)
                .Select(rebar => rebar.GetClassDiameterInfoString())
                .Distinct()
                .ToList();

            int counter = 0;
            while (counter < allRebarClasses.Count)
            {
                worksheet.Cells[1, 7 + counter].Value = allRebarClasses[counter];
                counter++;
            }

            #endregion Получаем список всех уникальных классов и диаметров арматуры

            for (int i = 0; i < rebarAssemblies.Count; i++)
            {
                var rebarAssembly = rebarAssemblies[i];

                worksheet.Range["B" + (2 + i)].Value = rebarAssembly.Mark;
                worksheet.Range["C" + (2 + i)].Value = rebarAssembly.Type;
                worksheet.Range["D" + (2 + i)].Value = rebarAssembly.Count;
                worksheet.Range["E" + (2 + i)].Value = rebarAssembly.Mass;
                worksheet.Range["F" + (2 + i)].Value = rebarAssembly.GetRebarsInfoString();

                foreach (var rebar in rebarAssembly.Rebars)
                {
                    var rebarClassDiameterInfo = rebar.GetClassDiameterInfoString();
                    var rebarTotalCount = rebar.Count * rebarAssembly.Count;

                    if (rebarClassDiameterInfo == worksheet.Range["G1"].Value) { var condition = worksheet.Range["G" + (2 + i)].Value == null? worksheet.Range["G" + (2 + i)].Value = rebarTotalCount : worksheet.Range["G" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["H1"].Value) { var condition = worksheet.Range["H" + (2 + i)].Value == null? worksheet.Range["H" + (2 + i)].Value = rebarTotalCount : worksheet.Range["H" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["I1"].Value) { var condition = worksheet.Range["I" + (2 + i)].Value == null? worksheet.Range["I" + (2 + i)].Value = rebarTotalCount : worksheet.Range["I" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["J1"].Value) { var condition = worksheet.Range["J" + (2 + i)].Value == null? worksheet.Range["J" + (2 + i)].Value = rebarTotalCount : worksheet.Range["J" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["K1"].Value) { var condition = worksheet.Range["K" + (2 + i)].Value == null? worksheet.Range["K" + (2 + i)].Value = rebarTotalCount : worksheet.Range["K" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["L1"].Value) { var condition = worksheet.Range["L" + (2 + i)].Value == null? worksheet.Range["L" + (2 + i)].Value = rebarTotalCount : worksheet.Range["L" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["M1"].Value) { var condition = worksheet.Range["M" + (2 + i)].Value == null? worksheet.Range["M" + (2 + i)].Value = rebarTotalCount : worksheet.Range["M" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["N1"].Value) { var condition = worksheet.Range["N" + (2 + i)].Value == null? worksheet.Range["N" + (2 + i)].Value = rebarTotalCount : worksheet.Range["N" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["O1"].Value) { var condition = worksheet.Range["O" + (2 + i)].Value == null? worksheet.Range["O" + (2 + i)].Value = rebarTotalCount : worksheet.Range["O" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["P1"].Value) { var condition = worksheet.Range["P" + (2 + i)].Value == null? worksheet.Range["P" + (2 + i)].Value = rebarTotalCount : worksheet.Range["P" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["Q1"].Value) { var condition = worksheet.Range["Q" + (2 + i)].Value == null? worksheet.Range["Q" + (2 + i)].Value = rebarTotalCount : worksheet.Range["Q" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["R1"].Value) { var condition = worksheet.Range["R" + (2 + i)].Value == null? worksheet.Range["R" + (2 + i)].Value = rebarTotalCount : worksheet.Range["R" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["S1"].Value) { var condition = worksheet.Range["S" + (2 + i)].Value == null? worksheet.Range["S" + (2 + i)].Value = rebarTotalCount : worksheet.Range["S" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["T1"].Value) { var condition = worksheet.Range["T" + (2 + i)].Value == null? worksheet.Range["T" + (2 + i)].Value = rebarTotalCount : worksheet.Range["T" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["U1"].Value) { var condition = worksheet.Range["U" + (2 + i)].Value == null? worksheet.Range["U" + (2 + i)].Value = rebarTotalCount : worksheet.Range["U" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["V1"].Value) { var condition = worksheet.Range["V" + (2 + i)].Value == null? worksheet.Range["V" + (2 + i)].Value = rebarTotalCount : worksheet.Range["V" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["W1"].Value) { var condition = worksheet.Range["W" + (2 + i)].Value == null? worksheet.Range["W" + (2 + i)].Value = rebarTotalCount : worksheet.Range["W" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["X1"].Value) { var condition = worksheet.Range["X" + (2 + i)].Value == null? worksheet.Range["X" + (2 + i)].Value = rebarTotalCount : worksheet.Range["X" + (2 + i)].Value += rebarTotalCount; }
                    else if (rebarClassDiameterInfo == worksheet.Range["Y1"].Value) { var condition = worksheet.Range["Y" + (2 + i)].Value == null? worksheet.Range["Y" + (2 + i)].Value = rebarTotalCount : worksheet.Range["Y" + (2 + i)].Value += rebarTotalCount; }
                }
            }
        }

        private static void FormatDataWorksheet(Worksheet worksheet)
        {
            #region Задаем выравнивание текста в ячейках

            worksheet.UsedRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            worksheet.UsedRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Range["C1"].EntireColumn.IndentLevel = 1;
            worksheet.Range["F1"].EntireColumn.IndentLevel = 1;

            worksheet.Range["C1"].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            worksheet.Range["F1"].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            worksheet.Range["A1"].EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            worksheet.Range["A1"].EntireRow.WrapText = true;
            worksheet.Range["B1"].EntireColumn.WrapText = true;

            #endregion Задаем выравнивание текста в ячейках

            #region Задаем ширину столбцов и высоту строк

            worksheet.Range["B1"].ColumnWidth = 8; //Поз.
            worksheet.Range["C1"].ColumnWidth = 35; //Наименование
            worksheet.Range["D1"].ColumnWidth = 8; //Кол.
            worksheet.Range["E1"].ColumnWidth = 10; //Масса ед., кг
            worksheet.Range["F1"].ColumnWidth = 30; //Арматура изделия
            worksheet.Range["G1:Y1"].ColumnWidth = 10; //Классы арматуры

            worksheet.Range["A1"].RowHeight = 40; //Высота шапки

            foreach (Range row in worksheet.UsedRange.Rows)
            {
                if (row.Height < 30) row.RowHeight = 30;
            }

            #endregion Задаем ширину столбцов и высоту строк

            #region Задаем границы ячеек

            worksheet.UsedRange.Borders.Weight = XlBorderWeight.xlThin;

            for (int i = 0; i < worksheet.UsedRange.Rows.Count; i++)
            {
                worksheet.Range["F" + (1 + i)].Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThick;
            }

            for (int i = 0; i < worksheet.UsedRange.Columns.Count; i++)
            {
                worksheet.Cells[1, 1 + i].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            }

            #endregion Задаем границы ячеек

            #region Задаем цвета ячеек

            worksheet.Range["A1:F1"].Interior.Color = firstRowCommonСolor; //Ячейки с данными каркасов
            worksheet.Range[worksheet.Cells[1, 7], worksheet.Cells[1, worksheet.UsedRange.Columns.Count]].Interior.Color = firstRowRebarClassСolor; //Ячейки с классами арматуры каркасов

            for (int i = 0; i < worksheet.UsedRange.Columns.Count - 6; i++)
            {
                worksheet.Range[worksheet.Cells[2, 7 + i], worksheet.Cells[worksheet.UsedRange.Rows.Count, 7 + i]].Interior.Color = rebarClassDarkСolor;
                i++;
            }

            #endregion Задаем цвета ячеек

            worksheet.Range["A1"].EntireRow.Font.Bold = true;

            //Закрепляем первую строку
            worksheet.Activate();
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;
        }

        private static void SaveFile(Microsoft.Office.Interop.Excel.Application app, Workbook workbook, string fileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                DefaultExt = "xlsx",
                AddExtension = true,
                Filter = "Книга Excel |*.xlsx",
                FileName = fileName,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            try
            {
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(Filename: saveFileDialog.FileName, AccessMode: XlSaveAsAccessMode.xlNoChange, ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);

                    DialogResult dialogResult = MessageBox.Show("Открыть созданный файл?", "Открыть файл", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        OpenCreatedFile(saveFileDialog.FileName);
                    }
                }
            }
            catch
            {
                MessageBox.Show("Не удалось перезаписать файл. Закройте текущую таблицу Excel", "Ошибка");
            }
            finally
            {
                workbook.Close(0);
                app.Quit();
            }
        }

        private static void OpenCreatedFile(string path)
        {
            FileInfo file = new FileInfo(path);

            if (file.Exists)
            {
                System.Diagnostics.Process.Start(path);
            }
            else
            {
                MessageBox.Show("Не удалось открыть файл. Проверьте расположение файла.", "Ошибка");
            }
        }
    }
}