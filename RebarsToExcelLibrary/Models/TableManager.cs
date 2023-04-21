using Microsoft.Office.Interop.Excel;
using Prism.Mvvm;
using RebarsToExcel.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using IExcel = Microsoft.Office.Interop.Excel;

namespace RebarsToExcel.Models
{
    public class TableManager : BindableBase
    {
        private readonly IList<Bar> _bars;
        private readonly IList<RebarAssembly> _rebarAssemblies;

        private readonly Color headerСolor = Color.FromArgb(191, 228, 204);
        private readonly Color rebarOfAssemblyСolor = Color.FromArgb(232, 233, 233);

        #region ЗАГОЛОВКИ ТАБЛИЦЫ
        private readonly string CODE        = "A"; //Код изделия
        private readonly string MARK        = "B"; //Марка
        private readonly string TYPE        = "C"; //Тип изделия
        private readonly string UNITS       = "D"; //Ед. измерения
        private readonly string COUNT       = "E"; //Кол.
        private readonly string MASS        = "F"; //Масса ед.
        private readonly string DIAMETER    = "G"; //Диаметр
        private readonly string CLASS       = "H"; //Класс
        private readonly string LENGTH      = "I"; //Длина
        private readonly string ITEMCOUNT   = "J"; //Кол. заготовок
        private readonly string ITEMMASS    = "K"; //Масса заготовки
        private readonly string MASSTOTAL   = "L"; //Всего, кг
        private readonly string COUNSTRTYPE = "M"; //Тип основы
        private readonly string CONSTRMARK  = "N"; //Метка основы
        private readonly string LEVEL       = "O"; //Этаж
        private readonly string DRAWING     = "P"; //Чертеж
        private readonly string IMAGE       = "Q"; //Эскиз
        #endregion

        #region СВОЙСТВА PROGRESSBAR
        private double _progressCounter;
        public double ProgressCounter
        {
            get => _progressCounter;
            set
            {
                _progressCounter = value;
                RaisePropertyChanged(nameof(ProgressCounter));
            }
        }

        public double ProgressTotalCount { get; private set; }
        #endregion

        public TableManager(IList<Bar> bars, IList<RebarAssembly> rebarAssemblies)
        {
            _bars = bars;
            _rebarAssemblies = rebarAssemblies;

            ProgressTotalCount = _bars.Count + _rebarAssemblies.Count + 9; //Добавляем 9 позиций на форматирование таблицы
        }

        public void CreateTable()
        {
            IExcel.Application app = new IExcel.Application();
            app.Visible = false;
            app.DisplayAlerts = false;

            Workbook workbook = app.Workbooks.Add();
            Worksheet worksheet = workbook.Worksheets["Лист1"];

            var tableCreationProgressWindow = new TableCreationProgressWindow(this);
            var backgroundWorker = new BackgroundWorker();

            backgroundWorker.DoWork += (sender, e) =>
            {
                SetDataToWorksheet(worksheet);
                FormatDataWorksheet(worksheet);
            };

            backgroundWorker.RunWorkerCompleted += (sender, e) =>
            {
                if (tableCreationProgressWindow != null)
                {
                    tableCreationProgressWindow.Close();
                }

                SaveFile(app, workbook);
            };

            tableCreationProgressWindow.Show();

            backgroundWorker.RunWorkerAsync();
        }

        private void SetDataToWorksheet(Worksheet worksheet)
        {
            //Предварительное заполнение шапки для получения корректного значения UsedRange
            worksheet.Range["B1"].Value = ProjectData.ProjectCode;
            worksheet.Range["B2"].Value = ProjectData.ProjectName;
            worksheet.Range["B3"].Value = ProjectData.BuildingName;
            worksheet.Range["A4"].Value = "Заглушка";

            var allUniqSizes = GetAllUniqSizes(_bars);
            SetSizeHeader(worksheet, allUniqSizes);

            int counter = 0;

            //Заполнение сборочных единиц
            foreach (var rebarAssembly in _rebarAssemblies)
            {
                worksheet.Range[CODE + (5 + counter)].Value = GetRebarAssemblyCode(rebarAssembly);
                worksheet.Range[MARK + (5 + counter)].Value = rebarAssembly.Mark;
                worksheet.Range[TYPE + (5 + counter)].Value = rebarAssembly.Type;
                worksheet.Range[UNITS + (5 + counter)].Value = "шт.";
                worksheet.Range[COUNT + (5 + counter)].Value = rebarAssembly.Count;
                worksheet.Range[MASS + (5 + counter)].Value = rebarAssembly.Mass;
                worksheet.Range[MASSTOTAL + (5 + counter)].Value = rebarAssembly.Mass * rebarAssembly.Count;
                worksheet.Range[COUNSTRTYPE + (5 + counter)].Value = rebarAssembly.ConstructionType;
                worksheet.Range[CONSTRMARK + (5 + counter)].Value = rebarAssembly.ConstructionMark == "(нет)" ? string.Empty : rebarAssembly.ConstructionMark;
                worksheet.Range[LEVEL + (5 + counter)].Value = rebarAssembly.Level.Name == "(нет)" ? string.Empty : rebarAssembly.Level.Name;
                worksheet.Range[DRAWING + (5 + counter)].Value = rebarAssembly.Definition;

                var rebars = rebarAssembly.Rebars;
                if (rebars.Any())
                {
                    foreach (var rebar in rebars)
                    {
                        counter++;

                        worksheet.Range[DIAMETER + (5 + counter)].Value = rebar.Diameter;
                        worksheet.Range[CLASS + (5 + counter)].Value = rebar.Class;
                        worksheet.Range[LENGTH + (5 + counter)].Value = rebar.Length;
                        worksheet.Range[ITEMCOUNT + (5 + counter)].Value = rebar.Count;
                        worksheet.Range[ITEMMASS + (5 + counter)].Value = rebar.Mass;
                        worksheet.UsedRange.Rows[5 + counter].Columns.Interior.Color = rebarOfAssemblyСolor;
                    }
                }
                counter++;
                ProgressCounter++;
            }

            //Заполнение деталей
            foreach (var bar in _bars)
            {
                worksheet.Range[MARK + (5 + counter)].Value = bar.PositionWithShapeMark;
                worksheet.Range[TYPE + (5 + counter)].Value = GetBarType(bar);
                worksheet.Range[UNITS + (5 + counter)].Value = bar.CountTypeInfo;
                worksheet.Range[COUNT + (5 + counter)].Value = bar.CountType == 2 ? Math.Round(bar.Count, 0) : bar.Count;
                worksheet.Range[MASS + (5 + counter)].Value = bar.Mass;
                worksheet.Range[DIAMETER + (5 + counter)].Value = bar.Diameter;
                worksheet.Range[CLASS + (5 + counter)].Value = bar.Class;
                worksheet.Range[LENGTH + (5 + counter)].Value = bar.CountType == 2 ? string.Empty : bar.Length.ToString();
                worksheet.Range[MASSTOTAL + (5 + counter)].Value = bar.CountType == 2 ? Math.Round(Math.Round(bar.Count, 0) * bar.Mass, 2) : bar.Count * bar.Mass;
                worksheet.Range[COUNSTRTYPE + (5 + counter)].Value = bar.ConstructionType;
                worksheet.Range[CONSTRMARK + (5 + counter)].Value = bar.ConstructionMark == "(нет)" ? string.Empty : bar.ConstructionMark;
                worksheet.Range[LEVEL + (5 + counter)].Value = bar.Level.Name == "(нет)" ? string.Empty : bar.Level.Name;

                try
                {
                    SetShapeImage(worksheet, bar, counter);
                }
                catch { }

                SetSizesInCells(worksheet, allUniqSizes, counter, bar);

                counter++;
                ProgressCounter++;
            }
        }

        private void FormatDataWorksheet(Worksheet worksheet)
        {
            //Шапка таблицы
            worksheet.Range["A1"].Value = "Шифр объекта:";
            worksheet.Range["A2"].Value = "Наименование объекта:";
            worksheet.Range["A3"].Value = "Наименование здания:";
            worksheet.Range[CODE + 4].Value        = "Код изделия";
            worksheet.Range[MARK + 4].Value        = "Марка (Поз.)";
            worksheet.Range[TYPE + 4].Value        = "Тип изделия";
            worksheet.Range[UNITS + 4].Value       = "Ед. измерения";
            worksheet.Range[COUNT + 4].Value       = "Кол.";
            worksheet.Range[MASS + 4].Value        = "Масса ед., кг";
            worksheet.Range[DIAMETER + 4].Value    = "Диаметр, мм";
            worksheet.Range[CLASS + 4].Value       = "Класс";
            worksheet.Range[LENGTH + 4].Value      = "Длина, мм";
            worksheet.Range[ITEMCOUNT + 4].Value   = "Кол. заготовок";
            worksheet.Range[ITEMMASS + 4].Value    = "Масса заготовки, кг";
            worksheet.Range[MASSTOTAL + 4].Value   = "Всего, кг";
            worksheet.Range[COUNSTRTYPE + 4].Value = "Тип основы";
            worksheet.Range[CONSTRMARK + 4].Value  = "Метка основы";
            worksheet.Range[LEVEL + 4].Value       = "Этаж";
            worksheet.Range[DRAWING + 4].Value     = "Чертеж (ссылка)";
            worksheet.Range[IMAGE + 4].Value       = "Эскиз\n(размеры даны по наружным габаритам)";
            ProgressCounter++;

            //Удаление префикса "_" у заголовка размеров
            foreach (Range cell in worksheet.UsedRange.Rows[4].Columns)
            {
                cell.Value = cell.Value.TrimStart('_');
            }
            ProgressCounter++;

            //Выравнивание текста
            worksheet.UsedRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            worksheet.UsedRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Range["A1"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["A2"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["A3"].HorizontalAlignment = XlHAlign.xlHAlignRight;
            worksheet.Range["B1:B3"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
            worksheet.Range["A4"].EntireRow.WrapText = true;
            ProgressCounter++;

            //Габариты ячеек
            worksheet.Range[CODE + 4].ColumnWidth = 35;
            worksheet.Range[MARK + 4].ColumnWidth = 14;
            worksheet.Range[TYPE + 4].ColumnWidth = 25;
            worksheet.Range[UNITS + 4].ColumnWidth = 12;
            worksheet.Range[COUNT + 4].ColumnWidth = 10;
            worksheet.Range[MASS + 4].ColumnWidth = 10;
            worksheet.Range[DIAMETER + 4].ColumnWidth = 10;
            worksheet.Range[CLASS + 4].ColumnWidth = 10;
            worksheet.Range[LENGTH + 4].ColumnWidth = 10;
            worksheet.Range[ITEMCOUNT + 4].ColumnWidth = 10;
            worksheet.Range[ITEMMASS + 4].ColumnWidth = 10;
            worksheet.Range[MASSTOTAL + 4].ColumnWidth = 10;
            worksheet.Range[COUNSTRTYPE + 4].ColumnWidth = 14;
            worksheet.Range[CONSTRMARK + 4].ColumnWidth = 14;
            worksheet.Range[LEVEL + 4].ColumnWidth = 20;
            worksheet.Range[DRAWING + 4].ColumnWidth = 25;
            worksheet.Range[IMAGE + 4].ColumnWidth = 38;
            worksheet.Range["A4"].RowHeight = 50;
            ProgressCounter++;

            foreach (Range row in worksheet.UsedRange.Rows)
            {
                if (row.Height < 30) row.RowHeight = 30;
            }
            ProgressCounter++;

            //Формат границ ячеек
            worksheet.UsedRange.Borders.Weight = XlBorderWeight.xlThin;
            worksheet.Range["A1:A3"].EntireRow.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            worksheet.Range["A3"].Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            worksheet.UsedRange.Rows[4].Columns.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            ProgressCounter++;

            //Цвет фона ячеек
            worksheet.UsedRange.Rows[4].Columns.Interior.Color = headerСolor;
            ProgressCounter++;

            //Стиль текста
            worksheet.Range["A1"].Font.Bold = true;
            worksheet.Range["A2"].Font.Bold = true;
            worksheet.Range["A3"].Font.Bold = true;
            worksheet.Range["A4"].EntireRow.Font.Bold = true;
            ProgressCounter++;

            //Закрепление шапки
            worksheet.Activate();
            worksheet.Application.ActiveWindow.SplitRow = 4;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            //Масштаб листа
            worksheet.Application.ActiveWindow.Zoom = 75;
            ProgressCounter++;
        }

        private void SaveFile(IExcel.Application app, Workbook workbook)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog()
            {
                DefaultExt = "xlsx",
                AddExtension = true,
                Filter = "Книга Excel |*.xlsx",
                FileName = ProjectData.FileName,
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

        private void OpenCreatedFile(string path)
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

        private string GetRebarAssemblyCode(RebarAssembly rebarAssembly)
        {
            switch (rebarAssembly.ConstructionTypeEnum)
            {
                case ConstructionType.Beam:
                    return string.Join("/", ProjectData.ProjectCode, "Бм", rebarAssembly.Mark);
                case ConstructionType.Column:
                    return string.Join("/", ProjectData.ProjectCode, "Км", rebarAssembly.Mark);
                case ConstructionType.Floor:
                    return string.Join("/", ProjectData.ProjectCode, "Пм", rebarAssembly.Mark);
                case ConstructionType.Foundation:
                    return string.Join("/", ProjectData.ProjectCode, "Фм", rebarAssembly.Mark);
                case ConstructionType.Wall:
                    return string.Join("/", ProjectData.ProjectCode, "СТм", rebarAssembly.Mark);
                default:
                    return string.Join("/", ProjectData.ProjectCode, string.Empty, rebarAssembly.Mark);
            }
        }

        private IList<string> GetAllUniqSizes(IList<Bar> bars)
        {
            return bars.SelectMany(bar => bar.Sizes.Select(s => s))
                .Where(size => size.Value != 0)
                .Select(size => size.Key.Name)
                .Distinct()
                .OrderBy(size => size)
                .ToList();
        }

        private void SetSizeHeader(Worksheet worksheet, IList<string> allUniqSizes)
        {
            var uniqSizeStartCell = worksheet.Range["R4"];

            foreach (var uniqSize in allUniqSizes)
            {
                uniqSizeStartCell.Value = uniqSize;
                uniqSizeStartCell = uniqSizeStartCell.Next;
            }
        }

        private void SetShapeImage(Worksheet worksheet, Bar bar, int counter)
        {
            var barImageShapeRange = worksheet.Range["Q" + (5 + counter)];
            float Left = (float)(double)barImageShapeRange.Left + 1;
            float Top = (float)(double)barImageShapeRange.Top + 1;

            worksheet.Shapes.AddPicture(bar.ShapeImagePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, -1, -1);

            barImageShapeRange.RowHeight = 62;
        }

        private void SetSizesInCells(Worksheet worksheet, IList<string> allUniqSizes, int counter, Bar bar)
        {
            if (bar.Shape != "0.1" && bar.Shape != "0.2")
            {
                var currentCell = worksheet.Range["R" + (5 + counter)];

                foreach (var uniqSize in allUniqSizes)
                {
                    var sizeValue = bar.Sizes.Where(size => size.Key.Name.Equals(uniqSize)).First().Value;
                    if (sizeValue != 0)
                    {
                        currentCell.Value = sizeValue;
                    }

                    currentCell = currentCell.Next;
                }
            }
        }

        private string GetBarType(Bar bar)
        {
            if (bar.Shape == "0.1" || bar.Shape == "0.2")
                return "Мерная";

            return "Гнутая";
        }
    }
}