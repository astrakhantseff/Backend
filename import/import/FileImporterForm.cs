using ClosedXML.Excel;
using DevExpress.XtraSplashScreen;
using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using FileImport.Models;
using System.Text;
using FileImport.Pages;

namespace FileImport
{
    public partial class FileImporterForm : Form
    {
        internal string SavePath => GetSavePath();
        internal string FileName { get; set; }

        internal Company CompanyInfo { get; set; } = new Company();
        internal Dictionary<int, Field> FieldsInfo { get; set; } = new Dictionary<int, Field>();
        internal Dictionary<int, Protocol> ProtocolsInfo { get; set; } = new Dictionary<int, Protocol>();
        internal Dictionary<int, Reserve> ReservesInfo { get; set; } = new Dictionary<int, Reserve>();

        public FileImporterForm()
        {
            InitializeComponent();

            tabPage1.Text = ConfigurationManager.AppSettings["Page1"];
            tabPage2.Text = ConfigurationManager.AppSettings["Page2"];
            tabPage3.Text = ConfigurationManager.AppSettings["Page3"];
            tabPage4.Text = ConfigurationManager.AppSettings["Page4"];
            tabPage5.Text = ConfigurationManager.AppSettings["Page5"];
            tabPage6.Text = ConfigurationManager.AppSettings["Page6"];
            tabPage7.Text = ConfigurationManager.AppSettings["Page7"];

            //listBox1.Items.Add(@"C:\Новая папка\Пример отчета 6ГР Ванадий.xlsx");
            ////listBox1.Items.Add(@"C:\Новая папка\Пример отчета 6ГР Ванадий.xlsx");

            //listBox2.Items.Add(@"C:\Новая папка\6-ГР_конденсат_ПНГ_01.01.2023.xlsx");
            ////listBox2.Items.Add(@"C:\Новая папка\6-ГР_конденсат_ПНГ_01.01.2023.xlsx");

            //listBox3.Items.Add(@"C:\Новая папка\6-ГР_нефть_ПНГ_01.01.2023.xlsx");
            ////listBox3.Items.Add(@"C:\Новая папка\6-ГР_нефть_ПНГ_01.01.2023.xlsx");

            //listBox4.Items.Add(@"C:\Новая папка\6-ГР_ресурсы категории Д0_ПНГ_01 01 2023.xlsx");
            ////listBox4.Items.Add(@"C:\Новая папка\6-ГР_ресурсы категории Д0_ПНГ_01 01 2023.xlsx");

            //listBox5.Items.Add(@"C:\Новая папка\Книга 2. 6-гр нефть на 01.01.2024.xlsx");
            ////listBox5.Items.Add(@"C:\Новая папка\Книга 2. 6-гр нефть на 01.01.2024.xlsx");

            //listBox6.Items.Add(@"C:\Новая папка\Книга 6 D0 на 01.01.2024.xlsx");
            //listBox6.Items.Add(@"C:\Новая папка\Книга 6 D0 на 01.01.2024.xlsx");

            listBox7.Items.Add(@"C:\Новая папка\Книга 3. 6-гр газ, гелий на 01.01.2024.xlsx");
            //listBox7.Items.Add(@"C:\Новая папка\Книга 3. 6-гр газ, гелий на 01.01.2024.xlsx");
        }

        private IOverlaySplashScreenHandle ShowProgressPanel()
        {
            // Создаем экземпляр OverlayWindowOptions с заданными параметрами
            var options = new OverlayWindowOptions(
                startupDelay: 500,              // Уменьшаем задержку для более быстрого отображения панели
                backColor: Color.LightGray,     // Легкий цвет фона для более приятного визуального восприятия
                opacity: 0.7,                   // Более высокая прозрачность для улучшенной видимости элементов на фоне
                fadeIn: true,                   // Активируем плавное появление для более плавного UX
                fadeOut: true,                  // Активируем плавное затухание для завершения UX
                imageSize: new Size(64, 64));   // Размер изображения остается прежним

            // Возвращаем управление отображением OverlayForm с заданными параметрами
            return SplashScreenManager.ShowOverlayForm(this, options);
        }

        private void CloseProgressPanel(IOverlaySplashScreenHandle handle)
        {
            if (handle != null)
                SplashScreenManager.CloseOverlayForm(handle);
        }

        private ListBox GetTargetListBox()
        {
            // Словарь для сопоставления вкладок с соответствующими ListBox
            var tabToListBoxMap = new Dictionary<TabPage, ListBox>
            {
                { tabPage1, listBox1 },
                { tabPage2, listBox2 },
                { tabPage3, listBox3 },
                { tabPage4, listBox4 },
                { tabPage5, listBox5 },
                { tabPage6, listBox6 },
                { tabPage7, listBox7 }
            };

            // Получение текущей выбранной вкладки
            var selectedTab = tabControl.SelectedTab;

            // Попытка найти ListBox, соответствующий выбранной вкладке
            tabToListBoxMap.TryGetValue(selectedTab, out var targetListBox);

            return targetListBox;
        }

        private static string GetSavePath()
        {
            // Попробуем получить путь из настроек конфигурации.
            string path = ConfigurationManager.AppSettings["SavePath"];

            // Если путь пустой или отсутствует, возвращаем текущий рабочий каталог.
            if (string.IsNullOrEmpty(path))
            {
                return Directory.GetCurrentDirectory();
            }

            // Возвращаем найденный путь.
            return path;
        }

        public void AddButton_Click(object sender, EventArgs e)
        {
            // Создание OpenFileDialog с возможностью множественного выбора файлов
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xlsx;*.xlsm"
            };

            // Если пользователь выбрал файлы, добавляем их в ListBox
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ListBox targetListBox = GetTargetListBox();

                // Если целевой список определён, добавляем новые файлы
                if (targetListBox != null)
                {
                    foreach (var file in openFileDialog.FileNames)
                    {
                        // Проверяем наличие файла в списке; если файл уже есть, пропускаем его
                        if (!targetListBox.Items.Contains(file))
                        {
                            targetListBox.Items.Add(file);
                        }
                    }
                }
            }
        }

        private async void CopyButton_Click(object sender, EventArgs e)
        {
            // Массив для хранения всех списков файлов
            var listBoxes = new[] { listBox1, listBox2, listBox3, listBox4, listBox5, listBox6, listBox7 };

            // Массив для хранения PageNum
            var pageNumbers = new[] { PageNum.Number1, PageNum.Number2, PageNum.Number3, PageNum.Number4, PageNum.Number5, PageNum.Number6, PageNum.Number7 };

            // Проверка на наличие выбранных файлов
            if (listBoxes.All(listBox => listBox.Items.Count == 0))
            {
                MessageBox.Show("Сначала выберите файлы для копирования.", "Нет файлов", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Создаем загрузочный экран
            IOverlaySplashScreenHandle handle = null;

            try
            {
                handle = ShowProgressPanel();

                outputListBox.Items.Clear();

                var template = new Template(this);
                // Асинхронное создание копии шаблона
                await template.CreateTemplateCopyAsync();

                string path = SavePath;

                // Метод для обработки файлов в переданном списке
                async Task ProcessFilesAsync(ListBox listBox, PageNum pageNumber)
                {
                    foreach (var file in listBox.Items)
                    {
                        try
                        {
                            // Получение имени файла
                            string fileName = file.ToString();
                            var num = (int)pageNumber;

                            string dictionary = ConfigurationManager.AppSettings[$"Sheet{num}_Source3"];
                            string source = ConfigurationManager.AppSettings[$"Sheet{num}_Source2"];

                            DictionaryValues dictionaryValues = new DictionaryValues
                            {
                                Index = ConfigurationManager.AppSettings[$"Sheet{num}_FieldIndex"],
                                FieldName = ConfigurationManager.AppSettings[$"Sheet{num}_FieldName"],
                                RegionOfRussia = ConfigurationManager.AppSettings[$"Sheet{num}_RegionOfRussia"],
                                DevelopmentStage = ConfigurationManager.AppSettings[$"Sheet{num}_DevelopmentStage"]
                            };

                            // Асинхронная заполнение информации о компании
                            await template.WriteCompanyInfoAsync(fileName, pageNumber);

                            // В зависимости от типа операции, выполнение дополнительных действий
                            switch (num)
                            {
                                case 2:
                                case 3:
                                    {
                                        var page2_3 = new Page2_3(this);
                                        // Асинхронное заполнение информации о полях
                                        await page2_3.FillFieldsInfoAsync(fileName, dictionary, source, dictionaryValues);
                                        // Асинхронное заполнение информации о протоколах
                                        await page2_3.FillProtocolInfoAsync(fileName, source, num);

                                        await template.WriteAsync();
                                        break;
                                    }
                                case 4:
                                    {
                                        var page4 = new Page4(this);
                                        // Асинхронное заполнение информации о полях
                                        await page4.FillFieldsInfoAsync(fileName, dictionary, source, dictionaryValues);

                                        await template.WriteAsync();
                                        break;
                                    }
                                case 5:
                                    {
                                        var page5 = new Page5(this);
                                        // Асинхронное заполнение информации о полях
                                        await page5.FillFieldsInfoAsync(fileName, source);
                                        // Асинхронное заполнение информации о запасах
                                        await page5.FillReserves(fileName, source);

                                        await template.WriteAsync();
                                        break;
                                    }

                                case 6:
                                    {
                                        var page6 = new Page6(this);
                                        // Асинхронное заполнение информации о полях
                                        await page6.FillFieldsInfoAsync(fileName, source);
                                        
                                        await template.WriteAsync();
                                        break;
                                    }
                                case 7:
                                    {
                                        var page7 = new Page7(this);
                                        // Асинхронное заполнение информации о полях
                                        await page7.FillFieldsInfoAsync(fileName, source);
                                        // Асинхронное заполнение информации о запасах
                                        await page7.FillReserves(fileName, source);
                                        
                                        await template.WriteAsync();
                                        break;
                                    }
                            }

                            outputListBox.Items.Add($"Ok: {file} скопирован в {Path.Combine(path, FileName)}.");
                        }
                        catch (Exception xcp)
                        {
                            outputListBox.Items.Add($"Ошибка: {xcp.Message}");
                        }
                        finally
                        {
                            UpdateScrollBars();
                        }
                    }
                }

                // Асинхронная обработка каждого списка файлов с использованием цикла
                for (int i = 0; i < listBoxes.Length; i++)
                {
                    // Обработка i-го списка файлов
                    await ProcessFilesAsync(listBoxes[i], pageNumbers[i]);
                }
            }
            finally
            {
                Template.ResetRows();
                await SaveWithBorderAsync();

                // Закрываем загрузочный экран
                CloseProgressPanel(handle);
            }
        }

        internal bool LicExist(License license)
        {
            return FieldsInfo.Values
                .SelectMany(field => field.Licenses)
                .Any(lic => lic.LicenseNumber == license.LicenseNumber);
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            ListBox targetListBox = GetTargetListBox();

            if (targetListBox != null)
            {
                // Получаем массив индексов выбранных элементов в ListBox.
                var selectedIndices = targetListBox.SelectedIndices;

                // Удаляем элементы с конца, чтобы избежать изменения индексов других элементов.
                for (int i = selectedIndices.Count - 1; i >= 0; i--)
                {
                    targetListBox.Items.RemoveAt(selectedIndices[i]);
                }
            }
        }

        private void DeleteListBox_KeyDown(object sender, KeyEventArgs e)
        {
            // Проверяем, нажата ли клавиша Delete
            if (e.KeyCode == Keys.Delete)
            {
                DeleteButton_Click(sender, e);
            }
        }

        private void outputListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            // Проверяем, что индекс элемента находится в допустимом диапазоне
            if (e.Index < 0 || e.Index >= outputListBox.Items.Count)
                return;

            // Получаем элемент
            string itemText = outputListBox.Items[e.Index].ToString();

            // Определяем условия для изменения цвета заливки
            bool isError = itemText.StartsWith("Ошибка: ");
            bool isOk = itemText.StartsWith("Ok: ");

            // Предварительное определение цветов
            Color backgroundColor = outputListBox.BackColor;
            Color textColor = outputListBox.ForeColor;

            // Изменяем цвета в зависимости от содержания текста
            if (isError)
            {
                backgroundColor = Color.Tomato;
            }
            else if (isOk)
            {
                backgroundColor = Color.LightGreen;
            }

            // Отрисовка фонового цвета и настройка цвета текста
            using (SolidBrush backgroundBrush = new SolidBrush(backgroundColor))
            using (SolidBrush textBrush = new SolidBrush(textColor))
            {
                // Устанавливаем цвет фона
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);

                // Устанавливаем цвет текста
                e.Graphics.DrawString(itemText, e.Font, textBrush, e.Bounds, StringFormat.GenericDefault);
            }

            // Рисуем фокус, если элемент выделен
            e.DrawFocusRectangle();
        }

        private void UpdateScrollBars()
        {
            // Рассчет максимальной ширины текста элементов для горизонтальной полосы прокрутки
            int maxWidth = 0;
            using (Graphics g = outputListBox.CreateGraphics())
            {
                foreach (string item in outputListBox.Items)
                {
                    int itemWidth = (int)g.MeasureString(item, outputListBox.Font).Width;
                    if (itemWidth > maxWidth)
                    {
                        maxWidth = itemWidth;
                    }
                }
            }

            // Условие для горизонтальной полосы прокрутки
            outputListBox.HorizontalExtent = maxWidth;
        }

        private void FileImporterForm_Resize(object sender, EventArgs e)
        {
            // Обновление полосы прокрутки при изменении размера формы
            UpdateScrollBars();
        }

        private async Task SaveWithBorderAsync()
        {
            // Получаем названия листов из конфигурации
            var destinations = new List<string>
            {
                ConfigurationManager.AppSettings["Destination1"],
                ConfigurationManager.AppSettings["Destination2"],
                ConfigurationManager.AppSettings["Destination3"],
                ConfigurationManager.AppSettings["Destination4"],
                ConfigurationManager.AppSettings["Destination5"],
                ConfigurationManager.AppSettings["Destination6"],
                ConfigurationManager.AppSettings["Destination7"]
            };

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                // Открываем существующую рабочую книгу
                using (var workbook = new XLWorkbook(FileName))
                {
                    // Определяем функцию для обработки листа
                    void ProcessWorksheet(IXLWorksheet worksheet)
                    {
                        // Получаем диапазон занятых ячеек на листе
                        var usedRange = worksheet.RangeUsed();

                        // Проверяем, что диапазон не пустой
                        if (usedRange != null)
                        {
                            // Устанавливаем границы
                            usedRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            usedRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            usedRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            usedRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            usedRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                            // Центрируем текст по горизонтали
                            usedRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        }
                    }

                    // Проходим по всем листам из конфигурации
                    foreach (var destination in destinations)
                    {
                        // Пытаемся получить рабочий лист и обработать его
                        if (workbook.TryGetWorksheet(destination, out var worksheet))
                        {
                            ProcessWorksheet(worksheet);
                        }
                        else
                        {
                            throw new Exception($"Лист с именем '{destination}' не найден в '{FileName}'.");
                        }
                    }

                    // Сохраняем изменения в книге
                    workbook.Save();
                }
            });
        }
    }
}

