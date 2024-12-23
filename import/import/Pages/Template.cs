using ClosedXML.Excel;
using FuzzyString;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileImport
{
    internal class Template
    {
        private readonly FileImporterForm _init;
        
        private static int RowCompany { get; set; } = 7;
        private static int RowFields { get; set; } = 7;
        private static int RowLicenses { get; set; } = 7;
        private static int RowProtocols { get; set; } = 7;
        private static int RowReserves { get; set; } = 7;
        private static int RowDeposits { get; set; } = 7;
        private static int RowCharacteristics { get; set; } = 8;

        internal Template(FileImporterForm init)
        {
            _init = init;
        }

        internal async Task CreateTemplateCopyAsync()
        {
            // Создаем уникальное имя для файла с использованием текущей даты и времени
            _init.FileName = string.Format(ConfigurationManager.AppSettings["FileName"], DateTime.UtcNow.ToString("yyyy-MM-dd_HH-mm-ss"));

            // Объединяем путь сохранения с именем файла
            _init.FileName = Path.Combine(_init.SavePath, _init.FileName);

            string template = ConfigurationManager.AppSettings["Template"];
            string destination = ConfigurationManager.AppSettings["Destination1"];

            // Используем асинхронный метод для загрузки и сохранения Excel файла
            await Task.Run(() =>
            {
                // Открываем шаблонный Excel файл
                using (var workbook = new XLWorkbook(template))
                {
                    // Проверяем, существует ли рабочий лист с заданным именем
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        // Если рабочий лист не найден, сообщаем об этом
                        throw new Exception($"Лист с именем '{destination}' не найден в '{template}'.");
                    }

                    // Сохраняем копию рабочей книги асинхронно
                    workbook.SaveAs(_init.FileName);
                }
            });
        }

        private async Task ReadAsync(string file, PageNum pageNumber)
        {
            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            var num = (int)pageNumber;

            // Получение конфигурационных значений для заданного листа
            var configKeys = new[] { "Source", "CompanyName", "PostalAddress", "Inn", "Okpo", "Oktmo", "Kpp" };
            var configValues = configKeys.ToDictionary(key => key, key => ConfigurationManager.AppSettings[$"Sheet{num}_{key}"]);

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверяем, существует ли рабочий лист с заданным именем
                    var worksheet = workbook.Worksheets.FirstOrDefault(sheet =>
                        sheet.Name == configValues["Source"] || sheet.Name.Contains(configValues["Source"]));

                    if (worksheet == null)
                    {
                        throw new Exception($"Лист с именем '{configValues["Source"]}' не найден в '{file}'.");
                    }

                    // Считывание и присваивание данных, если конфигурационное значение не пустое
                    void SetCompanyInfoIfNotEmpty(string key, Action<string> setValue)
                    {
                        if (!string.IsNullOrWhiteSpace(configValues[key]))
                        {
                            setValue(worksheet.Cell(configValues[key]).GetString());
                        }
                    }

                    SetCompanyInfoIfNotEmpty("CompanyName", value => _init.CompanyInfo.CompanyName = value);
                    SetCompanyInfoIfNotEmpty("PostalAddress", value => _init.CompanyInfo.PostalAddress = value);
                    SetCompanyInfoIfNotEmpty("Inn", value => _init.CompanyInfo.Inn = value);
                    SetCompanyInfoIfNotEmpty("Okpo", value => _init.CompanyInfo.Okpo = value);
                    SetCompanyInfoIfNotEmpty("Oktmo", value => _init.CompanyInfo.Oktmo = value);
                    SetCompanyInfoIfNotEmpty("Kpp", value => _init.CompanyInfo.Kpp = value);
                }
            });
        }

        private async Task WriteAsync(int row)
        {
            // Извлекаем путь назначения из конфигурации
            string destination = ConfigurationManager.AppSettings["Destination1"];

            // Используем асинхронный метод для загрузки и сохранения Excel файла
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    // Пытаемся получить лист с указанным именем
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        // Выбрасываем исключение, если листа с таким именем нет
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    // Записываем соответствующие данные в ячейки
                    worksheet.Cell($"A{row}").Value = _init.CompanyInfo.CompanyName;
                    worksheet.Cell($"C{row}").Value = _init.CompanyInfo.PostalAddress;
                    worksheet.Cell($"D{row}").Value = _init.CompanyInfo.Inn;
                    worksheet.Cell($"F{row}").Value = _init.CompanyInfo.Okpo;
                    worksheet.Cell($"I{row}").Value = _init.CompanyInfo.Oktmo;
                    worksheet.Cell($"H{row}").Value = _init.CompanyInfo.Kpp;

                    // Сохраняем изменения в файл
                    workbook.Save();
                }
            });
        }

        internal async Task WriteCompanyInfoAsync(string file, PageNum pageNumber)
        {
            // Асинхронно читаем данные из текущего файла
            await ReadAsync(file, pageNumber);

            // Пишем данные в общий файл, передавая текущий номер строки
            await WriteAsync(RowCompany);

            // Увеличиваем номер строки для следующего файла
            RowCompany++;
        }

        private async Task WriteFieldsInfoAsync()
        {
            // Получаем путь назначения из конфигурационного файла
            string destination = ConfigurationManager.AppSettings["Destination2"];

            // Используем асинхронный метод для загрузки и сохранения Excel файла
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                // Используем конструкцию 'using' для гарантированной очистки ресурсов
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    // Пробуем получить рабочий лист по указанному имени
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    // Используем StringBuilder для форматирования строковых значений,
                    // чтобы избежать многократного создания строк
                    var stringBuilder = new StringBuilder();

                    // Перебираем все значения из FieldsInfo и записываем их в соответствующие ячейки
                    foreach (var field in _init.FieldsInfo.Values)
                    {
                        // Форматируем адрес ячейки для каждой записи
                        stringBuilder.Clear();
                        stringBuilder.AppendFormat("A{0}", RowFields);
                        worksheet.Cell(stringBuilder.ToString()).Value = field.FieldName;

                        stringBuilder.Clear();
                        stringBuilder.AppendFormat("B{0}", RowFields);
                        worksheet.Cell(stringBuilder.ToString()).Value = field.DevelopmentStage;

                        stringBuilder.Clear();
                        stringBuilder.AppendFormat("C{0}", RowFields);
                        worksheet.Cell(stringBuilder.ToString()).Value = field.FieldType;

                        stringBuilder.Clear();
                        stringBuilder.AppendFormat("D{0}", RowFields);
                        worksheet.Cell(stringBuilder.ToString()).Value = field.RegionOfRussia;

                        stringBuilder.Clear();
                        stringBuilder.AppendFormat("E{0}", RowFields);
                        worksheet.Cell(stringBuilder.ToString()).Value = field.Location;

                        // Увеличиваем счетчик строки для следующей записи
                        RowFields++;
                    }

                    // Сохраняем изменения в файл
                    workbook.Save();
                }
            });
        }

        private async Task WriteLicenseInfoAsync()
        {
            string destination = ConfigurationManager.AppSettings["Destination3"];

            // Используем асинхронный метод для загрузки и сохранения Excel файла
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    foreach (var field in _init.FieldsInfo.Values)
                    {
                        // Обрабатываем каждый LicenseNumber во всех полях field
                        foreach (var license in field.Licenses)
                        {
                            worksheet.Cell($"A{RowLicenses}").Value = license.LicenseNumber;
                            worksheet.Cell($"B{RowLicenses}").Value = license.Series;
                            worksheet.Cell($"C{RowLicenses}").Value = license.Number;
                            worksheet.Cell($"D{RowLicenses}").Value = license.Type;
                            worksheet.Cell($"H{RowLicenses}").Value = license.RegistrationDate;
                            worksheet.Cell($"F{RowLicenses}").Value = "Действующая";

                            RowLicenses++;
                        }
                    }

                    // Сохраняем изменения в файл
                    workbook.Save();
                }
            });
        }

        private async Task WriteProtocolInfoAsync()
        {
            string destination = ConfigurationManager.AppSettings["Destination4"];

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    foreach (var protocol in _init.ProtocolsInfo.Values)
                    {
                        worksheet.Cell($"A{RowProtocols}").Value = protocol.ProtocolNumber;
                        worksheet.Cell($"F{RowProtocols}").Value = protocol.ApprovingAuthority;
                        
                        switch (protocol.ApprovingAuthority)
                        {
                            case "ГКЗ":
                                worksheet.Cell($"C{RowProtocols}").Value = "Отчёт по подсчету(пересчету) запасов";
                                break;

                            case "ФАН":
                                worksheet.Cell($"C{RowProtocols}").Value = "Оперативный подсчет запасов";
                                break;
                        }

                        worksheet.Cell($"D{RowProtocols}").Value = protocol.ProtocolNumber;
                        worksheet.Cell($"B{RowProtocols}").Value = "Действующий";

                        RowProtocols++;
                    }

                    // Сохраняем изменения в файл асинхронно
                    workbook.Save();
                }
            });
        }

        internal async Task WriteReservesInfoAsync()
        {
            string destination = ConfigurationManager.AppSettings["Destination5"];

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    foreach (var reserve in _init.ReservesInfo)
                    {
                        foreach (var category in reserve.Value.Categories)
                        {
                            worksheet.Cell($"A{RowReserves}").Value = reserve.Key + 1;
                            worksheet.Cell($"B{RowReserves}").Value = reserve.Value.Ouz;
                            worksheet.Cell($"C{RowReserves}").Value = reserve.Value.Layer;
                            worksheet.Cell($"D{RowReserves}").Value = reserve.Value.Field;
                            worksheet.Cell($"E{RowReserves}").Value = reserve.Value.Year;
                            worksheet.Cell($"F{RowReserves}").Value = reserve.Value.MineralComponent;

                            worksheet.Cell($"G{RowReserves}").Value = category.Name;
                            worksheet.Cell($"H{RowReserves}").Value = category.Geological;
                            worksheet.Cell($"I{RowReserves}").Value = category.Recoverable;
                            worksheet.Cell($"J{RowReserves}").Value = category.Production;

                            RowReserves++;
                        }
                    }

                    // Сохраняем изменения в файл асинхронно
                    workbook.Save();
                }
            });
        }

        internal async Task WriteDepositInfoAsync()
        {
            string destination = ConfigurationManager.AppSettings["Destination6"];

            var stratigraphies = await GetDictStratigraphyListAsync();

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    foreach (var reserve in _init.ReservesInfo)
                    {
                        foreach (var category in reserve.Value.Categories)
                        {
                            worksheet.Cell($"A{RowDeposits}").Value = reserve.Key + 1;
                            worksheet.Cell($"B{RowDeposits}").Value = reserve.Value.Ouz;
                            worksheet.Cell($"C{RowDeposits}").Value = "Пласт в целом";
                            worksheet.Cell($"D{RowDeposits}").Value = reserve.Value.Layer;
                            worksheet.Cell($"E{RowDeposits}").Value = reserve.Value.Field;
                            worksheet.Cell($"F{RowDeposits}").Value = "Разрабатываемое";
                            worksheet.Cell($"J{RowDeposits}").Value = "Утвержден";

                            var field = _init.FieldsInfo.FirstOrDefault(f => f.Value.FieldName == reserve.Value.Field);
                            if (field.Value != null)
                            {
                                worksheet.Cell($"G{RowDeposits}").Value = field.Value.FieldType;
                                worksheet.Cell($"H{RowDeposits}").Value = field.Value.RegionOfRussia;

                                var license = field.Value.License?.LicenseNumber;
                                worksheet.Cell($"M{RowDeposits}").Value = license ?? string.Empty;

                                var protocolNumber = field.Value.ProtocolNumber;
                                worksheet.Cell($"S{RowDeposits}").Value = protocolNumber ?? string.Empty;
                            }

                            worksheet.Cell($"K{RowDeposits}").Value = _init.CompanyInfo.CompanyName;
                            worksheet.Cell($"I{RowDeposits}").Value = reserve.Value.Year;
                            worksheet.Cell($"N{RowDeposits}").Value = reserve.Value.DepositType;

                            worksheet.Cell($"O{RowDeposits}").Value = reserve.Value.DepositCollector;
                            //worksheet.Cell($"P{RowDeposits}").Value = reserve.Value.DepositName;

                            worksheet.Cell($"P{RowDeposits}").Value = reserve.Value.DepositName + ": " + Format(stratigraphies, reserve.Value.DepositName, reserve.Value.Ouz);

                            worksheet.Cell($"T{RowDeposits}").Value = reserve.Value.DepositMinDepth;
                            worksheet.Cell($"U{RowDeposits}").Value = reserve.Value.DepositMaxDepth;

                            worksheet.Cell($"Q{RowDeposits}").Value = "Распределенный фонд";
                            worksheet.Cell($"R{RowDeposits}").Value = 0;

                            RowDeposits++;
                        }
                    }

                    // Сохраняем изменения в файл асинхронно
                    workbook.Save();
                }
            });
        }

        private static string Format(List<string> stratigraphies, string depositName, string ouz)
        {
            // Цикл по каждому слову в словаре для поиска лучшего совпадения
            string bestMatch = null;
            double bestSimilarity = 0.0;

            double similarityThreshold = 0.8;

            foreach (var stratigraphy in stratigraphies)
            {
                //List<FuzzyStringComparisonOptions> options = new List<FuzzyStringComparisonOptions>();
                // Вычисляем коэффициент схожести с каждому словом из словаря
                //double currentScore = FuzzyString.ComparisonMetrics.JaccardIndex(depositName, stratigraphy);

                //options.Add(FuzzyStringComparisonOptions.UseOverlapCoefficient);
                //options.Add(FuzzyStringComparisonOptions.UseLongestCommonSubsequence);
                //options.Add(FuzzyStringComparisonOptions.UseLongestCommonSubstring);
                //options.Add(FuzzyStringComparisonOptions.UseJaccardDistance);

                //// Choose the relative strength of the comparison - is it almost exactly equal? or is it just close?
                //FuzzyStringComparisonTolerance tolerance = FuzzyStringComparisonTolerance.Normal;

                //var result = _bindingList.Where(x => x.Text.ApproximatelyEquals(txtSearch.Text, tolare, options.ToArray())).ToList();

                //bool similarity = depositName.ApproximatelyEquals(stratigraphy, FuzzyStringComparisonTolerance.Strong, FuzzyStringComparisonOptions.UseRatcliffObershelpSimilarity);
                //if (similarity)
                //    return stratigraphy;
                if (stratigraphy.Contains(ouz))
                    return stratigraphy;
                // Обновляем лучшее совпадение, если текущая схожесть лучше
                //if (similarity > bestSimilarity)
                //{
                //    bestSimilarity = similarity;
                //    bestMatch = stratigraphy;
                //}
            }
            return bestMatch;
        }

        internal async Task WriteAsync()
        {
            // Асинхронная запись информации о месторождении
            await WriteFieldsInfoAsync();

            // Асинхронная запись информации о лицензии
            await WriteLicenseInfoAsync();

            // Асинхронная запись информации о протоколе
            await WriteProtocolInfoAsync();

            // Асинхронная запись информации о запасах
            await WriteReservesInfoAsync();

            // Асинхронная запись информации о залежах
            await WriteDepositInfoAsync();

            // Асинхронная запись информации о залежах
            await WriteCharacteristicsInfoAsync();
        }

        private async Task WriteCharacteristicsInfoAsync()
        {
            string destination = ConfigurationManager.AppSettings["Destination7"];

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                // Открываем существующий шаблон
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    if (!workbook.TryGetWorksheet(destination, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{destination}' не найден в '{_init.FileName}'.");
                    }

                    foreach (var reserve in _init.ReservesInfo)
                    {
                        foreach (var category in reserve.Value.Categories)
                        {
                            worksheet.Cell($"G{RowCharacteristics}").Value = 1;
                            worksheet.Cell($"AN{RowCharacteristics}").Value = 0;
                            worksheet.Cell($"BP{RowCharacteristics}").Value = 0;

                            worksheet.Cell($"A{RowCharacteristics}").Value = reserve.Key + 1;
                            worksheet.Cell($"B{RowCharacteristics}").Value = reserve.Value.Ouz;
                            worksheet.Cell($"C{RowCharacteristics}").Value = reserve.Value.Layer;
                            worksheet.Cell($"D{RowCharacteristics}").Value = reserve.Value.Field;
                            worksheet.Cell($"E{RowCharacteristics}").Value = reserve.Value.Year;
                            worksheet.Cell($"F{RowCharacteristics}").Value = category.Name;

                            worksheet.Cell($"AH{RowCharacteristics}").Value = reserve.Value.DepositMinDepth;
                            worksheet.Cell($"AI{RowCharacteristics}").Value = reserve.Value.DepositMaxDepth;
                            worksheet.Cell($"AK{RowCharacteristics}").Value = reserve.Value.DepositMinAbsDepth;
                            worksheet.Cell($"AL{RowCharacteristics}").Value = reserve.Value.DepositMaxAbsDepth;

                            worksheet.Cell($"R{RowCharacteristics}").Value = 0;

                            RowCharacteristics++;
                        }
                    }

                    // Сохраняем изменения в файл асинхронно
                    workbook.Save();
                }
            });
        }

        private async Task<List<string>> GetDictStratigraphyListAsync()
        {
            var result = new List<string>();

            // Проверка существования файла
            if (!File.Exists(_init.FileName))
            {
                throw new FileNotFoundException($"Файл не найден {_init.FileName}.");
            }

            string source = ConfigurationManager.AppSettings[$"Destination8"];

            // Используем асинхронный метод для длительных операций
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(_init.FileName))
                {
                    // Проверяем, что нужный лист присутствует
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{_init.FileName}'.");
                    }

                    // Получаем используемые ячейки в колонке A
                    var column = worksheet.Column("A");

                    // Поиск информации о поле
                    foreach (var cell in column.CellsUsed().Skip(1))
                    {
                        string cellValue = cell.GetValue<string>();
                        result.Add(cellValue);
                    }
                }
            });

            return result; // Возвращаем список найденных значений
        }

        internal static void ResetRows()
        {
            RowCompany = 7;
            RowFields = 7;
            RowLicenses = 7;
            RowProtocols = 7;
            RowReserves = 7;
            RowDeposits = 7;
            RowCharacteristics = 8;
        }
    }
}