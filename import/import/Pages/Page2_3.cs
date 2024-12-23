using ClosedXML.Excel;
using FileImport.Models;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FileImport.Pages
{
    internal class Page2_3
    {
        private FileImporterForm _init { get; set; }

        internal Page2_3(FileImporterForm init)
        {
            _init = init;
        }

        internal async Task FillFieldsInfoAsync(string file, string dictionary, string source, DictionaryValues dictionaryValues)
        {
            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверка существования листа
                    if (!workbook.TryGetWorksheet(dictionary, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{dictionary}' не найден в '{file}'.");
                    }

                    bool result = false;

                    _init.FieldsInfo.Clear();

                    int i = 1;

                    // Перебор всех ячеек на листе
                    foreach (var row in worksheet.RowsUsed().Skip(2))
                    {
                        var field = new Field
                        {
                            Index = row.Cell(dictionaryValues.Index).GetValue<string>(),
                            FieldName = row.Cell(dictionaryValues.FieldName).GetValue<string>(),
                            RegionOfRussia = row.Cell(dictionaryValues.RegionOfRussia).GetValue<string>(),
                            DevelopmentStage = row.Cell(dictionaryValues.DevelopmentStage).GetValue<string>()
                        };

                        result = !string.IsNullOrWhiteSpace(field.FieldName) &&
                                 !string.IsNullOrWhiteSpace(field.RegionOfRussia) &&
                                 !string.IsNullOrWhiteSpace(field.DevelopmentStage);

                        if (field.DevelopmentStage == "Разведываемые месторождения")
                            continue;

                        if (!result)
                            throw new Exception($"Ряд {dictionary}:{field.FieldName} содержит пустые ячейки в '{file}'.");

                        _init.FieldsInfo.Add(i, field);

                        i++;
                    }

                    // Заполняем тип месторождения
                    if (!workbook.TryGetWorksheet(source, out worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    // Выбираем столбец, с которым будем работать. Здесь берем столбец 'A'
                    var column = worksheet.Column("A");

                    foreach (var key in _init.FieldsInfo.Keys)
                    {
                        // Проходим по каждой ячейке в колонке, исключая пустые ячейки
                        foreach (var cell in column.CellsUsed())
                        {
                            // Получаем значение ячейки как строку
                            string cellValue = cell.GetValue<string>();

                            // Проверяем, является ли значение одной цифрой
                            if (cellValue.All(char.IsDigit))
                            {
                                // Выводим адрес ячейки и её значение
                                if (_init.FieldsInfo[key].Index == cellValue)
                                {
                                    // Получаем адрес следующей ячейки справа
                                    var nextCell = cell.Worksheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + 1);

                                    // Возвращаем значение из этой ячейки
                                    string val = nextCell.GetValue<string>();
                                    string[] arr = val.Split('\n');

                                    if (arr.Length > 2) // Проверяем, что есть хотя бы три элемента в массиве
                                    {
                                        // Устанавливаем FieldType и Location из второго элемента массива
                                        _init.FieldsInfo[key].FieldType = arr[1];
                                        _init.FieldsInfo[key].Location = arr.Length > 3 ? arr[3] : null;

                                        // Разбиваем третий элемент на части, используя запятую в качестве разделителей
                                        string[] arrLic = arr[2].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                        foreach (var lic in arrLic)
                                        {
                                            var parts = lic.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                            var license = new License()
                                            {
                                                LicenseNumber = parts[0],
                                                RegistrationDate = parts[2]
                                            };
                                            license.Parse();

                                            if (!_init.LicExist(license))
                                            {
                                                _init.FieldsInfo[key].Licenses.Add(license);
                                            }
                                        }
                                    }

                                    // Регулярное выражение для поиска чисел после 'а)' и 'б)'
                                    const string pattern = @"а\) (\d+); б\) (\d+)";

                                    // Создание объекта для поиска совпадений в тексте
                                    Match match = Regex.Match(val, pattern);

                                    if (match.Success)
                                    {
                                        // Извлечение чисел, найденных регулярным выражением
                                        _init.FieldsInfo[key].DiscoveryYear = match.Groups[1].Value;
                                        _init.FieldsInfo[key].DevelopmentStartYear = match.Groups[2].Value;
                                    }
                                }

                                if (_init.FieldsInfo[key].FieldType != null)
                                    break;
                            }
                        }
                    }
                }
            });
        }

        internal async Task FillProtocolInfoAsync(string file, string source, int pageNumber)
        {
            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            string destination = ConfigurationManager.AppSettings["Destination4"];

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    _init.ProtocolsInfo.Clear();

                    // Заполняем тип месторождения
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    // Выбираем столбец, с которым будем работать. Здесь берем столбец 'A'
                    var column = worksheet.Column("A");

                    int index = 1;
                    // Проходим по каждой ячейке в колонке, исключая пустые ячейки
                    foreach (var cell in column.CellsUsed().Skip(2))
                    {
                        // Получаем значение ячейки как строку
                        string cellValue = cell.GetValue<string>();

                        // Проверяем, содержит ли значение точку
                        if (cellValue.Contains("."))
                        {
                            // Получаем адрес ячейки справа того же ряда
                            var nextCell = cell.Worksheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + 19);

                            // Возвращаем значение из этой ячейки
                            string value = nextCell.GetValue<string>();

                            switch (pageNumber)
                            {
                                case 2:
                                    {
                                        // Регулярное выражение для поиска паттерна утверждающего органа и следующего за ним номера
                                        const string pattern = @"(.{3})\n.*\n\s*№\s*(.+)";

                                        // Создаем объект Regex с указанным паттерном
                                        Regex regex = new Regex(pattern);

                                        // Выполняем поиск всех совпадений в строке
                                        MatchCollection matches = regex.Matches(value);

                                        // Используем LINQ для обработки совпадений
                                        // Обрабатываем совпадения и получаем коллекцию протоколов
                                        var protocols = matches.Cast<Match>()
                                            .Where(match => match.Success)
                                            .Select(match => new Protocol
                                            {
                                                ProtocolNumber = match.Groups[2].Value,
                                                ApprovingAuthority = match.Groups[1].Value
                                            });

                                        // Добавляем найденные протоколы в список
                                        foreach (var protocol in protocols)
                                        {
                                            if (_init.ProtocolsInfo.Any(p => p.Value.ProtocolNumber == protocol.ProtocolNumber &&
                                                                       p.Value.ApprovingAuthority == protocol.ApprovingAuthority))
                                                continue;

                                            // Добавляем Protocol в словарь
                                            _init.ProtocolsInfo.Add(index++, protocol);
                                        }

                                        break;
                                    }

                                case 3:
                                    {
                                        // Определяем регулярное выражение с поиском даты
                                        const string pattern = @"№\s*(.+)\n(.{3})";

                                        // Создаем объект Regex для полного поиска
                                        Regex regex = new Regex(pattern);

                                        // Выполняем поиск всех совпадений в строке с заданным шаблоном
                                        MatchCollection matches = regex.Matches(value);

                                        // Перебор всех совпадений
                                        foreach (Match match in matches)
                                        {
                                            // Если совпадение успешно
                                            if (match.Success)
                                            {
                                                // Извлекаем группы совпадений
                                                var groups = match.Groups;

                                                // Инициализируем переменные без дублирования
                                                string number = groups[1].Value;
                                                string authority = groups[2].Value;

                                                // Создаем новый объект Protocol с извлеченными данными
                                                var protocol = new Protocol
                                                {
                                                    ProtocolNumber = number,
                                                    ApprovingAuthority = authority
                                                };

                                                if (_init.ProtocolsInfo.Any(p => p.Value.ProtocolNumber == protocol.ProtocolNumber &&
                                                                                 p.Value.ApprovingAuthority == protocol.ApprovingAuthority))
                                                    continue;

                                                // Добавляем Protocol в словарь
                                                _init.ProtocolsInfo.Add(index++, protocol);
                                            }
                                        }

                                        break;
                                    }
                            }
                        }
                    }
                }
            });
        }
    }
}
