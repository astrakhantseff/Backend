using ClosedXML.Excel;
using FileImport.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FileImport.Pages
{
    internal class Page5
    {
        private readonly FileImporterForm _init;

        internal PageType PageType { get; set; } = PageType.Oil;

        internal Page5(FileImporterForm init)
        {
            _init = init;
        }

        internal async Task FillFieldsInfoAsync(string file, string source)
        {
            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    _init.FieldsInfo.Clear();
                    _init.ProtocolsInfo.Clear();

                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    var column = worksheet.Column("A");

                    var developmentStage = string.Empty;

                    switch (PageType)
                    {
                        case PageType.Oil:
                            developmentStage = worksheet.Cell("A7").GetString();
                            developmentStage = char.ToUpper(developmentStage[0]) + developmentStage.ToLower().Substring(1);
                            break;

                        case PageType.Gaz:
                            developmentStage = "Разрабатываемое";
                            break;
                    }

                    var regexPattern1 = @"^\d+\s+";
                    var regexPattern2 = @"\d+\s+(.+?),(\s+(.{3})\s+(.+?)\s+(.{2})\s+(\d{2}.\d{2}.\d{4}))?(\s+({(.+?)}))?(\s+\/(.+?)\/)?(\s+<Протокол\s+№\s+(.+)\s+(.{3})\s+(.+)>)?";
                    var regexYearPattern = @"а\)\s+(\d+)\s+б\)\s+(\d+)";

                    int fieldKey = 0;
                    int protocolKey = 0;

                    foreach (var cell in column.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();

                        if (cellValue.ToLower().Trim() == "разведываемые месторождения".ToLower())
                            break;

                        if (!Regex.IsMatch(cellValue, regexPattern1))
                            continue;

                        Match match = Regex.Match(cellValue, regexPattern2);
                        if (match.Success)
                        {
                            Process(match, fieldKey, protocolKey);
                            protocolKey++;
                        }
                        else
                        {
                            throw new Exception($"Pattern {nameof(regexPattern2)} not match: row {cellValue}");
                        }

                        if (_init.FieldsInfo.Keys.Last() == fieldKey)
                        {
                            var nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 1, cell.Address.ColumnNumber);
                            ProcessYear(nextCell, fieldKey, regexYearPattern);

                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 2, cell.Address.ColumnNumber);
                            _init.FieldsInfo[fieldKey].RegionOfRussia = nextCell.GetValue<string>();
                            _init.FieldsInfo[fieldKey].DevelopmentStage = developmentStage;
                        }
                        else
                        {
                            throw new Exception("Last key is supposed to be equal another value");
                        }

                        fieldKey++;
                    }
                }
            });
        }

        private void Process(Match match, int fieldKey, int protocolKey)
        {
            var groups = match.Groups;

            var field = new Field
            {
                FieldName = groups[1].Value,
                Location = groups[9].Value,
                FieldType = groups[11].Value
            };

            var license = new License
            {
                LicenseNumber = groups[3].Value + groups[4].Value + groups[5].Value,
                Series = groups[3].Value,
                Number = groups[4].Value,
                Type = groups[5].Value,
                RegistrationDate = groups[6].Value
            };

            if (!string.IsNullOrWhiteSpace(license.LicenseNumber))
            {
                field.License = license;
                field.Licenses.Add(license);
            }

            _init.FieldsInfo.Add(fieldKey, field);

            var protocol = new Protocol
            {
                ProtocolNumber = groups[13].Value,
                ApprovingAuthority = groups[14].Value
            };

            if (!string.IsNullOrWhiteSpace(protocol.ProtocolNumber))
            {
                field.ProtocolNumber = protocol.ProtocolNumber;
                _init.ProtocolsInfo.Add(protocolKey, protocol);
            }
        }

        private void ProcessYear(IXLCell nextCell, int fieldKey, string pattern)
        {
            string value = nextCell.GetValue<string>();
            Match match = Regex.Match(value, pattern);
            if (match.Success)
            {
                _init.FieldsInfo[fieldKey].DiscoveryYear = match.Groups[1].Value;
                _init.FieldsInfo[fieldKey].DevelopmentStartYear = match.Groups[2].Value;
            }
        }

        internal async Task<List<(string id, string fieldName)>> FillFieldIdReservesAsync(string file, string source)
        {
            var fieldIdList = new List<(string FieldName, string Id)>();

            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            // Используем асинхронный метод для длительных операций
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверяем, что нужный лист присутствует
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    // Получаем используемые ячейки в колонке A
                    var column = worksheet.Column("A");

                    // Регулярное выражение для поиска идентификатора и имени поля
                    var regexPattern1 = @"^(\d+)\s+(.+),";

                    // Поиск информации о поле
                    foreach (var cell in column.CellsUsed())
                    {
                        string cellValue = cell.GetValue<string>();
                        if (cellValue.ToLower().Trim() == "разведываемые месторождения".ToLower())
                            break;

                        if (!Regex.IsMatch(cellValue, regexPattern1))
                            continue;

                        Match match = Regex.Match(cellValue, regexPattern1);
                        if (match.Success)
                        {
                            var fieldId = match.Groups[1].Value;
                            var fieldName = match.Groups[2].Value;

                            // Добавляем найденные данные в список
                            fieldIdList.Add((fieldId, fieldName));
                        }
                    }
                }
            });

            return fieldIdList; // Возвращаем список найденных значений
        }

        private bool ContainsOnlyLatinLetters(string input)
        {
            // Проверяем каждый символ в строке
            foreach (char c in input)
            {
                // Если символ не является латинской буквой, возвращаем false
                if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || char.IsDigit(c)))
                {
                    return false;
                }
            }
            // Если все символы - латинские буквы, возвращаем true
            return true;
        }

        internal async Task FillReserves(string file, string source)
        {
            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            _init.ReservesInfo.Clear();

            var listFields = await FillFieldIdReservesAsync(file, source);
            var year = await GetYearReservesAsync(file);
            
            // Используем асинхронный метод для долгих операций
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверяем, что нужный лист присутствует
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    // Получаем используемые ячейки в колонке A
                    var column = worksheet.Column("A");

                    int reserveKey = 0;

                    // Поиск информации о поле
                    foreach (var item in listFields)
                    {
                        (string fieldId, string fieldName) = item;

                        // Регулярное выражение для дальнейшего анализа
                        string regexPattern = $"^{fieldId}" + @"[.]\d+[.]\d+([.]\d+)?";

                        // Поиск и обработка резервов
                        foreach (var cell in column.CellsUsed())
                        {
                            string cellValue = cell.GetValue<string>();
                            if (cellValue.ToLower().Trim() == "разведываемые месторождения".ToLower())
                                break;

                            if (!Regex.IsMatch(cellValue, regexPattern))
                                continue;

                            Match match = Regex.Match(cellValue, regexPattern);
                            if (!match.Success)
                            {
                                continue;
                            }

                            string s = cellValue.Replace(match.Groups[0].Value, string.Empty);
                            var arr = s.Split(new[] { ',', ' ', '.' }, StringSplitOptions.RemoveEmptyEntries);
                            if (arr.Length == 0 || !ContainsOnlyLatinLetters(arr[0]))
                            {
                                continue;
                            }

                            Reserve reserve = new Reserve
                            {
                                Ouz = arr[0],
                                Layer = arr[0],
                                MineralComponent = "Нефть",
                                Field = fieldName,
                                Year = year
                            };

                            // Загрузка следующей ячейки для сдвига на 1 строки вниз от текущей
                            var nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 1, cell.Address.ColumnNumber);
                            cellValue = nextCell.GetValue<string>();
                            reserve.DepositName = GetDepositName(cellValue);

                            // Загрузка следующей ячейки для сдвига на 2 строки вниз от текущей
                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 2, cell.Address.ColumnNumber);
                            cellValue = nextCell.GetValue<string>();

                            var typeAndCollector = GetTypeAndCollector(cellValue);
                            reserve.DepositType = Format(typeAndCollector.Item1);
                            reserve.DepositCollector = typeAndCollector.Item2;

                            var depth = GetDepthNum(cellValue);
                            //reserve.DepositDepth = depositDepth;
                            reserve.DepositMinDepth = depth.Item1;
                            reserve.DepositMaxDepth = depth.Item3;
                            reserve.DepositMinAbsDepth = depth.Item3;
                            reserve.DepositMaxAbsDepth = depth.Item3;

                            // Загрузка следующей ячейки для сдвига на 3 строки вниз от текущей
                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 3, cell.Address.ColumnNumber);
                            var depositYear = nextCell.GetValue<string>();

                            if (PageType == PageType.Oil)
                            {
                                // Загрузка следующей ячейки для сдвига на 4 строки вниз от текущей
                                nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 4, cell.Address.ColumnNumber);
                                cellValue = nextCell.GetValue<string>();
                            }

                            var pattern = @"^[a-zA-Z](\d*)(\+[a-zA-Z]\d*)*$";
                            int shift = 0;

                            string geological = string.Empty;
                            string recoverable = string.Empty;
                            string production = string.Empty;

                            while (!Regex.IsMatch(cellValue, regexPattern) &&
                                   !string.IsNullOrWhiteSpace(cellValue))
                            {
                                nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 4 + shift, cell.Address.ColumnNumber);
                                cellValue = nextCell.GetValue<string>();

                                if (Regex.IsMatch(cellValue, pattern))
                                {
                                    switch (PageType)
                                    {
                                        case PageType.Oil:
                                            // Извлечение данных о категории
                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 4 + shift, cell.Address.ColumnNumber + 1);
                                            geological = nextCell.GetValue<string>();

                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 5 + shift, cell.Address.ColumnNumber + 1);
                                            recoverable = nextCell.GetValue<string>();

                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 6 + shift, cell.Address.ColumnNumber + 1);
                                            production = nextCell.GetValue<string>();
                                            break;

                                        case PageType.Gaz:
                                            // Извлечение данных о категории
                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 4 + shift, cell.Address.ColumnNumber + 1);
                                            geological = nextCell.GetValue<string>();

                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 5 + shift, cell.Address.ColumnNumber + 1);
                                            recoverable = nextCell.GetValue<string>();

                                            nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 6 + shift, cell.Address.ColumnNumber + 1);
                                            production = nextCell.GetValue<string>();
                                            break;
                                    }

                                    // Добавление новой категории в резерв
                                    var category = new Category()
                                    {
                                        Name = cellValue,
                                        Geological = geological,
                                        Recoverable = recoverable,
                                        Production = production
                                    };

                                    reserve.Categories.Add(category);
                                    shift += 3;
                                }
                                else
                                {
                                    shift++;
                                }
                            }

                            _init.ReservesInfo.Add(reserveKey++, reserve);
                        }
                    }
                }
            });

            string GetDepositName(string input)
            {
                //Метод IndexOf находит индекс первого пробела в строке
                int spaceIndex = input.IndexOf(' ');

                // Если пробел найден (index != -1), извлекаем подстроку от начала до пробела
                // Метод Substring извлекает подстроку, начиная с индекса 0 и до индекса пробела
                string result = input.Substring(spaceIndex + 1);

                return result;
            }

            string Format(string input)
            {
                string result = input;

                if (result.EndsWith("ая"))
                {
                    result = result.Replace("ая", "ое");
                }

                return result;
            }

            (string, string) GetTypeAndCollector(string input)
            {
                (string, string) result = (string.Empty, string.Empty);

                const string pattern = @"Залежь\s+:\s+(.+),\s+коллектор\s+:\s+(.+), Глубина залегания:.+";

                // Создание объекта для поиска совпадений в тексте
                Match match = Regex.Match(input, pattern);

                if (match.Success)
                {
                    result.Item1 = match.Groups[1].Value;
                    result.Item2 = match.Groups[2].Value;
                }

                return result;
            }

            (string, string, string, string) GetDepthNum(string input)
            {
                (string, string, string, string) result = (string.Empty, string.Empty, string.Empty, string.Empty);

                const string pattern = @"(\d+)\s+[(](-(\d+))[)]\s+-\s+(\d+)\s+[(](-(\d+))[)]";

                // Создание объекта для поиска совпадений в тексте
                Match match = Regex.Match(input, pattern);

                if (match.Success)
                {
                    // Извлечение чисел, найденных регулярным выражением
                    result.Item1 = match.Groups[1].Value;
                    result.Item2 = match.Groups[3].Value;
                    result.Item3 = match.Groups[4].Value;
                    result.Item4 = match.Groups[6].Value;
                }

                return result;
            }
        }

        private async Task<string> GetYearReservesAsync(string file)
        {
            string result = string.Empty;
            string source = string.Empty;

            // Проверка существования файла
            if (!File.Exists(file))
            {
                throw new FileNotFoundException($"Файл не найден {file}.");
            }

            int num = 0;

            switch (PageType)
            {
                case PageType.Oil:
                    num = 5;
                    break;

                case PageType.Gaz:
                    num = 7;
                    break;
            }

            if (num != 0)
            {
                source = ConfigurationManager.AppSettings[$"Sheet{num}_Source"];
            }
            else
            {
                return result;
            }

            // Используем асинхронный метод для длительных операций
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверяем, что нужный лист присутствует
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    var value = worksheet.Cell("E6").Value.ToString();

                    const string pattern = @"(\d+)";

                    Match match = Regex.Match(value, pattern);
                    if (match.Success)
                    {
                        result = match.Groups[0].Value;
                    }
                }
            });

            return result; // Возвращаем список найденных значений
        }
    }
}
