using ClosedXML.Excel;
using FileImport.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FileImport.Pages
{
    internal class Page6
    {
        private readonly FileImporterForm _init;

        internal Page6(FileImporterForm init)
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

            // Открытие и чтение из Excel файла в асинхронном контексте
            await Task.Run(() =>
            {
                _init.FieldsInfo.Clear();
                _init.ProtocolsInfo.Clear();

                // Используем переменную key для отслеживания ключей
                int key = 0;

                // Открываем workbook с использованием блока using
                using (var workbook = new XLWorkbook(file))
                {
                    // Проверяем наличие листа исходных данных
                    if (!workbook.TryGetWorksheet(source, out var worksheet))
                    {
                        throw new Exception($"Лист с именем '{source}' не найден в '{file}'.");
                    }

                    // Выбираем столбец 'A' для обработки
                    var column = worksheet.Column("A");

                    string region = worksheet.Cell("B9").GetString();

                    // Перебираем ячейки, исключая пустые
                    foreach (var cell in column.CellsUsed())
                    {
                        // Получаем строковое значение ячейки
                        string cellValue = cell.GetValue<string>();

                        // Проверяем, является ли значение числом
                        if (!cellValue.All(char.IsDigit))
                            continue;

                        // Получаем следующую ячейку справа
                        var nextCell = cell.Worksheet.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber + 1);

                        // Проверяем, если следующая ячейка тоже число, продолжаем
                        string val = nextCell.GetValue<string>();
                        if (val.All(char.IsDigit))
                            continue;

                        // Получаем значение следующей ячейки
                        string value = nextCell.GetValue<string>();

                        // Задаём шаблон для регулярного выражения
                        const string pattern = @"(.+)?,(\s+(.{3})\s+(.[^\s]+)\s+(.{2})\s+(от\s+)?(\d{2}.\d{2}.\d{4}))?";

                        // Выполняем сопоставление регулярного выражения
                        Match match = Regex.Match(value, pattern);
                        if (match.Success)
                        {
                            var groups = match.Groups;

                            // Инициализируем объект Field с именем месторождения
                            var field = new Field
                            {
                                FieldName = groups[1].Value,
                                RegionOfRussia = region
                            };

                            // Инициализируем объект License с параметрами лицензии
                            var license = new License
                            {
                                LicenseNumber = groups[3].Value + groups[4].Value + groups[5].Value,
                                Series = groups[3].Value,
                                Number = groups[4].Value,
                                Type = groups[5].Value,
                                RegistrationDate = groups[7].Value
                            };

                            // Проверяем наличие номера лицензии и добавляем её в список лицензий
                            if (!string.IsNullOrWhiteSpace(license.LicenseNumber))
                            {
                                field.Licenses.Add(license);
                            }

                            // Добавляем информацию о поле в список
                            _init.FieldsInfo.Add(key++, field);
                        }
                    }
                }
            });
        }
    }
}
