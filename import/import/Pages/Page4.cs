using ClosedXML.Excel;
using FileImport.Models;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FileImport.Pages
{
    internal class Page4
    {
        private readonly FileImporterForm _init;

        internal Page4(FileImporterForm init)
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

                        _init.FieldsInfo.Add(i++, field);
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
                                    string value = nextCell.GetValue<string>();
                                    if (value != _init.FieldsInfo[key].FieldName)
                                        continue;

                                    nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 1, cell.Address.ColumnNumber + 1);
                                    _init.FieldsInfo[key].FieldType = nextCell.GetValue<string>();

                                    nextCell = cell.Worksheet.Cell(cell.Address.RowNumber + 2, cell.Address.ColumnNumber + 1);
                                    value = nextCell.GetValue<string>();
                                    string[] arr = value.Split(' ');

                                    var license = new License()
                                    {
                                        LicenseNumber = arr[0],
                                        RegistrationDate = arr[2]
                                    };
                                    license.Parse();

                                    if (!_init.LicExist(license))
                                    {
                                        _init.FieldsInfo[key].Licenses.Add(license);
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
    }
}
