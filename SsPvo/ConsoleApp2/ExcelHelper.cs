using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;

namespace ConsoleApp2
{
    public static class ExcelHelper
    {
        public static XlsxAppInfo[] GetAppsFromExcel(string file, Func<ExcelWorksheet, int, Dictionary<string, int>, XlsxAppInfo> mapping)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(file)))
            {
                string sheetName = "Список заявлений";
                var sheet = excelPackage.Workbook.Worksheets[sheetName];
                if (sheet == null)
                    throw new ArgumentException($"{nameof(file)}: в файле \"{file}\" не найден лист \"{sheetName}\"");

                int numRows = sheet.Dimension.Rows;
                int numCols = sheet.Dimension.Columns;

                var headerToColumnIndex = new Dictionary<string, int>();

                for (int i = 1; i <= numCols; i++)
                {
                    headerToColumnIndex.Add(sheet.Cells[1, i].Value?.ToString(), i);
                }

                var data = new List<XlsxAppInfo>();

                for (int r = 2; r <= numRows; r++)
                {
                    data.Add(mapping(sheet, r, headerToColumnIndex));
                }

                return data.ToArray();
            }
        }

        public static SnilsWithEpguIds[] GetSnilsWithAppUidsFromXlsFile(string file)
        {
            var data = GetAppsFromExcel(file, XlsxAppInfo.FromExcelRowMin);
            return data.GroupBy(x => x.Snils)
                .Select(gr =>
                {
                    return new SnilsWithEpguIds
                    {
                        Snils = gr.Key,
                        EpguIds = gr.Select(x => x.EpguId).ToArray()
                    };
                })
                .ToArray();
        }

        public struct XlsxAppInfo
        {
            [DisplayName("Номер заявления")]
            public string AppNum { get; set; }
            [DisplayName("Фамилия")]
            public string LastName { get; set; }
            [DisplayName("Имя")]
            public string FirstName { get; set; }
            [DisplayName("Отчество")]
            public string MiddleName { get; set; }
            [DisplayName("Снилс")]
            public string Snils { get; set; }
            [DisplayName("Конкурс")]
            public string CgName { get; set; }
            [DisplayName("Уровень")]
            public string Level { get; set; }
            [DisplayName("Форма")]
            public string Form { get; set; }
            [DisplayName("Источник финанс.")]
            public string FinSource { get; set; }
            [DisplayName("Статус")]
            public string EpguStatus { get; set; }
            [DisplayName("Дата регистр.")]
            public string RegDate { get; set; }
            [DisplayName("Дата изменения")]
            public string LastModDate { get; set; }
            [DisplayName("Оригинал док. об образ-нии")]
            public string IsOriginal { get; set; }
            [DisplayName("Согласие на зачисл.")]
            public string IsAgreed { get; set; }
            [DisplayName("Дата согласия")]
            public string IsAgreedDate { get; set; }
            [DisplayName("Отзыв согласия на зачисл.")]
            public string IsRevoked { get; set; }
            [DisplayName("Дата отзыва согласия")]
            public string IsRevokedDate { get; set; }
            [DisplayName("Общежитие")]
            public string NeedHostel { get; set; }
            [DisplayName("Рейтинг")]
            public string Rating { get; set; }
            [DisplayName("ЕПГУ")]
            public string EpguId { get; set; }

            public Dictionary<string, string> OtherColumns { get; set; }

            public static XlsxAppInfo FromExcelRow(ExcelWorksheet ws, int r, Dictionary<string, int> columnMap)
            {
                return new XlsxAppInfo
                {
                    AppNum = $"{ws.Cells[r, columnMap["Номер заявления"]].Value}",
                    LastName = $"{ws.Cells[r, columnMap["Фамилия"]].Value}",
                    FirstName = $"{ws.Cells[r, columnMap["Имя"]].Value}",
                    MiddleName = $"{ws.Cells[r, columnMap["Отчество"]].Value}",
                    Snils = $"{ws.Cells[r, columnMap["Снилс"]].Value}",
                    CgName = $"{ws.Cells[r, columnMap["Конкурс"]].Value}",
                    Level = $"{ws.Cells[r, columnMap["Уровень"]].Value}",
                    Form = $"{ws.Cells[r, columnMap["Форма"]].Value}",
                    FinSource = $"{ws.Cells[r, columnMap["Источник финанс."]].Value}",
                    EpguStatus = $"{ws.Cells[r, columnMap["Статус"]].Value}",
                    RegDate = $"{ws.Cells[r, columnMap["Дата регистр."]].Value}",
                    LastModDate = $"{ws.Cells[r, columnMap["Дата изменения"]].Value}",
                    IsOriginal = $"{ws.Cells[r, columnMap["Оригинал док. об образ-нии"]].Value}",
                    IsAgreed = $"{ws.Cells[r, columnMap["Согласие на зачисл."]].Value}",
                    IsAgreedDate = $"{ws.Cells[r, columnMap["Дата согласия"]].Value}",
                    IsRevoked = $"{ws.Cells[r, columnMap["Отзыв согласия на зачисл."]].Value}",
                    IsRevokedDate = $"{ws.Cells[r, columnMap["Дата отзыва согласия"]].Value}",
                    NeedHostel = $"{ws.Cells[r, columnMap["Общежитие"]].Value}",
                    Rating = $"{ws.Cells[r, columnMap["Рейтинг"]].Value}",
                    EpguId = $"{ws.Cells[r, columnMap["ЕПГУ"]].Value}"
                };
            }

            public static XlsxAppInfo FromExcelRowMin(ExcelWorksheet ws, int r, Dictionary<string, int> columnMap)
            {
                return new XlsxAppInfo
                {
                    AppNum = $"{ws.Cells[r, columnMap["Номер заявления"]].Value}",
                    Snils = $"{ws.Cells[r, columnMap["Снилс"]].Value}",
                    EpguId = $"{ws.Cells[r, columnMap["ЕПГУ"]].Value}"
                };
            }
        }

        public struct SnilsWithEpguIds
        {
            public string Snils { get; set; }
            public string[] EpguIds { get; set; }
        }
    }
}
