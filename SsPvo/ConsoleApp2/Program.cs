using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // api host
            string host = ConfigurationManager.AppSettings["host"];
            // ОГРН
            string ogrn = ConfigurationManager.AppSettings["ogrn"];
            // КПП
            string kpp = ConfigurationManager.AppSettings["kpp"];
            // путь сохранения xml/json
            string savepath = ConfigurationManager.AppSettings["savepath"];
            // x509 subject
            string certsubj = ConfigurationManager.AppSettings["certsubj"];
            // СНИЛС
            string snils = "";
            // файл xlsx с заявлениями
            string xlsxfile = "";

            // аргументы cmd перегружают параметры из конфига
            foreach (string arg in args)
            {
                if (arg.Contains("host=")) host = arg.Split('=')[1];
                if (arg.Contains("ogrn=")) ogrn = arg.Split('=')[1];
                if (arg.Contains("kpp=")) kpp = arg.Split('=')[1];
                if (arg.Contains("savepath=")) savepath = arg.Split('=')[1];
                if (arg.Contains("certsubj=")) certsubj = arg.Split('=')[1];
                if (arg.Contains("snils=")) snils = arg.Split('=')[1];
                if (arg.Contains("xlsxfile=")) xlsxfile = arg.Split('=')[1];
            }

            var apiClient = new SSClient(ogrn, kpp, host, new Crypto { X509SubjectFragment = certsubj }, savepath);

            Console.WriteLine("==================================");
            Console.WriteLine($"Текущие настройки:");
            Console.WriteLine($"{nameof(host)}:{host}");
            Console.WriteLine($"{nameof(ogrn)}:{ogrn}");
            Console.WriteLine($"{nameof(kpp)}:{kpp}");
            Console.WriteLine($"{nameof(savepath)}:{savepath}");
            Console.WriteLine($"{nameof(certsubj)}:{certsubj}");
            Console.WriteLine($"{nameof(snils)}:{snils}");
            Console.WriteLine($"{nameof(xlsxfile)}:{xlsxfile}");
            Console.WriteLine("==================================");
            Console.WriteLine();


                if (!string.IsNullOrWhiteSpace(xlsxfile))
                {
                    Console.WriteLine($"{nameof(xlsxfile)}: Указан файл с заявлениями, обработка..");

                    if (!xlsxfile.Contains(".xlsx"))
                    {
                        Console.WriteLine($"{nameof(xlsxfile)}: неподдерживаемыей формат файла!");
                        return;
                    }

                    Console.WriteLine($"Пробуем прочитать файл \"{xlsxfile}\"");

                    var excelData = ExcelHelper.GetSnilsWithAppUidsFromXlsFile(xlsxfile);
                    if (excelData == null || !excelData.Any())
                    {
                        throw new InvalidOperationException(
                            $"{nameof(xlsxfile)}: не удалось прочитать данные из \"{xlsxfile}\"");
                    }

                    foreach (var item in excelData)
                    {
                        await apiClient.saveXmlBySnils(item.Snils.Trim(),  item.EpguIds);
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(snils))
                    {
                        Console.WriteLine($"{nameof(snils)}: Не указан ни один СНИЛС");
                    }
                    else
                    {
                        if (snils.Contains(","))
                        {
                            Console.WriteLine($"{nameof(snils)}: Указано несколько СНИЛС, обработка..");
                            foreach (var curSnils in snils.Split(','))
                            {
                                await apiClient.saveXmlBySnils(curSnils.Trim());
                            }
                        }
                        else
                        {
                            Console.WriteLine($"{nameof(snils)}: Указан СНИЛС:{snils}, обработка..");
                            await apiClient.saveXmlBySnils(snils.Trim());
                        }
                    }
                }

#if DEBUG
            Console.WriteLine("Нажмите любую клавишу для завершения..");
            Console.ReadKey();
#endif
        }
    }
}
