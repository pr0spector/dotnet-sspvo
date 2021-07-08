using Microsoft.Extensions.Logging;
using SsPvo.Client;
using SsPvo.Client.Enums;
using SsPvo.Client.Messages.Serialization;
using SsPvo.Client.Models;
using SsPvo.Ui.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;
using SsPvo.Client.Messages;

namespace SsPvo.Ui.Common.Batch.Handlers
{
    public class ExcelImportScenarioHandler : IBatchScenarioHandler
    {
        #region fields
        // private readonly ConcurrentQueue<SsPvoMessage> _msgQueue;
        private Dictionary<string, MsgWrapper> templates = new Dictionary<string, MsgWrapper>
        {
            {
                "serviceEntrant",
                new MsgWrapper
                {
                    Action = "get",
                    EntityType = "serviceEntrant",
                    Xml = "<PackageData><ServiceEntrant><IDEntrantChoice><SNILS>$arg0</SNILS></IDEntrantChoice></ServiceEntrant></PackageData>"
                }
            },
            {
                "document",
                new MsgWrapper
                {
                    Action = "get",
                    EntityType = "document",
                    Xml = "<PackageData><Document><IDEntrantChoice><SNILS>$arg0</SNILS></IDEntrantChoice><IDDocChoice><UIDEpgu>$arg1</UIDEpgu></IDDocChoice></Document></PackageData>"
                }
            }
        };
        #endregion

        #region ctor
        public ExcelImportScenarioHandler(SsPvoApiClient apiClient)
        {
            ApiClient = apiClient;
        }
        #endregion

        #region props
        public SsPvoApiClient ApiClient { get; set; }
        public BatchAction.Scenario Scenario => BatchAction.Scenario.ExcelImport;
        #endregion

        #region methods
        public async Task<BatchAction.Status> ProcessItemAsync(BatchAction.Item item, BatchAction.Options options, CancellationToken token)
        {
            if (ApiClient == null) return item.SetResult(BatchAction.Status.Error, "Клиент API недоступен!");

            var logger = options.GetValueOrDefault<ILogger<BatchAction>>($"{BatchAction.Options.CommonOptions.Log}");

            logger?.LogDebug($"Абитуриент {item.Description}..");

            var pe = item.GetProcessedEntity<SsAppFromExcel>();

            // TODO: token.ThrowIfCancellationRequested();

            try
            {
                // 1.1 запрос профиля по СНИЛС
                var t1 = templates["serviceEntrant"];

                uint? t1IdJwt = await RequestDataIdJwtAsync(
                    t1.Action, t1.EntityType, t1.Xml.Replace("$arg0", pe.Snils), logger, token);

                if (t1IdJwt == null)
                {
                    return item.SetResult(BatchAction.Status.Error,
                        $"Студент {pe.FullName}[СНИЛС:{pe.Snils}] ошибка получения idJwt для запроса профиля!");
                }

                // 1.2. запрос профиля по idJwt
                var xDocProfile = await RequestActualData((uint)t1IdJwt, logger, token);

                if (xDocProfile == null)
                {
                    return item.SetResult(BatchAction.Status.Error,
                        $"Студент {pe.FullName}[СНИЛС:{pe.Snils}] ошибка получения профиля абитуриента по idJwt {t1IdJwt}!");
                }

                // 1.3. сохранение xml-файла
                SaveXDocument(xDocProfile, $"{pe.Snils}.xml", logger);

                // 1.4. подтверждение
                if (Settings.Default.AutoConfirmBatchMessages)
                {
                    await ApiClient.SendQueueConfirmMessage((uint)t1IdJwt, token);
                }

                // запросы дополнительных сведений
                await RequestDocuments(pe, xDocProfile.XPathSelectElements("//Documents//UIDEpgu"), logger, token);

                return item.SetResult(BatchAction.Status.Completed,
                    $"Студент {pe.FullName}[СНИЛС:{pe.Snils}] обработка успешно завершена!");
            }
            catch (OperationCanceledException e)
            {
                return item.SetResult(BatchAction.Status.Canceled,
                    $"Студент {pe.FullName}[СНИЛС:{pe.Snils}] обработка прервана! {e.Message}");
            }
            catch (Exception e)
            {
                return item.SetResult(BatchAction.Status.Error, e.Message);
            }
        }

        private async Task<uint?> RequestDataIdJwtAsync(
            string dataRequestAction,
            string dataRequestEntityType,
            string dataRequestXml,
            ILogger<BatchAction> logger,
            CancellationToken token)
        {
            logger?.LogDebug($"Запрос {dataRequestEntityType}..");

            var msgServiceEntrantResponseData = await ApiClient.SendActionMessage(
                dataRequestAction, dataRequestEntityType, XDocument.Parse(dataRequestXml), token);

            logger?.LogDebug($"Ответ {dataRequestEntityType} получен. Извлекаем данные..");

            var idJwtResponse = ApiClient.TryExtractResponse<IdJwtResponse>(msgServiceEntrantResponseData);

            logger?.LogDebug($"idJwt: {idJwtResponse.Item1?.IdJwt}");

            if (string.IsNullOrWhiteSpace(idJwtResponse.Item1?.IdJwt)) return null;

            return uint.Parse(idJwtResponse.Item1.IdJwt);
        }

        private async Task<XDocument> RequestActualData(uint idJwt, ILogger<BatchAction> logger, CancellationToken token)
        {
            logger?.LogDebug($"Запрос данных по idJwt {idJwt} из service..");

            var msgServiceEntrantProfileResponseData =
                await ApiClient.SendGetQueueItemMessage(SsPvoQueue.Service, idJwt, token);

            logger?.LogDebug($"Ответ получен. Извлекаем данные..");

            var tokenResponse =
                ApiClient.TryExtractResponse<ResponseTokenResponse>(msgServiceEntrantProfileResponseData);

            if (string.IsNullOrWhiteSpace(tokenResponse.Item1?.ResponseToken)) return null;
            
            return new JwtToken(tokenResponse.Item1.ResponseToken).Decode().Item2;
        }

        private void SaveXDocument(XDocument xdoc, string fileName, ILogger<BatchAction> logger)
        {
            EnsureFolderExist();
            string fullPath = Path.GetFullPath(Path.Combine(Settings.Default.DefaultOutputFolder, fileName));
            logger?.LogDebug($"Сохраняем файл \"{fullPath}\"");
            xdoc.Save(fullPath);

        }

        private void EnsureFolderExist()
        {
            if (!Directory.Exists(Settings.Default.DefaultOutputFolder)) Directory.CreateDirectory(Settings.Default.DefaultOutputFolder);
        }


        private async Task RequestDocuments(SsAppFromExcel pe, IEnumerable<XElement> elements, ILogger<BatchAction> logger, CancellationToken token)
        {
            foreach (var el in elements)
            {
                logger?.LogDebug($"Загрузка документа UIDEpgu {el?.Value}");
            }
            return;

            foreach (var el in elements)
            {
                string uidEpgu = el.Value;

                logger?.LogDebug($"Загрузка документа UIDEpgu {uidEpgu}");

                var doc = templates["document"];

                uint? docIdJwt = await RequestDataIdJwtAsync(
                    doc.Action, doc.EntityType, doc.Xml.Replace("$arg0", pe.Snils).Replace($"$arg1", uidEpgu), logger, token);

                if (docIdJwt == null)
                {
                    logger?.LogWarning($"Студент {pe.FullName}[СНИЛС:{pe.Snils}] ошибка получения idJwt для запроса документа UIDEpgu {uidEpgu}!");
                }

                // 1.2. запрос документа по idJwt
                var xdoc = await RequestActualData((uint)docIdJwt, logger, token);

                if (xdoc == null)
                {
                    logger?.LogWarning($"Студент {pe.FullName}[СНИЛС:{pe.Snils}] ошибка получения данных по UIDEpgu {uidEpgu}!");
                }

                // 1.3. сохранение xml-файла
                SaveXDocument(xdoc, $"{pe.Snils}-document-{uidEpgu}.xml", logger);

                // 1.4. подтверждение
                if (Settings.Default.AutoConfirmBatchMessages)
                {
                    await ApiClient.SendQueueConfirmMessage((uint)docIdJwt, token);
                }
            }
        }

        private class MsgWrapper
        {
            public string Action { get; set; }
            public string EntityType { get; set; }
            public string Xml { get; set; }
        }

        #endregion
    }
}
