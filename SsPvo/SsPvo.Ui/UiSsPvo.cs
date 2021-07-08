using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using SsPvo.Client;
using SsPvo.Client.Enums;
using SsPvo.Client.Extensions;
using SsPvo.Client.Messages;
using SsPvo.Client.Messages.Serialization;
using SsPvo.Ui.Common;
using SsPvo.Ui.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using SsPvo.Ui.Annotations;
using SsPvo.Ui.Common.Batch;
using SsPvo.Ui.Common.Batch.Handlers;
using SsPvo.Ui.Common.Logging;
using Formatting = Newtonsoft.Json.Formatting;
using SsPvo.Ui.Extensions;

namespace SsPvo.Ui
{
    public partial class UiSsPvo : Form, INotifyPropertyChanged
    {
        #region fields
        private readonly LoggerFactory _loggerFactory;
        private readonly Crypto _csp;
        private SsPvoApiClient _ssPvoClient;
        private BatchAction _importBatch;
        private ExcelImportScenarioHandler _jobHandler;
        private CancellationTokenSource _cancellationTokenSource;
        #endregion

        #region ctor
        public UiSsPvo(LoggerFactory loggerFactory)
        {
            _loggerFactory = loggerFactory;
            InitializeComponent();
            _csp = new Crypto();
            ExcelEntries = new List<SsAppFromExcel>();

            LogUtils.CustomLogEventSink.LogEvent += CustomLogEventSink_LogEvent;

            this.Load += UiEpguTests_Load;
        }
        #endregion

        #region props
        public SsPvoMessage CurrentMessage { get; set; }
        public List<SsAppFromExcel> ExcelEntries { get; set; }

        public SortableBindingList<BatchAction.Item> ImportItems { get; private set; }

        public bool AllowStartImport => _importBatch?.AllowStart ?? false;
        public bool AllowStopImport => _importBatch?.AllowStop ?? false;
        public bool AllowPauseImport => _importBatch?.AllowPause ?? false;
        public bool AllowResumeImport => _importBatch?.AllowResume ?? false;
        public bool AllowAddItemsImport => _importBatch?.AllowAddItems ?? false;
        public bool AllowRemoveItemsImport => _importBatch?.AllowRemoveItems ?? false;
        #endregion

        #region methods

        #region general
        public void InitBindings()
        {
            var bs = new BindingSource { DataSource = this };

            TsbRestart.DataBindings.Add(nameof(Control.Enabled), bs, nameof(AllowStartImport),
                true, DataSourceUpdateMode.Never);
            TsbCancel.DataBindings.Add(nameof(Control.Enabled), bs, nameof(AllowStopImport),
                true, DataSourceUpdateMode.Never);
            TsbRemoveSelectedItems.DataBindings.Add(nameof(Control.Enabled), bs, nameof(AllowRemoveItemsImport),
                true, DataSourceUpdateMode.Never);
            TsbClearItems.DataBindings.Add(nameof(Control.Enabled), bs, nameof(AllowRemoveItemsImport),
                true, DataSourceUpdateMode.Never);

            DgvImportFiles.DataSource = new BindingSource(bs, nameof(ImportItems));
        }
        private void ReInitClient()
        {
            _csp.X509SubjectFragment = TbSettingCertNameFragment.Text;

            _ssPvoClient = new SsPvoApiClient(
                Settings.Default.OGRN,
                Settings.Default.KPP,
                Settings.Default.SelectedApiUrl,
                _csp,
                _loggerFactory.CreateLogger<SsPvoApiClient>());

            if (_jobHandler != null) _jobHandler.ApiClient = _ssPvoClient;
        }
        private void SaveSettings()
        {
            // TODO: доработать, чтобы значения менялись автоматически
            Settings.Default.OGRN = TbSettingOGRN.Text;
            Settings.Default.KPP = TbSettingKPP.Text;
            Settings.Default.SelectedApiUrl = TbSettingSelectedApiUrl.Text;
            Settings.Default.LastImportedXlsxFile = TbApplicationsFile.Text;
            Settings.Default.LastEntityType = TbEntityType.Text;
            Settings.Default.AutoConfirmBatchMessages = ChkbSettingAutoConfirmBatchMessages.Checked;
            Settings.Default.DefaultInputFolder = TbSettingDefaultInputFolder.Text;
            Settings.Default.DefaultOutputFolder = TbSettingDefaultOutputFolder.Text;
            Settings.Default.LastIdJwt = TbIdJwt.Text;
            Settings.Default.CertNameFragment = TbSettingCertNameFragment.Text;

            Settings.Default.Save();
            Settings.Default.Reload();
        }

        private void CustomLogEventSink_LogEvent(object sender, Serilog.Events.LogEvent e)
        {
            try
            {
                TbExcelJobsLog.AppendText($@"{DateTime.Now:G} [{e.Level}]: {e.RenderMessage()}");
                TbExcelJobsLog.AppendText($"{Environment.NewLine}");
                Application.DoEvents();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
            }
        }

        public void IndicateState(bool busy, string message = "") =>
            this.IndicateState(busy, message, TssLabel, TssStatusIcon);
        #endregion

        #region handlers
        private void UiEpguTests_Load(object sender, EventArgs evts)
        {
            ReInitClient();

            _importBatch =
                new BatchAction(BatchAction.Scenario.ExcelImport, _loggerFactory.CreateLogger<BatchAction>());
            _jobHandler = new ExcelImportScenarioHandler(_ssPvoClient);
            _importBatch.Handlers.Add(_jobHandler);
            ImportItems = new SortableBindingList<BatchAction.Item>();

            TsbRestart.Click += async (s, e) => await RunJobsAsync();
            TsbCancel.Click += (s, e) => _cancellationTokenSource?.Cancel();
            TsbRemoveSelectedItems.Click += (s, e) =>
                RemoveImportItems(DgvImportFiles.SelectedRowsDataBoundItems<BatchAction.Item>());
            TsbClearItems.Click += (s, e) => RemoveImportItems(ImportItems);
            TsbCopyToClipboard.Click += (s, e) => Clipboard.SetText(TbExcelJobsLog.Text);
            TsbClearLog.Click += (s, e) => TbExcelJobsLog.Clear();

            CbMessageType.ValueMember = "Id";
            CbMessageType.DisplayMember = "Name";
            CbMessageType.DataSource = Enum.GetValues(typeof(SsPvoMessageType))
                .OfType<SsPvoMessageType>()
                .Select(x => new { Id = x, Name = $"{x.GetDescription()} [{x}]" })
                .ToList();
            CbCls.DataSource = Enum.GetValues(typeof(SsPvoCls));
            CbActionType.DataSource = Enum.GetValues(typeof(SsPvoAction));

            CbMessageType.SelectedValueChanged += CbMessageType_SelectedValueChanged;

            BtnSaveSettings.Click += (s, ev) =>
            {
                SaveSettings();
                ReInitClient();
            };

            BtnResetSettings.Click += (s, ev) =>
            {
                Settings.Default.Reset();
                Settings.Default.Save();
                Settings.Default.Reload();
                MessageBox.Show(@"Требуется перезапуск");
            };

            BtnSelectImportFile.Click += (s, ev) =>
            {
                string filePath = DialogHelper.OpenFile("(*.xlsx)|*.xlsx", initialDirectory: Settings.Default.DefaultOutputFolder) as string;
                TbApplicationsFile.Text = $@"{filePath}";
            };

            BtnCreateMessage.Click += (s, ev) => Create();
            BtnSendMessage.Click += async (s, ev) => await Send();
            BtnCreateAndSendMessage.Click += async (s, ev) => await CreateAndSend();

            DgvImportFiles.AutoGenerateColumns = false;
            DgvImportFiles.RowPostPaint += (s, e) => DgvImportFiles.PaintRowNumbers(this.Font, e);

            InitBindings();

            TsbImportEntriesFromExcelFile.Click += (s, e) => LoadApplicationsFile(TbApplicationsFile.Text);
        }
        private void CbMessageType_SelectedValueChanged(object sender, EventArgs e)
        {
            if (CbMessageType.SelectedValue == null || (SsPvoMessageType)CbMessageType.SelectedValue == SsPvoMessageType.Cert)
            {
                CbCls.Enabled = false;
                CbActionType.Enabled = false;
                TbEntityType.Enabled = false;
                TbIdJwt.Enabled = false;
                TbRequestPayload.Enabled = false;
                RbAllMessages.Enabled = false;
                RbConcreteMessage.Enabled = false;
            }
            else
            {
                switch ((SsPvoMessageType)CbMessageType.SelectedValue)
                {
                    case SsPvoMessageType.Cls:
                        CbCls.Enabled = true;
                        CbActionType.Enabled = false;
                        TbEntityType.Enabled = false;
                        TbIdJwt.Enabled = false;
                        TbRequestPayload.Enabled = false;
                        RbAllMessages.Enabled = false;
                        RbConcreteMessage.Enabled = false;
                        break;
                    case SsPvoMessageType.Action:
                        CbCls.Enabled = false;
                        CbActionType.Enabled = true;
                        TbEntityType.Enabled = true;
                        TbIdJwt.Enabled = true;
                        TbRequestPayload.Enabled = true;
                        RbAllMessages.Enabled = false;
                        RbConcreteMessage.Enabled = false;
                        break;
                    case SsPvoMessageType.ServiceQueue:
                    case SsPvoMessageType.EpguQueue:
                    case SsPvoMessageType.Confirm:
                        CbCls.Enabled = false;
                        CbActionType.Enabled = false;
                        TbEntityType.Enabled = false;
                        TbIdJwt.Enabled = true;
                        TbRequestPayload.Enabled = false;
                        RbAllMessages.Enabled = true;
                        RbConcreteMessage.Enabled = true;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }
        private void ImportBatchStatusChanged(object sender, BatchAction.Status e)
        {
            switch (e)
            {
                case BatchAction.Status.Completed:
                    this.IndicateState(false, "", TssLabel, TssStatusIcon);
                    break;
                case BatchAction.Status.NotStarted:
                case BatchAction.Status.Paused:
                case BatchAction.Status.Canceled:
                case BatchAction.Status.Error:
                    this.IndicateState(false, "", TssLabel, TssStatusIcon);
                    break;
                case BatchAction.Status.InProgress:
                    this.IndicateState(true, "", TssLabel, TssStatusIcon);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(e), e, null);
            }

            TsbRestart.Image = e == BatchAction.Status.NotStarted
                ? Properties.Resources.button_green_play_16x16
                : Properties.Resources.refresh_update;
        }
        private void RemoveImportItems(IEnumerable<BatchAction.Item> itemsToRemove)
        {
            if (!_importBatch.AllowRemoveItems) return;

            try
            {
                foreach (var item in itemsToRemove.ToList())
                {
                    _importBatch.Items.Remove(item);
                    ImportItems.Remove(item);
                }
            }
            catch (Exception e)
            {
                DialogHelper.ShowError(e.Message);
            }
            finally
            {
                OnPropertyChanged(nameof(ImportItems));
                OnPropertyChanged(nameof(AllowStartImport));
                OnPropertyChanged(nameof(AllowStopImport));
                OnPropertyChanged(nameof(AllowPauseImport));
                OnPropertyChanged(nameof(AllowResumeImport));
                OnPropertyChanged(nameof(AllowAddItemsImport));
                OnPropertyChanged(nameof(AllowRemoveItemsImport));
            }
        }
        #endregion


        #region manual messaging
        private void Create()
        {
            TbMessageGuid.Text = string.Empty;
            TbRequestPayload.Text = string.Empty;
            TbRequestHeader.Text = string.Empty;

            CurrentMessage = null;

            try
            {
                CurrentMessage = CreateMessageInternal();
                if (CurrentMessage == null) return;
                CurrentMessage.PrepareRequestData(_csp);

                TbMessageGuid.Text = $@"{CurrentMessage?.Guid}";
                TbRequestHeader.Text = Utils.GetSerialized(CurrentMessage.RequestData.JHeader);
                if (CurrentMessage?.RequestData?.XPayload != null)
                {
                    TbRequestPayload.Text = Utils.GetSerialized(CurrentMessage.RequestData.XPayload);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($@"{e.Message}", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private async Task Send()
        {
            TbResponseHeader.Text = string.Empty;
            TbResponsePayload.Text = string.Empty;
            TbResponseContent.Text = string.Empty;

            if (CurrentMessage == null)
            {
                MessageBox.Show($@"Нет сообщения для отправки!");
                return;
            }

            IndicateState(true, "Загрузка..");

            try
            {
                var response = await _ssPvoClient.SendMessage(CurrentMessage);
                DisplayResponse(response);
            }
            catch (Exception e)
            {
                MessageBox.Show($@"{e.Message}", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                IndicateState(false);
            }
        }
        private async Task CreateAndSend()
        {
            Create();
            await Send();
        }
        private SsPvoMessage CreateMessageInternal()
        {
            try
            {
                var msgType = (SsPvoMessageType)CbMessageType.SelectedValue;

                SsPvoMessage msg = null;
                var msgOptions = new SsPvoMessage.Options
                {
                    Type = msgType
                };

                switch (msgType)
                {
                    case SsPvoMessageType.Cls:
                        msgOptions.Cls = $"{(SsPvoCls)CbCls.SelectedValue}";
                        break;
                    case SsPvoMessageType.Cert:
                        break;
                    case SsPvoMessageType.Action:
                        msgOptions.Action = $"{(SsPvoAction)CbActionType.SelectedValue}".ToLowerFirstChar();
                        msgOptions.EntityType = TbEntityType.Text;
                        if (!string.IsNullOrWhiteSpace(TbRequestHeader.Text))
                        {
                            msgOptions.Payload = XDocument.Parse(TbRequestPayload.Text);
                        }
                        break;
                    case SsPvoMessageType.ServiceQueue:
                    case SsPvoMessageType.EpguQueue:
                        msgOptions.QueueMsgType = RbAllMessages.Checked
                            ? SsPvoQueueMsgSubType.AllMessages
                            : SsPvoQueueMsgSubType.SingleMessage;
                        if (msgOptions.QueueMsgType == SsPvoQueueMsgSubType.SingleMessage)
                        {
                            msgOptions.Action = "getMessage";
                            msgOptions.IdJwt = uint.Parse(TbIdJwt.Text);
                        }
                        break;
                    case SsPvoMessageType.Confirm:
                        msgOptions.IdJwt = uint.Parse(TbIdJwt.Text);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                return _ssPvoClient.MessageFactory.Create(msgOptions);
            }
            catch (Exception e)
            {
                TbResponseHeader.Text = e.Message;
                TbResponsePayload.Text = e.Message;
                TbResponseContent.Text = string.Empty;
            }

            return null;
        }

        private void DisplayResponse(ResponseData response)
        {
            TbResponseContent.Text = string.IsNullOrWhiteSpace(response?.Metadata?.Content)
                ? response?.Metadata?.ErrorMessage
                : response?.Metadata?.Content;

            /*
            switch (response.Message.MessageType)
            {
                case SsPvoMessageType.Cls:
                    var xdoc = await _ssPvoClient.GetDictionary((SsPvoCls)CbCls.SelectedValue);
                    TbResponseHeader.Text = string.Empty;
                    TbResponsePayload.Text = xdoc?.ToString(SaveOptions.None);
                    TbResponseContent.Text = string.Empty;
                    return;
                case SsPvoMessageType.Cert:
                    bool isRegistered = await _ssPvoClient.GetIsCertificateRegistered();
                    TbResponseHeader.Text = $@"{isRegistered}";
                    TbResponsePayload.Text = string.Empty;
                    TbResponseContent.Text = string.Empty;
                    break;
                case SsPvoMessageType.Action:
                    var idJwtResponse = await _ssPvoClient.SendAction(
                        (SsPvoAction)CbActionType.SelectedValue,
                        TbEntityType.Text,
                        XDocument.Parse(TbRequestPayload.Text));
                    if (idJwtResponse.Item1?.IdJwt != null)
                    {
                        TbResponseHeader.Text = idJwtResponse.Item1.IdJwt.ToString();
                        TbResponsePayload.Text = string.Empty;
                        TbResponseContent.Text = string.Empty;
                    }
                    else
                    {
                        TbResponseHeader.Text = idJwtResponse.Item2.Error;
                        TbResponsePayload.Text = idJwtResponse.Item2.Error;
                        TbResponseContent.Text = string.Empty;
                    }

                    break;
                case SsPvoMessageType.ServiceQueue:
                case SsPvoMessageType.EpguQueue:
                    var targetQueue = msgType == SsPvoMessageType.ServiceQueue
                        ? SsPvoQueue.Service
                        : SsPvoQueue.Epgu;
                    if (RbAllMessages.Checked)
                    {
                        var qmr = await _ssPvoClient.CheckQueueMessages(targetQueue);
                        if (qmr != null)
                        {
                            TbResponseHeader.Text =
                                $@"Messages: {qmr.Messages},{Environment.NewLine}IdJwts: [{(qmr.IdJwts != null ? string.Join(",", qmr.IdJwts) : null)}]";
                        }
                        else
                        {
                            TbResponseHeader.Text = @"Error!";
                        }
                        TbResponsePayload.Text = string.Empty;
                        TbResponseContent.Text = string.Empty;
                    }
                    else
                    {
                            //if (env.Metadata.IsSuccessful)
                            //{
                            //    var token = env.GetJwt("responseToken");
                            //    var decoded = token?.Decode();
                            //    if (decoded != null)
                            //    {
                            //        TbResponseHeader.Text = decoded.JHeader?.ToString();
                            //        TbResponsePayload.Text = decoded.XPayload?.ToString();
                            //    }
                            //    TbResponseContent.Text = string.Empty;
                            //}
                            //else
                            //{
                            //    TbResponseHeader.Text = env.Metadata.Content;
                            //    TbResponsePayload.Text = @"Error!";
                            //    TbResponseContent.Text = string.Empty;
                            //}
                    }
                    break;
                case SsPvoMessageType.Confirm:
                    break;
            }
            */
        }
        private void FillResponseFields(string jwt)
        {
            string[] splitted = !string.IsNullOrWhiteSpace(jwt) ? jwt.Split('.') : null;

            try
            {
                TbResponseHeader.Text = (splitted != null && splitted.Length > 0)
                    ? (string.IsNullOrWhiteSpace(splitted[0]) ? string.Empty : JObject.Parse(splitted[0].FromBase64String()).ToString(Formatting.Indented))
                    : string.Empty;
                TbResponsePayload.Text = (splitted != null && splitted.Length >= 1)
                    ? (string.IsNullOrWhiteSpace(splitted[1]) ? string.Empty : XDocument.Parse(splitted[1].FromBase64String()).ToString(SaveOptions.None))
                    : string.Empty;
                TbResponseContent.Text = jwt;
            }
            finally
            {
            }
        }
        #endregion


        #region excel processing
        private void LoadApplicationsFile(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return;

            ExcelEntries.Clear();

            using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
            {
                var sheet = excelPackage.Workbook.Worksheets["Список заявлений"];
                if (sheet == null) return;

                int numRows = sheet.Dimension.Rows;
                int numCols = sheet.Dimension.Columns;

                var headerToColumnIndex = new Dictionary<string, int>();

                for (int i = 1; i <= numCols; i++)
                {
                    headerToColumnIndex.Add(sheet.Cells[1, i].Value?.ToString(), i);
                }

                for (int r = 2; r <= numRows; r++)
                {
                    ExcelEntries.Add(new SsAppFromExcel
                    {
                        AppNum = $"{sheet.Cells[r, headerToColumnIndex["Номер заявления"]].Value}",
                        LastName = $"{sheet.Cells[r, headerToColumnIndex["Фамилия"]].Value}",
                        FirstName = $"{sheet.Cells[r, headerToColumnIndex["Имя"]].Value}",
                        MiddleName = $"{sheet.Cells[r, headerToColumnIndex["Отчество"]].Value}",
                        Snils = $"{sheet.Cells[r, headerToColumnIndex["Снилс"]].Value}",
                        CgName = $"{sheet.Cells[r, headerToColumnIndex["Конкурс"]].Value}",
                        Level = $"{sheet.Cells[r, headerToColumnIndex["Уровень"]].Value}",
                        Form = $"{sheet.Cells[r, headerToColumnIndex["Форма"]].Value}",
                        FinSource = $"{sheet.Cells[r, headerToColumnIndex["Источник финанс."]].Value}",
                        EpguStatus = $"{sheet.Cells[r, headerToColumnIndex["Статус"]].Value}",
                        RegDate = $"{sheet.Cells[r, headerToColumnIndex["Дата регистр."]].Value}",
                        LastModDate = $"{sheet.Cells[r, headerToColumnIndex["Дата изменения"]].Value}",
                        IsOriginal = $"{sheet.Cells[r, headerToColumnIndex["Оригинал док. об образ-нии"]].Value}",
                        IsAgreed = $"{sheet.Cells[r, headerToColumnIndex["Согласие на зачисл."]].Value}",
                        IsAgreedDate = $"{sheet.Cells[r, headerToColumnIndex["Дата согласия"]].Value}",
                        IsRevoked = $"{sheet.Cells[r, headerToColumnIndex["Отзыв согласия на зачисл."]].Value}",
                        IsRevokedDate = $"{sheet.Cells[r, headerToColumnIndex["Дата отзыва согласия"]].Value}",
                        NeedHostel = $"{sheet.Cells[r, headerToColumnIndex["Общежитие"]].Value}",
                        Rating = $"{sheet.Cells[r, headerToColumnIndex["Рейтинг"]].Value}",
                        EpguId = $"{sheet.Cells[r, headerToColumnIndex["ЕПГУ"]].Value}"
                    });
                }
            }

            CreateJobListFromCurrentApps();
        }

        private void CreateJobListFromCurrentApps()
        {
            if (!_importBatch.AllowAddItems) return;

            if (ExcelEntries == null) ExcelEntries = new List<SsAppFromExcel>();
            if (!ExcelEntries.Any()) return;

            var uniqueSnilsList = ExcelEntries.GroupBy(x => x.Snils).Select(gr => gr.First())
                .OrderBy(x => x.FullName).ToList();

            foreach (var newItem in uniqueSnilsList)
            {
                AddJobItem(newItem);
            }
        }

        private void AddJobItem(SsAppFromExcel newItem)
        {
            if (!_importBatch.AllowAddItems) return;

            bool alreadyInList = ImportItems.Where(x => x.GetProcessedEntity<SsAppFromExcel>() != null)
                .Select(x => x.GetProcessedEntity<SsAppFromExcel>())
                .Any(x => string.Equals(x.Snils, newItem.Snils));

            if (alreadyInList)
            {
                DialogHelper.ShowInfo($"Файл \"{newItem.Snils}\" уже в списке.");
                return;
            }

            var item = new BatchAction.Item(newItem, BatchAction.Scenario.ExcelImport)
            {
                Description = $"{newItem.FullName}; СНИЛС: {newItem.Snils}; ЕПГУ UID: {newItem.EpguId}"
            };
            ImportItems.Add(item);
        }

        public async Task RunJobsAsync()
        {
            if (_importBatch.BatchStatus == BatchAction.Status.InProgress) return;

            _importBatch.StatusChanged += ImportBatchStatusChanged;

            _cancellationTokenSource = new CancellationTokenSource();

            var options = new BatchAction.Options();
            options.AddOrUpdate($"{BatchAction.Options.CommonOptions.HandleItemsInSeparateThread}", true);

            _importBatch.Items.Clear();
            _importBatch.Items.Load(ImportItems);

            await _importBatch.Run(options, _cancellationTokenSource.Token);

            _importBatch.Items.Clear();

            _importBatch.StatusChanged -= ImportBatchStatusChanged;
        }
        #endregion

        #endregion

        #region INPC
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
