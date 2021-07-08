using SsPvo.Client.Enums;
using SsPvo.Client.Messages.Base;
using System;
using System.Xml.Linq;

namespace SsPvo.Client.Messages
{
    public class SsPvoMessageFactory : ISsPvoMessageFactory
    {
        private readonly string _ogrn;
        private readonly string _kpp;

        public SsPvoMessageFactory(string ogrn, string kpp)
        {
            _ogrn = ogrn;
            _kpp = kpp;
        }

        public SsPvoMessage Create(SsPvoMessage.Options options)
        {
            var msg = new SsPvoMessage(options.Type, options.QueueMsgType, _ogrn, _kpp);

            switch (options.Type)
            {
                case SsPvoMessageType.Cls:
                    msg.RequestData.JHeader.Add("cls", options.Cls);
                    break;
                case SsPvoMessageType.Cert:
                    break;
                case SsPvoMessageType.Action:
                    msg.RequestData.JHeader.Add("action", options.Action);
                    msg.RequestData.JHeader.Add("entityType", options.EntityType);
                    msg.RequestData.XPayload = options.Payload;
                    break;
                case SsPvoMessageType.ServiceQueue:
                case SsPvoMessageType.EpguQueue:
                    if (options.QueueMsgType == SsPvoQueueMsgSubType.SingleMessage)
                    {
                        msg.RequestData.JHeader.Add("action", options.Action);
                        msg.RequestData.JHeader.Add("idJwt", options.IdJwt);
                    }
                    break;
                case SsPvoMessageType.Confirm:
                    msg.RequestData.JHeader.Add("action", "messageConfirm");
                    msg.RequestData.JHeader.Add("idJwt", options.IdJwt);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            return msg;
        }

        #region helpers
        public SsPvoMessage CreateDictionaryMessage(string cls) =>
            Create(new SsPvoMessage.Options { Type = SsPvoMessageType.Cls, Cls = cls });

        public SsPvoMessage CreateCertCheckMessage() =>
            Create(new SsPvoMessage.Options { Type = SsPvoMessageType.Cert });

        public SsPvoMessage CreateActionMessage(string action, string entityType, XDocument data) => Create(
            new SsPvoMessage.Options
                {Type = SsPvoMessageType.Action, Action = action, EntityType = entityType, Payload = data});

        public SsPvoMessage CreateCheckQueueMessage(SsPvoQueue queue)
        {
            var msgType = queue == SsPvoQueue.Epgu
                ? SsPvoMessageType.EpguQueue
                : SsPvoMessageType.ServiceQueue;

            return Create(new SsPvoMessage.Options {Type = msgType, QueueMsgType = SsPvoQueueMsgSubType.AllMessages});
        }

        public SsPvoMessage CreateGetQueueItemMessage(SsPvoQueue queue, uint idJwt = 0)
        {
            var msgType = queue == SsPvoQueue.Epgu
                ? SsPvoMessageType.EpguQueue
                : SsPvoMessageType.ServiceQueue;

            return Create(new SsPvoMessage.Options
            {
                Type = msgType, Action = "getMessage", QueueMsgType = SsPvoQueueMsgSubType.SingleMessage, IdJwt = idJwt
            });
        }

        public SsPvoMessage CreateQueueConfirmMessage(uint idJwt)
        {
            return Create(new SsPvoMessage.Options {Type = SsPvoMessageType.Confirm, IdJwt = idJwt});
        }
        #endregion        
    }
}
