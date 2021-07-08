using SsPvo.Client.Enums;
using System;
using System.Xml.Linq;

namespace SsPvo.Client.Messages
{
    public class SsPvoMessage
    {
        #region ctors
        public SsPvoMessage(SsPvoMessageType msgType, SsPvoQueueMsgSubType? subType)
        {
            Guid = Guid.NewGuid();
            MessageType = msgType;
            QueueMessageSubType = subType;
            RequestData = new RequestData();
        }

        public SsPvoMessage(SsPvoMessageType msgType, SsPvoQueueMsgSubType? subType, string ogrn = null, string kpp = null) : this(msgType, subType)
        {
            if (ogrn != null) RequestData.JHeader.Add(nameof(ogrn), ogrn);
            if (kpp != null) RequestData.JHeader.Add(nameof(kpp), kpp);
        }
        #endregion

        #region props
        public Guid Guid { get; }
        public SsPvoMessageType MessageType { get; protected set; }
        public SsPvoQueueMsgSubType? QueueMessageSubType { get; set; }
        public RequestData RequestData { get; protected set; }
        public ResponseData ResponseData { get; set; }
        #endregion

        #region methods
        public void PrepareRequestData(ICsp csp) => RequestData?.Prepare(MessageType, QueueMessageSubType, csp);
        #endregion

        #region classes and structs
        public struct Options
        {
            public SsPvoMessageType Type { get; set; }
            public SsPvoQueueMsgSubType? QueueMsgType { get; set; }
            public string Action { get; set; }
            public uint IdJwt { get; set; }
            public string Cls { get; set; }
            public string EntityType { get; set; }
            public XDocument Payload { get; set; }
        }
        #endregion
    }
}