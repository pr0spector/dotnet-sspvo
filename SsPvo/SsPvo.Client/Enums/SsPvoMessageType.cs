using System.ComponentModel;

namespace SsPvo.Client.Enums
{
    public enum SsPvoMessageType
    {
        [Description("Справочники")]
        Cls,
        [Description("Сертификат")]
        Cert,
        [Description("Действие")]
        Action,
        [Description("Очередь ОО")]
        ServiceQueue,
        [Description("Очередь ЕПГУ")]
        EpguQueue,
        [Description("Подтверждение")]
        Confirm
    }
}
