# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора требований API для надстройки Outlook

Вложенный набор API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> Примечание. Эта документация относится к **предварительной версии** [набора требований](/javascript/office/requirement-sets/outlook-api-requirement-sets). Этот набор требований еще не полностью реализован, и клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор требований в манифесте надстройки. Прежде чем использовать методы и свойства, добавленные в этом наборе требований, следует отдельно проверять их на доступность.

Предварительная версия набора требований включает все возможности [набора требований 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

- [SharedProperties](/javascript/api/outlook/office.sharedproperties): добавлен новый объект, который представляет свойства для элемента appointment или message в общей папке, календаре или почтовом ящике.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-): новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением (`allowEvent`). Это значение используется для отмены выполнения события.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback): добавлен новый метод, который прикрепляет файл из кодирования base64 к сообщению или встрече.
- [Office.context.mailbox.item.getInitializationContextAsync.](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback): добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки сообщением с действиями](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback): добавлен новый метод, который возвращает объект, представляющий sharedProperties элемента appointment или message.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference): добавлен доступ к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions): добавлено новое перечисление битового флага, указывающее делегированные разрешения.
- [Office.EventType](/javascript/api/office/office.eventtype): изменено для поддержки событий OfficeThemeChanged посредством добавления записи `OfficeThemeChanged`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)