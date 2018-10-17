# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора требований API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> Эта документация является **предварительной версией** [набора требований](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Этот набор требований еще не полностью реализован, и клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор требований в манифесте надстройки. Методы и свойства, представленные в этом наборе требований, должны быть по отдельности протестированы на доступность перед их использованием.

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
- [Office.EventType](/javascript/api/office/office.eventtype): изменен для поддержки события OfficeThemeChanged посредством добавления записи `OfficeThemeChanged`.
- [Элемент манифеста SupportsSharedFolders](../../manifest/supportssharedfolders.md): добавлен дочерний элемент к элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md). Он определяет, является ли надстройка доступной в сценарии делегирования.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)