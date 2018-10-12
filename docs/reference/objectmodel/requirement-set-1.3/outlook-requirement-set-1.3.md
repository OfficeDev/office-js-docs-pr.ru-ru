# <a name="outlook-add-in-api-requirement-set-13"></a>Набор требований API для надстройки Outlook 1.3

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). 

## <a name="whats-new-in-13"></a>Новые возможности в версии 1.3

Набор требований 1.3 включает все возможности [набора требований версии 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). К нему добавлены перечисленные ниже возможности.

- Добавлена поддержка [команд надстроек](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook).
- Добавлена возможность сохранять и закрывать создаваемый элемент.
- Расширенный объект [Body](/javascript/api/outlook_1_3/office.body) позволяет надстройкам получать или задавать текст целиком.
- Добавлены методы для преобразования идентификаторов из формата EWS в формат REST и наоборот.
- Появилась возможность добавлять сообщения уведомлений на информационную панель элементов.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-). Он возвращает текущий текст в указанном формате.
- Добавлен метод [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-). Он заменяет весь текст указанным текстом.
- Добавлено свойство [Office.context.officeTheme](office.context.md#officetheme-object). Оно предоставляет доступ к цветам темы Office.
- Добавлен объект [Event](/javascript/api/office/office.addincommands.event). Он передается как параметр в функции команд, не требующих пользовательского интерфейса, в надстройке Outlook. Используется для уведомления о завершении обработки.
- Добавлен метод [Office.context.mailbox.item.close](office.context.mailbox.item.md#close). Он закрывает текущий создаваемый элемент.
- Добавлен метод [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback). Он асинхронно сохраняет элемент.
- Добавлено свойство [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages). Оно получает сообщения уведомления для элемента.
- Добавлен метод [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string). Он преобразует идентификатор элемента из формата REST в формат EWS.
- Добавлен метод [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Он преобразует идентификатор элемента из формата EWS в формат REST.
- Добавлено свойство [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype). Оно указывает тип сообщения уведомления для встречи или сообщения.
- Добавлено свойство [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion). Оно указывает версию REST API, которая соответствует идентификатору элемента в формате REST.
- Добавлен объект [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages). Он предоставляет методы для доступа к сообщениям уведомления в надстройке Outlook.
- Добавлен тип [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails). Он возвращается методом `NotificationMessages.getAllAsync`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)