# <a name="outlook-add-in-api-requirement-set-17"></a>Набор требований к API надстройки Outlook 1.7

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-17"></a>Новые возможности в версии 1.7

Набор обязательных элементов включает все возможности [набора обязательных элементов 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). Также добавлены следующие возможности:

- Добавлены новые API для расписания повторения встреч и сообщений с запросом о встрече.
- Изменено свойство item.from — теперь оно также доступно в режиме создания.
- Добавлена поддержка событий RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [From](/javascript/api/outlook_1_7/office.from): добавление нового объекта, который предоставляет метод получения начального значения.
- Добавлен объект [Organizer](/javascript/api/outlook_1_7/office.organizer): добавление нового объекта, который предоставляет метод получения значения "Организатор".
- Добавлен объект [Recurrence](/javascript/api/outlook_1_7/office.recurrence): добавление нового объекта, который предоставляет методы получения и установки расписания повторения встреч и методы получения расписания повторения сообщений с запросом об организации встреч.
- Добавлен объект [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone): добавление нового объекта, который представляет формат часового пояса расписания повторения.
- Добавлен объект [SeriesTime](/javascript/api/outlook_1_7/office.seriestime): добавление нового объекта, который предоставляет методы получения и установки даты и времени встреч в повторяющейся серии и методы получения значения даты и времени приглашений на встречу в повторяющейся серии.
- Добавлен объект [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback): добавление нового метода, который добавляет обработчик событий для поддерживаемого события.
- Изменен объект [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom): изменение для получения начального значения в режиме создания.
- Изменен объект [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer): изменение для получения значения "Организатор" в режиме создания.
- Добавлен объект [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence): добавление нового свойства, которое получает или задает объект, предоставляющий методы управления расписанием повторения встреч. Это свойство можно также использовать для получения расписания повторения запросов об организации встречи.
- Добавлен объект [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback): добавление нового метода, который удаляет обработчик событий.
- Добавлен объект [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string): добавление нового свойства, получающего идентификатор серии, к которой относится событие.
- Добавлен объект [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days): добавление нового перечисления, указывающего день недели или тип дня.
- Добавлен объект [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month): добавление нового перечисления, указывающего месяц.
- Добавлен объект [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone): добавление нового перечисления, указывающего часовой пояс повторения.
- Добавлен объект [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype): добавление нового перечисления, определяющего тип повторения.
- Добавлен объект [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber): добавление нового перечисления, указывающего неделю месяца.
- Изменен объект [Office.EventType](/javascript/api/office/office.eventtype): изменение объекта с целью поддержки событий RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged путем добавления элементов `RecurrenceChanged`, `RecipientsChanged` и `AppointmentTimeChanged`, соответственно.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)