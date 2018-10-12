 

# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose (создание) или read (чтение)|

##### <a name="members-and-methods"></a>Члены и методы

| Член | Тип |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Член |
| [CoercionType](#coerciontype-string) | Член |
| [EventType](#eventtype-string) | Член |
| [SourceProperty](#sourceproperty-string) | Член |

### <a name="namespaces"></a>Пространства имен

[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.

### <a name="members"></a>Члены

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| String|Вызов завершился успешно.|
|`Failed`| String|Вызов не удался.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose (создание) или read (чтение)|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Html`| String|Запрашивает возврат данных в формате HTML.|
|`Text`| String|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose (создание) или read (чтение)|

---

####  <a name="eventtype-string"></a>EventType :String

Указывает событие, связанное с обработчиком событий.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

| Имя | Тип | Описание | Минимальный набор требований |
|---|---|---|---|
|`AppointmentTimeChanged`| String | Дата или время выбранной встречи или серии была изменена. | 1.7 |
|`ItemChanged`| String | Выбранный элемент изменился. | 1.5 |
|`OfficeThemeChanged`| String | Выбранный элемент изменился. | Предварительная версия |
|`RecipientsChanged`| String | Изменился список получателей в выбранном элементе или изменилось расположение встречи. | 1.7 |
|`RecurrenceChanged`| String | Расписание повторения выбранной серии было изменено. | 1.7 |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose (создание) или read (чтение) |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Указывает источник данных, возвращаемых вызванным методом.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Body`| String|Источник данных — текст сообщения.|
|`Subject`| String|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Compose (создание) или read (чтение)|