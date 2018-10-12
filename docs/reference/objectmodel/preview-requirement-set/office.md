 

# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Члены и методы

| Член | Тип |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Член |
| [CoercionType](#coerciontype-string) | Член |
| [EventType](#eventtype-string) | Член |
| [SourceProperty](#sourceproperty-string) | Член |

### <a name="namespaces"></a>Пространства имен

[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.

### <a name="members"></a>Члены

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип:

*   Строка​

##### <a name="properties"></a>Свойства:

|Name| Тип| Описание|
|---|---|---|
|`Succeeded`| Строка​|Вызов завершился успешно.|
|`Failed`| Строка​|Вызов не удался.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

---

####  <a name="coerciontype-string"></a>CoercionType :String

Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.

##### <a name="type"></a>Тип:

*   Строка​

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Html`| Строка​|Запрашивает возврат данных в формате HTML.|
|`Text`| Строка​|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

---

####  <a name="eventtype-string"></a>EventType :String

Указывает событие, связанное с обработчиком событий.

##### <a name="type"></a>Тип:

*   Строка​

##### <a name="properties"></a>Свойства:

| Name | Тип | Описание | Минимальный набор требований |
|---|---|---|---|
|`AppointmentTimeChanged`| Строка​ | Дата или время выбранной встречи или серии была изменена. | 1.7 |
|`ItemChanged`| Строка​ | Выбранный элемент изменился. | 1.5 |
|`OfficeThemeChanged`| Строка​ | Выбранный элемент изменился. | Предварительная версия |
|`RecipientsChanged`| Строка​ | Список получателей в выбранном элементе или расположение встречи изменен(-о). | 1.7 |
|`RecurrenceChanged`| Строка​ | Расписание повторения выбранной серии было изменено. | 1.7 |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение |

---

####  <a name="sourceproperty-string"></a>SourceProperty :String

Указывает источник данных, возвращаемых вызванным методом.

##### <a name="type"></a>Тип:

*   Строка​

##### <a name="properties"></a>Свойства:

|Name| Тип| Описание|
|---|---|---|
|`Body`| Строка​|Источник данных — текст сообщения.|
|`Subject`| Строка​|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|