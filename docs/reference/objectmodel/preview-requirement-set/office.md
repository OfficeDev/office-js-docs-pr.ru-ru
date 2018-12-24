---
title: Пространство имен Office — предварительная версия набора обязательных элементов
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: a276af19ebd1816ad6bd59af5a75c39f13aa0b3c
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432901"
---
# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Член |
| [CoercionType](#coerciontype-string) | Член |
| [EventType](#eventtype-string) | Член |
| [SourceProperty](#sourceproperty-string) | Член |

### <a name="namespaces"></a>Пространства имен

[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.

[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.

### <a name="members"></a>Элементы

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| Для указания|Вызов завершился успешно.|
|`Failed`| Для указания|Вызов завершился ошибкой.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

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
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

---

####  <a name="eventtype-string"></a>EventType :String

Указывает событие, связанное с обработчиком.

##### <a name="type"></a>Тип:

*   String

##### <a name="properties"></a>Свойства:

| Имя | Тип | Описание | Минимальный набор обязательных элементов |
|---|---|---|---|
|`AppointmentTimeChanged`| String | Произошло изменение даты или времени выбранной встречи либо ряда встреч. | 1.7 |
|`AttachmentsChanged`| String | Было добавлено или удалено вложение для элемента. | Предварительная версия |
|`ItemChanged`| String | Пока область задач закреплена, для просмотра выбран другой элемент Outlook. | 1.5 |
|`OfficeThemeChanged`| String | Тема Office в почтовом ящике была изменена. | Предварительная версия |
|`RecipientsChanged`| String | Произошло изменение списка получателей выбранного элемента или места встречи. | 1.7 |
|`RecurrenceChanged`| String | Расписание повторения выбранного ряда элементов изменилось. | 1.7 |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение |

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
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|