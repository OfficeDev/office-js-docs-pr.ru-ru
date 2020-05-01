---
title: Пространство имен Office — Предварительная версия набора требований
description: Элементы пространства имен Office, доступные для надстроек Outlook с использованием набора обязательных элементов API почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d72e5c78a7fd8d3c00b8f84e7d9b05ee6defc0c5
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890860"
---
# <a name="office-mailbox-preview-requirement-set"></a>Office (набор требований предварительного просмотра почтового ящика)

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="properties"></a>Properties

| Свойство | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [контекст](office.context.md) | Создание<br>Чтение | [Context](/javascript/api/office/office.context?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Перечисления

| Перечисление | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Создание<br>Чтение | Строка | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Пространства имен

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.

## <a name="enumeration-details"></a>Сведения о перечислении

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: строка

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| Строка|Вызов завершился успешно.|
|`Failed`| Для указания|Вызов завершился ошибкой.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: строка

Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Html`| Строка|Запрашивает возврат данных в формате HTML.|
|`Text`| Строка|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: строка

Указывает событие, связанное с обработчиком.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

| Имя | Тип | Описание | Набор минимальных требований |
|---|---|---|:---:|
|`AppointmentTimeChanged`| Строка | Дата или время выбранной встречи или ряда изменились. | 1.7 |
|`AttachmentsChanged`| Строка | Вложение было добавлено или удалено из элемента. | 1.8 |
|`EnhancedLocationsChanged`| Строка | Расположение выбранной встречи изменилось. | 1.8 |
|`ItemChanged`| Строка | Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | 1.5 |
|`OfficeThemeChanged`| Строка | Тема Office в почтовом ящике изменилась. | Предварительная версия |
|`RecipientsChanged`| Строка | Список получателей выбранного элемента или места встречи изменился. | 1.7 |
|`RecurrenceChanged`| Строка | Шаблон повторения выбранного ряда изменился. | 1.7 |

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5 |
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="sourceproperty-string"></a>Перестрока: строка

Указывает источник данных, возвращаемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Body`| Строка|Источник данных — текст сообщения.|
|`Subject`| Строка|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|
