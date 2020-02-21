---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 23f3fb705c03eabd8ee7fce53f4c89a48128672f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165350"
---
# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="properties"></a>Properties

| Свойство | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [контекст](office.context.md) | Создание<br>Чтение | [Context](/javascript/api/office/office.context?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Перечисления

| Перечисление | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Создание<br>Чтение | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Создание<br>Чтение | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Пространства имен

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.

## <a name="enumeration-details"></a>Сведения о перечислении

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: строка

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| String|Вызов завершился успешно.|
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
|`Html`| String|Запрашивает возврат данных в формате HTML.|
|`Text`| String|Запрашивает возврат данных в формате текста.|

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
|`AppointmentTimeChanged`| String | Дата или время выбранной встречи или ряда изменились. | 1.7 |
|`ItemChanged`| String | Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | 1.5 |
|`RecipientsChanged`| String | Список получателей выбранного элемента или места встречи изменился. | 1.7 |
|`RecurrenceChanged`| String | Шаблон повторения выбранного ряда изменился. | 1.7 |

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5 |
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение |

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
|`Body`| String|Источник данных — текст сообщения.|
|`Subject`| String|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|
