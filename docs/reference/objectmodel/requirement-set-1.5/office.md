---
title: Пространство имен Office — набор обязательных элементов 1,5
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 848aa30c07b936c8454b2833d5dce3e1d15ee193
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891350"
---
# <a name="office-mailbox-requirement-set-15"></a>Office (набор требований для почтового ящика 1,5)

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

##### <a name="properties"></a>Properties

| Свойство | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [контекст](office.context.md) | Создание<br>Чтение | [Context](/javascript/api/office/office.context?view=outlook-js-1.5) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Перечисления

| Перечисление | Способов | Тип возвращаемых данных | Минимальные<br>набор требований |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Создание<br>Чтение | Строка | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Пространства имен

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.

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
|`ItemChanged`| Строка | Для просмотра выбран другой элемент Outlook, когда область задач закреплена. | 1.5 |

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
|`Body`| Строка|Источник данных — текст сообщения.|
|`Subject`| Строка|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Requirements

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|
