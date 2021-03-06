---
title: Office пространства имен — набор требований 1.5
description: Office пространства имен, доступных для Outlook надстройки с помощью API почтовых ящиков, установленного 1.5.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 46b70185ce983721c75093351e47a02eb8b9e7cd
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590857"
---
# <a name="office-mailbox-requirement-set-15"></a>Office (набор требований к почтовым ящикам 1.5)

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

## <a name="properties"></a>Свойства

| Свойство | Режимы | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|:---:|
| [контекст](office.context.md) | Создание<br>Чтение | [Context](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>Перечисления

| Перечисление | Режимы | Тип возвращаемых данных | Minimum<br>набор требований |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Создание<br>Чтение | Строка | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Создание<br>Чтение | Строка | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>Пространства имен

[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .

## <a name="enumeration-details"></a>Сведения о переумериях

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| Строка|Вызов завершился успешно.|
|`Failed`| String|Вызов завершился ошибкой.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: String

Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства

|Имя| Тип| Описание|
|---|---|---|
|`Html`| Строка|Запрашивает возврат данных в формате HTML.|
|`Text`| String|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: String

Указывает событие, связанное с обработчиком.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства

| Имя | Тип | Описание | Минимальный набор требований |
|---|---|---|:---:|
|`ItemChanged`| Строка | Другой элемент Outlook для просмотра при закреплении области задач. | 1.5 |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5 |
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение |

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty: String

Указывает источник данных, возвращаемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства

|Имя| Тип| Описание|
|---|---|---|
|`Body`| Строка|Источник данных — текст сообщения.|
|`Subject`| String|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Применимый режим Outlook](../../../outlook/outlook-add-ins-overview.md#extension-points)| Создание или чтение|
