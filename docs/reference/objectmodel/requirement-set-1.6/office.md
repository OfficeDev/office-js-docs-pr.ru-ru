---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: ae764e8cda2b3f14e33b883d054379db7b37a687
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696003"
---
# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Элементы и методы

| Элемент | Тип |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Member |
| [CoercionType](#coerciontype-string) | Member |
| [EventType](#eventtype-string) | Member |
| [SourceProperty](#sourceproperty-string) | Элемент |

### <a name="namespaces"></a>Пространства имен

[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.

### <a name="members"></a>Members

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: строка

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| String|Вызов завершился успешно.|
|`Failed`| Для указания|Вызов завершился ошибкой.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

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
|`Text`| String.|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

<br>

---
---

#### <a name="eventtype-string"></a>EventType: строка

Указывает событие, связанное с обработчиком.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

| Имя | Тип | Описание |
|---|---|---|
|`ItemChanged`| String | Для просмотра выбран другой элемент Outlook, когда область задач закреплена. |

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Минимальная версия набора обязательных элементов для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение |

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
|`Subject`| String.|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|
