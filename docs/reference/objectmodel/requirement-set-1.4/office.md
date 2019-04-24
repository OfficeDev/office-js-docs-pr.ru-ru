---
title: Пространство имен Office — набор обязательных элементов 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c60195ddfc42d962427127bf601bca3d41797566
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450326"
---
# <a name="office"></a>Office

Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="namespaces"></a>Пространства имен

[context.](Office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.

[MailboxEnums.](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.

### <a name="members"></a>Элементы

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

Указывает результат асинхронного вызова.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Succeeded`| Строка|Вызов завершился успешно.|
|`Failed`| Для указания|Вызов завершился ошибкой.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="coerciontype-string"></a>CoercionType :String

Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Html`| Строка|Запрашивает возврат данных в формате HTML.|
|`Text`| Строка|Запрашивает возврат данных в формате текста.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="sourceproperty-string"></a>SourceProperty :String

Указывает источник данных, возвращаемых вызванным методом.

##### <a name="type"></a>Тип

*   String

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`Body`| Строка|Источник данных — текст сообщения.|
|`Subject`| Строка|Источник данных — тема сообщения.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований к почтовому ящику](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](/outlook/add-ins/#extension-points)| Создание или чтение|
