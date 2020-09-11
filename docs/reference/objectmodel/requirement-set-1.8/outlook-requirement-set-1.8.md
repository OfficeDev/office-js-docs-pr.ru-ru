---
title: Набор обязательных элементов API для надстройки Outlook 1.8
description: Набор требований 1,8 для API надстроек Outlook
ms.date: 05/20/2020
localization_priority: Normal
ms.openlocfilehash: 6da1de2529f48fa15d93846d4afcca0718b81165
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431537"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Набор обязательных элементов API для надстройки Outlook 1.8

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-18"></a>Новые возможности в версии 1.8

Набор обязательных элементов 1.8 включает все возможности [набора обязательных элементов версии 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для вложений, категорий, делегирования доступа, расширенного расположения, заголовков Интернета и функций блокирования при отправке.
- Добавлен необязательный параметр `options` для метода Event.completed.
- Добавлена поддержка `AttachmentsChanged` и `EnhancedLocationsChanged` события.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий содержимое вложения.
- Добавлена [аттачментдетаилскомпосе](/javascript/api/outlook/office.attachmentdetailscompose?view=outlook-js-1.8&preserve-view=true): добавляет новый объект, представляющий сведения о вложении в режиме создания.
- Добавлен объект [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий категории элемента.
- Добавлен объект [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).
- Добавлен объект [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий набор местоположений для встречи.
- Добавлен объект [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий заголовки Интернета в элементе сообщения. Только в режиме создания.
- Добавлен объект [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий расположение. Только для чтения.
- Добавлен объект [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий идентификатор расположения.
- Добавлен объект [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий главный список категорий для почтового ящика.
- Добавлен объект [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8&preserve-view=true). Добавляет новый объект, представляющий свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.
- Добавлен [элемент манифеста SupportsSharedFolders](../../manifest/supportssharedfolders.md). Добавляет дочерний элемент к элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md). Он определяет, доступна ли надстройка в сценариях делегирования.
- Добавлен объект [Office.context.mailbox.masterCategories](office.context.mailbox.md#properties). Добавляет новое свойство, представляющее главный список категорий для почтового ящика.
- Добавлен объект [Office.context.mailbox.item.categories](office.context.mailbox.item.md#properties). Добавляет новое свойство, представляющее набор категорий для элемента.
- Добавлен объект [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#methods). Добавляет новый метод, позволяющий вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.
- Добавлен объект [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#properties). Добавляет новое свойство, представляющее набор местоположений для встречи.
- Добавлен объект [Office.context.mailbox.item.getAllInternetHeadersAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, получающий заголовки Интернета для элемента сообщения. Только в режиме чтения.
- Добавлен объект [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, позволяющий получить содержимое определенного вложения.
- Добавлен объект [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, получающий вложенные в элемент объекты в режиме создания.
- Добавлен объект [Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, получающий идентификатор сохраненного элемента встречи или сообщения.
- Добавлен объект [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, позволяющий получить объект, представляющий свойства sharedProperties элемента встречи или сообщения.
- Добавлен объект [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#properties). Добавляет новое свойство, представляющее настраиваемые заголовки Интернета в элементе сообщения. Только в режиме создания.
- Изменен объект [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-). Добавляет новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением (`allowEvent`). Это значение используется для отмены выполнения события.
- Добавлен объект [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true). Добавляет новое перечисление, указывающее форматирование, применяемое к содержимому вложения.
- Добавлен объект [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8&preserve-view=true). Добавляет новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.
- Добавлен объект [Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8&preserve-view=true). Добавляет новое перечисление, указывающее цвета, доступные для сопоставления с категориями.
- Добавлен объект [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8&preserve-view=true). Добавляет перечисление нового битового флага, указывающее разрешения на делегирование.
- Добавлен объект [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8&preserve-view=true). Добавляет новое перечисление, определяющее тип расположения встречи.
- Изменен объект [Office.EventType](/javascript/api/office/office.eventtype). Добавляет поддержку событий `AttachmentsChanged` и `EnhancedLocationsChanged`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
