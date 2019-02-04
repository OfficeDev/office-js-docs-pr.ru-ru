---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 01/30/2019
localization_priority: Priority
ms.openlocfilehash: 47db6824fa328a61e47d577a85375f43e73f16df
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701787"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки. Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.

Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent): добавлен новый объект, представляющий содержимое вложения.
- [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation): добавлен новый объект, представляющий набор местоположений для встречи.
- [InternetHeaders](/javascript/api/outlook/office.internetheaders): добавлен новый объект, представляющий заголовки Интернета в элементе сообщения.
- [LocationDetails](/javascript/api/outlook/office.locationdetails): добавлен новый объект, представляющий расположение. Только для чтения.
- [LocationDetails](/javascript/api/outlook/office.locationidentifier): добавлен новый объект, представляющий идентификатор расположения.
- [SharedProperties](/javascript/api/outlook/office.sharedproperties): добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) — новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением (`allowEvent`). Это значение используется для отмены выполнения события.
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback): добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.
- [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation): добавлено новое свойство, представляющее набор местоположений для встречи.
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent): добавлен новый метод, позволяющий получить содержимое определенного вложения.
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails): добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback): добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки сообщением с действиями](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback): добавлен новый метод, позволяющий получить объект, который представляет собой свойства sharedProperties элемента встречи или сообщения.
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders): добавлено новое свойство, представляющее заголовки Интернета в элементе сообщения.
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference): добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus): добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions): добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.
- [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype): добавлено новое перечисление, которое определяет тип расположения встречи.
- [Office.EventType](/javascript/api/office/office.eventtype): этот элемент изменен для поддержки событий AttachmentsChanged и OfficeThemeChanged путем добавления записей `AttachmentsChanged` и `OfficeThemeChanged` соответственно.
- [Элемент манифеста SupportsSharedFolders](../../manifest/supportssharedfolders.md): к элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент. Он определяет, доступна ли надстройка в сценариях делегирования.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)
