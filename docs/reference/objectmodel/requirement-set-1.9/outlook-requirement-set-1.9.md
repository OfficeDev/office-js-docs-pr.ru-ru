---
title: Набор обязательных элементов API для надстройки Outlook 1,9
description: Набор требований 1,9 для API надстройки Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: b2174052a60580a895ef82a4b5d8f00ed6899feb
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628080"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Набор обязательных элементов API для надстройки Outlook 1,9

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-19"></a>Новые возможности 1,9

Набор требований 1,9 включает все функции набора обязательных элементов [1,8](../requirement-set-1.8/outlook-requirement-set-1.8.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для функций "дописывать", "Отправить", "настраиваемые свойства" и "Отображение формы".
- Добавлена поддержка `Dialog.messageChild` .

### <a name="change-log"></a>Журнал изменений

- Добавлена функция [CustomProperties. жеталл](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--): добавляет новую функцию в `CustomProperties` объект, который получает все настраиваемые свойства.
- Добавлено [диалоговое окно Dialog. мессажечилд](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): добавляет новый метод, доставляющий сообщение со страницы узла, например область задач или файл функции без пользовательского интерфейса, в диалоговое окно, открытое на странице.
- Добавлен [элемент манифеста екстендедпермиссионс](../../manifest/extendedpermissions.md): добавляет дочерний элемент в элемент манифеста [VersionOverrides](../../manifest/versionoverrides.md) . Чтобы надстройка поддерживала [функцию Append-on-Send](../../../outlook/append-on-send.md), `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.
- Добавлен [Office. Context. Mailbox. дисплайаппоинтментформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-): добавляет новую функцию в `Mailbox` объект, отображающий существующую встречу. Это асинхронная версия `displayAppointmentForm` метода.
- Добавлено [приложение Office. Context. Mailbox. дисплаймессажеформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-): добавляет новую функцию в `Mailbox` объект, отображающий существующее сообщение. Это асинхронная версия `displayMessageForm` метода.
- Добавлен [Office. Context. Mailbox. дисплайневаппоинтментформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-): добавляет новую функцию в `Mailbox` объект, отображающий новую форму встречи. Это асинхронная версия `displayNewAppointmentForm` метода.
- Добавлен [Office. Context. Mailbox. дисплайневмессажеформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-): добавляет новую функцию в `Mailbox` объект, отображающий новую форму сообщения. Это асинхронная версия `displayNewMessageForm` метода.
- Добавлен элемент [Office. Context. Mailbox. Item. Body. аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-): добавляет в объект новую функцию `Body` , которая добавляет данные в конец тела элемента в режиме создания.
- Добавлен элемент [Office. Context. Mailbox. Item. дисплайрепляллформасинк](office.context.mailbox.item.md#methods): добавляет новую функцию в `Item` объект, отображающий форму "ответить всем" в режиме чтения. Это асинхронная версия `displayReplyAllForm` метода.
- Добавлен элемент [Office. Context. Mailbox. Item. дисплайреплиформасинк](office.context.mailbox.item.md#methods): добавляет новую функцию в `Item` объект, отображающий форму "ответ" в режиме чтения. Это асинхронная версия `displayReplyForm` метода.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
