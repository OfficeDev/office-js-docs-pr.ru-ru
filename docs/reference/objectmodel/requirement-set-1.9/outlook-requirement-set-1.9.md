---
title: Outlook API надстройки 1.9
description: Набор требований 1.9 для Outlook API надстройки.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e73f8805f87950b969be18214a570b747b1e1314
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590500"
---
# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook API надстройки 1.9

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-19"></a>Что нового в 1.9?

Набор требований 1.9 включает все функции набора [требований 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для приложений-на-отправке, настраиваемые свойства и функции отображения формы.
- Добавлена поддержка `Dialog.messageChild` .

### <a name="change-log"></a>Журнал изменений

- Добавлены [CustomProperties.getAll:](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#getall--)добавляет новую функцию в `CustomProperties` объект, который получает все настраиваемые свойства.
- Добавлен [Диалог.messageChild:](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)Добавляет новый метод, который доставляет сообщение со страницы хост-сайта, например области задач или файла функций без пользовательского интерфейса, в диалоговое окно, открытое со страницы.
- Добавлен [элемент манифеста ExtendedPermissions:](../../manifest/extendedpermissions.md)добавляет детский элемент в [элемент манифеста VersionOverrides.](../../manifest/versionoverrides.md) Чтобы надстройка поддержала функцию [приложения-на-отправке,](../../../outlook/append-on-send.md)расширенное разрешение должно быть включено в коллекцию `AppendOnSend` расширенных разрешений.
- Добавлен [Office.context.mailbox.displayAppointmentFormAsync:](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displayappointmentformasync-itemid--options--callback-)Добавляет новую функцию к объекту, отображаемму `Mailbox` существующую встречу. Это версия async `displayAppointmentForm` метода.
- Добавлен [Office.context.mailbox.displayMessageFormAsync:](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaymessageformasync-itemid--options--callback-)добавляет новую функцию к объекту, отображаемму `Mailbox` существующее сообщение. Это версия async `displayMessageForm` метода.
- Добавлен [Office.context.mailbox.displayNewAppointmentFormAsync:](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)добавляет новую функцию к объекту, который отображает новую `Mailbox` форму встречи. Это версия async `displayNewAppointmentForm` метода.
- Добавлен [Office.context.mailbox.displayNewMessageFormAsync:](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)добавляет новую функцию к объекту, который отображает новую `Mailbox` форму сообщения. Это версия async `displayNewMessageForm` метода.
- Добавлен [Office.context.mailbox.item.body.appendOnSendAsync:](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#appendonsendasync-data--options--callback-)добавляет новую функцию к объекту, который добавляет данные в конец тела элемента в режиме `Body` Compose.
- Добавлена [Office.context.mailbox.item.displayReplyAllFormAsync:](office.context.mailbox.item.md#methods)добавляет новую функцию к объекту, отображаемой в режиме "Ответить на все" в режиме `Item` Чтения. Это версия async `displayReplyAllForm` метода.
- Добавлена [Office.context.mailbox.item.displayReplyFormAsync:](office.context.mailbox.item.md#methods)добавляет новую функцию к объекту, отображаемом в режиме `Item` "Ответ" в режиме Чтения. Это версия async `displayReplyForm` метода.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
