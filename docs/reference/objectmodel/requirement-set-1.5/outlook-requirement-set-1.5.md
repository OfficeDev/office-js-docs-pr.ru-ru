---
title: Набор обязательных элементов API для надстройки Outlook 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 65453a0cff8db682f5f573c25a9afa4e9ff63f67
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163740"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Набор обязательных элементов API для надстройки Outlook 1.5

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-15"></a>Новые возможности в версии 1.5

Набор обязательных элементов 1.5 включает все возможности [набора обязательных элементов версии 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [закрепляемых областей задач](../../../outlook/pinnable-taskpane.md).
- Добавлена поддержка вызовов [REST API](../../../outlook/use-rest-api.md).
- Добавлена возможность отметить вложение как встроенное.
- Добавлена возможность закрыть область задач или диалоговое окно.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods). Добавляет обработчик для поддерживаемого события.
- Добавлено [приложение Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#methods): удаляет обработчики событий для поддерживаемого типа события.
- Добавлено свойство [Office.EventType](office.md#eventtype-string). Указывает событие, связанное с обработчиком, и включает поддержку события ItemChanged.
- Добавлен метод [Office.context.mailbox.restUrl](office.context.mailbox.md#properties). Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.
- Изменен метод [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods). Добавлен новый вариант этого метода с новой подписью (`getCallbackTokenAsync([options], callback)`). Исходная версия по-прежнему доступна и осталась без изменений.
- Добавлен метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Изменен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods). Новое значение в словаре `options` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
