---
title: Набор обязательных элементов API для надстройки Outlook 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0489e4efa763b3963fcdc78ec894db46fa06362
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870871"
---
# <a name="outlook-add-in-api-requirement-set-15"></a>Набор обязательных элементов API для надстройки Outlook 1.5

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-15"></a>Новые возможности в версии 1.5

Набор обязательных элементов 1.5 включает все возможности [набора обязательных элементов версии 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). В нем добавлены перечисленные ниже возможности.

- Добавлена поддержка [закрепляемых областей задач](/outlook/add-ins/pinnable-taskpane).
- Добавлена поддержка вызовов [REST API](/outlook/add-ins/use-rest-api).
- Добавлена возможность отметить вложение как встроенное.
- Добавлена возможность закрыть область задач или диалоговое окно.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback). Добавляет обработчик для поддерживаемого события.
- Добавлено [приложение Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): удаляет обработчики событий для поддерживаемого типа события.
- Добавлено свойство [Office.EventType](office.md#eventtype-string). Указывает событие, связанное с обработчиком, и включает поддержку события ItemChanged.
- Добавлен метод [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string). Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.
- Изменен метод [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback). Добавлен новый вариант этого метода с новой подписью (`getCallbackTokenAsync([options], callback)`). Исходная версия по-прежнему доступна и осталась без изменений.
- Добавлен метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Изменен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Новое значение в словаре `options` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/outlook/add-ins/quick-start)
