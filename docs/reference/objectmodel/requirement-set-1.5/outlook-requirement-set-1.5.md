---
title: Набор обязательных элементов API для надстройки Outlook 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0489e4efa763b3963fcdc78ec894db46fa06362
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451817"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="cbebb-102">Набор обязательных элементов API для надстройки Outlook 1.5</span><span class="sxs-lookup"><span data-stu-id="cbebb-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="cbebb-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="cbebb-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="cbebb-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="cbebb-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="cbebb-105">Новые возможности в версии 1.5</span><span class="sxs-lookup"><span data-stu-id="cbebb-105">What's new in 1.5?</span></span>

<span data-ttu-id="cbebb-p101">Набор обязательных элементов 1.5 включает все возможности [набора обязательных элементов версии 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). В нем добавлены перечисленные ниже возможности.</span><span class="sxs-lookup"><span data-stu-id="cbebb-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="cbebb-108">Добавлена поддержка [закрепляемых областей задач](/outlook/add-ins/pinnable-taskpane).</span><span class="sxs-lookup"><span data-stu-id="cbebb-108">Added support for [pinnable task panes](/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="cbebb-109">Добавлена поддержка вызовов [REST API](/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="cbebb-109">Added support for calling [REST APIs](/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="cbebb-110">Добавлена возможность отметить вложение как встроенное.</span><span class="sxs-lookup"><span data-stu-id="cbebb-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="cbebb-111">Добавлена возможность закрыть область задач или диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="cbebb-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="cbebb-112">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="cbebb-112">Change log</span></span>

- <span data-ttu-id="cbebb-113">Добавлен метод [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback). Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="cbebb-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="cbebb-114">Добавлено [приложение Office. Context. Mailbox. removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): удаляет обработчики событий для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="cbebb-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="cbebb-115">Добавлено свойство [Office.EventType](office.md#eventtype-string). Указывает событие, связанное с обработчиком, и включает поддержку события ItemChanged.</span><span class="sxs-lookup"><span data-stu-id="cbebb-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="cbebb-116">Добавлен метод [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string). Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="cbebb-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="cbebb-p102">Изменен метод [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback). Добавлен новый вариант этого метода с новой подписью (`getCallbackTokenAsync([options], callback)`). Исходная версия по-прежнему доступна и осталась без изменений.</span><span class="sxs-lookup"><span data-stu-id="cbebb-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="cbebb-119">Добавлен метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span><span class="sxs-lookup"><span data-stu-id="cbebb-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="cbebb-120">Изменен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Новое значение в словаре `options` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cbebb-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="cbebb-121">Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cbebb-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="cbebb-122">Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cbebb-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="cbebb-123">См. также</span><span class="sxs-lookup"><span data-stu-id="cbebb-123">See also</span></span>

- [<span data-ttu-id="cbebb-124">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="cbebb-124">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="cbebb-125">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="cbebb-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="cbebb-126">Начало работы</span><span class="sxs-lookup"><span data-stu-id="cbebb-126">Get started</span></span>](/outlook/add-ins/quick-start)
