---
title: Набор требований к API надстройки Outlook 1.2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dfd61738fe9e941d7d0f3a70173ab1f087f9f552
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30869934"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="06cde-102">Набор обязательных элементов API для надстройки Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="06cde-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="06cde-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="06cde-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="06cde-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="06cde-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="06cde-105">Новые возможности в версии 1.2</span><span class="sxs-lookup"><span data-stu-id="06cde-105">What's new in 1.2?</span></span>

<span data-ttu-id="06cde-p101">Набор обязательных элементов 1.2 включает все возможности [набора обязательных элементов версии 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Благодаря ему надстройки теперь могут вставлять текст на месте пользовательского указателя (как в теме, так и в тексте сообщения).</span><span class="sxs-lookup"><span data-stu-id="06cde-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="06cde-108">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="06cde-108">Change log</span></span>

- <span data-ttu-id="06cde-109">Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string). Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="06cde-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="06cde-110">Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback). Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="06cde-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="06cde-111">Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="06cde-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="06cde-112">Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="06cde-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="06cde-113">См. также</span><span class="sxs-lookup"><span data-stu-id="06cde-113">See also</span></span>

- [<span data-ttu-id="06cde-114">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="06cde-114">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="06cde-115">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="06cde-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="06cde-116">Начало работы</span><span class="sxs-lookup"><span data-stu-id="06cde-116">Get started</span></span>](/outlook/add-ins/quick-start)
