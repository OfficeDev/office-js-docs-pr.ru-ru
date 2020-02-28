---
title: Набор требований к API надстройки Outlook 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: d4fa18f3ab12e22ff30ef841d921f5dac89fd064
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325215"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="d1362-102">Набор обязательных элементов API для надстройки Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="d1362-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="d1362-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d1362-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="d1362-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d1362-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="d1362-105">Новые возможности в версии 1.2</span><span class="sxs-lookup"><span data-stu-id="d1362-105">What's new in 1.2?</span></span>

<span data-ttu-id="d1362-p101">Набор обязательных элементов 1.2 включает все возможности [набора обязательных элементов версии 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Благодаря ему надстройки теперь могут вставлять текст на месте пользовательского указателя (как в теме, так и в тексте сообщения).</span><span class="sxs-lookup"><span data-stu-id="d1362-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="d1362-108">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="d1362-108">Change log</span></span>

- <span data-ttu-id="d1362-109">Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="d1362-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="d1362-110">Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="d1362-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="d1362-111">Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="d1362-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="d1362-112">Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="d1362-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="d1362-113">См. также</span><span class="sxs-lookup"><span data-stu-id="d1362-113">See also</span></span>

- [<span data-ttu-id="d1362-114">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d1362-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="d1362-115">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d1362-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="d1362-116">Начало работы</span><span class="sxs-lookup"><span data-stu-id="d1362-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="d1362-117">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="d1362-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
