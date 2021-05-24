---
title: Набор требований к API надстройки Outlook 1.2
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590402"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="3fab4-103">Набор требований к API надстройки Outlook 1.2</span><span class="sxs-lookup"><span data-stu-id="3fab4-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="3fab4-104">Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="3fab4-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3fab4-105">В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3fab4-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="3fab4-106">Новые возможности в версии 1.2</span><span class="sxs-lookup"><span data-stu-id="3fab4-106">What's new in 1.2?</span></span>

<span data-ttu-id="3fab4-107">Набор требований 1.2 включает все функции набора [требований 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span><span class="sxs-lookup"><span data-stu-id="3fab4-107">Requirement set 1.2 includes all of the features of [requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span></span> <span data-ttu-id="3fab4-108">Благодаря ему надстройки теперь могут вставлять текст на месте пользовательского указателя (как в теме, так и в тексте сообщения).</span><span class="sxs-lookup"><span data-stu-id="3fab4-108">It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="3fab4-109">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="3fab4-109">Change log</span></span>

- <span data-ttu-id="3fab4-110">Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно возвращает данные, выбранные в теме или тексте сообщения.</span><span class="sxs-lookup"><span data-stu-id="3fab4-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="3fab4-111">Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно вставляет данные в текст или тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="3fab4-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="3fab4-112">Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="3fab4-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="3fab4-113">Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.</span><span class="sxs-lookup"><span data-stu-id="3fab4-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="3fab4-114">См. также</span><span class="sxs-lookup"><span data-stu-id="3fab4-114">See also</span></span>

- [<span data-ttu-id="3fab4-115">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3fab4-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="3fab4-116">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3fab4-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3fab4-117">Начало работы</span><span class="sxs-lookup"><span data-stu-id="3fab4-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="3fab4-118">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="3fab4-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
