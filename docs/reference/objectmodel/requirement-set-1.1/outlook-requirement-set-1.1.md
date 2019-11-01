---
title: Набор требований к API надстройки Outlook 1.1
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 312d40d499531eb6f93d3b1555bfb057cd4651d6
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901957"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="78b7a-102">Набор обязательных элементов API для надстройки Outlook 1.1</span><span class="sxs-lookup"><span data-stu-id="78b7a-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="78b7a-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="78b7a-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="78b7a-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="78b7a-105">Новые возможности в версии 1.1</span><span class="sxs-lookup"><span data-stu-id="78b7a-105">What's new in 1.1?</span></span>

<span data-ttu-id="78b7a-p101">Набор обязательных элементов 1.1 включает все возможности набора обязательных элементов версии 1.0. В нем надстройки получили возможность доступа к тексту сообщений и встреч, а также возможность изменения текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="78b7a-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="78b7a-108">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="78b7a-108">Change log</span></span>

- <span data-ttu-id="78b7a-109">Добавлен объект [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1). Предоставляет методы для добавления и изменения содержимого элемента в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-109">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="78b7a-110">Добавлен объект [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать место проведения собрания в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-110">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="78b7a-111">Добавлен объект [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать получателей для встречи или сообщения в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-111">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="78b7a-112">Добавлен объект [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать тему встречи или сообщения в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-112">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="78b7a-113">Добавлен объект [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать время начала и окончания собрания в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="78b7a-113">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="78b7a-114">Добавлен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="78b7a-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="78b7a-115">Добавлен метод [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback). Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="78b7a-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="78b7a-116">Добавлен метод [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback). Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="78b7a-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="78b7a-117">Добавлено свойство [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body). Получает объект, предоставляющий методы для работы с текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="78b7a-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="78b7a-118">Добавлена строка [Office. Context. Mailbox. Item. BCC](office.context.mailbox.item.md#bcc-recipients) сообщения.</span><span class="sxs-lookup"><span data-stu-id="78b7a-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="78b7a-119">Добавлено свойство [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1). Указывает тип получателя для встречи.</span><span class="sxs-lookup"><span data-stu-id="78b7a-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="78b7a-120">См. также</span><span class="sxs-lookup"><span data-stu-id="78b7a-120">See also</span></span>

- [<span data-ttu-id="78b7a-121">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="78b7a-121">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="78b7a-122">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="78b7a-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="78b7a-123">Начало работы</span><span class="sxs-lookup"><span data-stu-id="78b7a-123">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="78b7a-124">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="78b7a-124">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
