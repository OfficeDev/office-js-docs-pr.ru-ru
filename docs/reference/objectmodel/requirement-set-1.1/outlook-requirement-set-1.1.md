---
title: Набор требований к API надстройки Outlook 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: cd284a5871139b7f6bf006a9deb3671a937682f6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450305"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="d3ffa-102">Набор обязательных элементов API для надстройки Outlook 1.1</span><span class="sxs-lookup"><span data-stu-id="d3ffa-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="d3ffa-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="d3ffa-104">В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d3ffa-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="d3ffa-105">Новые возможности в версии 1.1</span><span class="sxs-lookup"><span data-stu-id="d3ffa-105">What's new in 1.1?</span></span>

<span data-ttu-id="d3ffa-p101">Набор обязательных элементов 1.1 включает все возможности набора обязательных элементов версии 1.0. В нем надстройки получили возможность доступа к тексту сообщений и встреч, а также возможность изменения текущего элемента.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="d3ffa-108">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="d3ffa-108">Change log</span></span>

- <span data-ttu-id="d3ffa-109">Добавлен объект [Body](/javascript/api/outlook_1_1/office.body). Предоставляет методы для добавления и изменения содержимого элемента в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-109">Added [Body](/javascript/api/outlook_1_1/office.body) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="d3ffa-110">Добавлен объект [Location](/javascript/api/outlook_1_1/office.location). Предоставляет методы, позволяющие получить и задать место проведения собрания в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-110">Added [Location](/javascript/api/outlook_1_1/office.location) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="d3ffa-111">Добавлен объект [Recipients](/javascript/api/outlook_1_1/office.recipients). Предоставляет методы, позволяющие получить и задать получателей для встречи или сообщения в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-111">Added [Recipients](/javascript/api/outlook_1_1/office.recipients) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="d3ffa-112">Добавлен объект [Subject](/javascript/api/outlook_1_1/office.subject). Предоставляет методы, позволяющие получить и задать тему встречи или сообщения в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-112">Added [Subject](/javascript/api/outlook_1_1/office.subject) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="d3ffa-113">Добавлен объект [Time](/javascript/api/outlook_1_1/office.time). Предоставляет методы, позволяющие получить и задать время начала и окончания собрания в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-113">Added [Time](/javascript/api/outlook_1_1/office.time) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="d3ffa-114">Добавлен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Добавляет файл в сообщение или встречу в качестве вложения.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="d3ffa-115">Добавлен метод [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback). Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="d3ffa-116">Добавлен метод [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback). Удаляет вложение из сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="d3ffa-117">Добавлено свойство [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body). Получает объект, предоставляющий методы для работы с текстом элемента.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="d3ffa-118">Добавлена строка [Office. Context. Mailbox. Item. BCC](office.context.mailbox.item.md#bcc-recipients) сообщения.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="d3ffa-119">Добавлено свойство [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype). Указывает тип получателя для встречи.</span><span class="sxs-lookup"><span data-stu-id="d3ffa-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="d3ffa-120">См. также</span><span class="sxs-lookup"><span data-stu-id="d3ffa-120">See also</span></span>

- [<span data-ttu-id="d3ffa-121">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d3ffa-121">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="d3ffa-122">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d3ffa-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="d3ffa-123">Начало работы</span><span class="sxs-lookup"><span data-stu-id="d3ffa-123">Get started</span></span>](/outlook/add-ins/quick-start)
