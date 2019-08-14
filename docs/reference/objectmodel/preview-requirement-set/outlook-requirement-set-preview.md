---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: b563d6cfc279a18a6a61f39c33a5ab42e1bd6984
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395710"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="839bb-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="839bb-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="839bb-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="839bb-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="839bb-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="839bb-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="839bb-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="839bb-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="839bb-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="839bb-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="839bb-107">Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.</span><span class="sxs-lookup"><span data-stu-id="839bb-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span>
>
> <span data-ttu-id="839bb-108">Чтобы использовать предварительные версии API:</span><span class="sxs-lookup"><span data-stu-id="839bb-108">To use preview APIs:</span></span>
>
> - <span data-ttu-id="839bb-109">Необходимо ссылаться на **бета-версию** библиотеки в сети CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span><span class="sxs-lookup"><span data-stu-id="839bb-109">You must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span>
> - <span data-ttu-id="839bb-110">Кроме того, может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider), чтобы получить более новые сборки Office.</span><span class="sxs-lookup"><span data-stu-id="839bb-110">You may also need to join the [Office Insider program](https://products.office.com/office-insider) for access to more recent Office builds.</span></span>

<span data-ttu-id="839bb-111">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="839bb-111">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="839bb-112">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="839bb-112">Features in preview</span></span>

<span data-ttu-id="839bb-113">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="839bb-113">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="839bb-114">Вложения</span><span class="sxs-lookup"><span data-stu-id="839bb-114">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="839bb-115">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="839bb-115">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="839bb-116">Добавлен новый объект, представляющий содержимое вложения.</span><span class="sxs-lookup"><span data-stu-id="839bb-116">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="839bb-117">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-117">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="839bb-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="839bb-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="839bb-119">Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="839bb-119">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="839bb-120">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="839bb-121">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-121">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="839bb-122">Добавлен новый метод, позволяющий получить содержимое определенного вложения.</span><span class="sxs-lookup"><span data-stu-id="839bb-122">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="839bb-123">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-123">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="839bb-124">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-124">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="839bb-125">Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="839bb-125">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="839bb-126">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-126">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="839bb-127">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="839bb-127">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="839bb-128">Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.</span><span class="sxs-lookup"><span data-stu-id="839bb-128">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="839bb-129">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-129">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="839bb-130">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="839bb-130">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="839bb-131">Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.</span><span class="sxs-lookup"><span data-stu-id="839bb-131">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="839bb-132">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="839bb-133">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="839bb-133">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="839bb-134">Добавлено событие `AttachmentsChanged` для объекта `Item`.</span><span class="sxs-lookup"><span data-stu-id="839bb-134">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="839bb-135">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="839bb-136">Блокировка при отправке</span><span class="sxs-lookup"><span data-stu-id="839bb-136">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="839bb-137">Event.completed</span><span class="sxs-lookup"><span data-stu-id="839bb-137">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="839bb-138">Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="839bb-138">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="839bb-139">Это значение используется для отмены выполнения события.</span><span class="sxs-lookup"><span data-stu-id="839bb-139">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="839bb-140">**Доступно в** Outlook в Интернете (классическая версия), Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-140">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="categories"></a><span data-ttu-id="839bb-141">Категории</span><span class="sxs-lookup"><span data-stu-id="839bb-141">Categories</span></span>

<span data-ttu-id="839bb-142">В Outlook пользователь может группировать сообщения и встречи, используя категории для выделения их цветом.</span><span class="sxs-lookup"><span data-stu-id="839bb-142">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="839bb-143">Пользователь определяет категории в главном списке своего почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="839bb-143">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="839bb-144">Затем он может применить одну или несколько категорий к элементу.</span><span class="sxs-lookup"><span data-stu-id="839bb-144">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="839bb-145">Эта возможность не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="839bb-145">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="839bb-146">Categories</span><span class="sxs-lookup"><span data-stu-id="839bb-146">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="839bb-147">Добавлен новый объект, представляющий категории элемента.</span><span class="sxs-lookup"><span data-stu-id="839bb-147">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="839bb-148">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="839bb-149">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="839bb-149">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="839bb-150">Добавлен новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).</span><span class="sxs-lookup"><span data-stu-id="839bb-150">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="839bb-151">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-151">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="839bb-152">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="839bb-152">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="839bb-153">Добавлен новый объект, представляющий главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="839bb-153">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="839bb-154">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-154">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="839bb-155">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="839bb-155">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="839bb-156">Добавлено новое свойство, представляющее главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="839bb-156">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="839bb-157">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-157">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="839bb-158">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="839bb-158">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="839bb-159">Добавлено новое свойство, представляющее набор категорий для элемента.</span><span class="sxs-lookup"><span data-stu-id="839bb-159">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="839bb-160">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="839bb-161">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="839bb-161">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="839bb-162">Добавлено новое перечисление, указывающее цвета, доступные для сопоставления с категориями.</span><span class="sxs-lookup"><span data-stu-id="839bb-162">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="839bb-163">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="839bb-164">Делегированный доступ</span><span class="sxs-lookup"><span data-stu-id="839bb-164">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="839bb-165">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="839bb-165">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="839bb-166">Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="839bb-166">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="839bb-167">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="839bb-168">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-168">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="839bb-169">Добавлен новый метод, получающий идентификатор сохраненного элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="839bb-169">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="839bb-170">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="839bb-171">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-171">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="839bb-172">Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="839bb-172">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="839bb-173">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="839bb-174">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="839bb-174">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="839bb-175">Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.</span><span class="sxs-lookup"><span data-stu-id="839bb-175">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="839bb-176">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="839bb-177">Элемент манифеста SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="839bb-177">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="839bb-178">К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент.</span><span class="sxs-lookup"><span data-stu-id="839bb-178">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="839bb-179">Он определяет, доступна ли надстройка в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="839bb-179">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="839bb-180">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="839bb-181">Расширенные функции расположения</span><span class="sxs-lookup"><span data-stu-id="839bb-181">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="839bb-182">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="839bb-182">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="839bb-183">Добавлен новый объект, представляющий набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="839bb-183">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="839bb-184">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="839bb-185">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="839bb-185">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="839bb-186">Добавлен новый объект, представляющий расположение.</span><span class="sxs-lookup"><span data-stu-id="839bb-186">Added a new object that represents a location.</span></span> <span data-ttu-id="839bb-187">Только для чтения.</span><span class="sxs-lookup"><span data-stu-id="839bb-187">Read only.</span></span>

<span data-ttu-id="839bb-188">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-188">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="839bb-189">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="839bb-189">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="839bb-190">Добавлен новый объект, представляющий идентификатор расположения.</span><span class="sxs-lookup"><span data-stu-id="839bb-190">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="839bb-191">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-191">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="839bb-192">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="839bb-192">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="839bb-193">Добавлено новое свойство, представляющее набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="839bb-193">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="839bb-194">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-194">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="839bb-195">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="839bb-195">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="839bb-196">Добавлено новое перечисление, которое определяет тип расположения встречи.</span><span class="sxs-lookup"><span data-stu-id="839bb-196">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="839bb-197">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-197">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="839bb-198">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="839bb-198">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="839bb-199">Добавлено событие `EnhancedLocationsChanged` для объекта `Item`.</span><span class="sxs-lookup"><span data-stu-id="839bb-199">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="839bb-200">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-200">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="839bb-201">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="839bb-201">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="839bb-202">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-202">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="839bb-203">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="839bb-203">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="839bb-204">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="839bb-204">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="839bb-205">Заголовки Интернета</span><span class="sxs-lookup"><span data-stu-id="839bb-205">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="839bb-206">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="839bb-206">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="839bb-207">Добавлен новый объект, представляющий пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="839bb-207">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="839bb-208">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-208">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="839bb-209">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="839bb-209">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="839bb-210">Добавлено новое свойство, представляющее пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="839bb-210">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="839bb-211">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-211">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="839bb-212">Тема Office</span><span class="sxs-lookup"><span data-stu-id="839bb-212">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="839bb-213">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="839bb-213">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="839bb-214">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="839bb-214">Added ability to get Office theme.</span></span>

<span data-ttu-id="839bb-215">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-215">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="839bb-216">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="839bb-216">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="839bb-217">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="839bb-217">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="839bb-218">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="839bb-218">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="839bb-219">Единый вход</span><span class="sxs-lookup"><span data-stu-id="839bb-219">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="839bb-220">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="839bb-220">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="839bb-221">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="839bb-221">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="839bb-222">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="839bb-222">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="839bb-223">См. также</span><span class="sxs-lookup"><span data-stu-id="839bb-223">See also</span></span>

- [<span data-ttu-id="839bb-224">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="839bb-224">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="839bb-225">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="839bb-225">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="839bb-226">Начало работы</span><span class="sxs-lookup"><span data-stu-id="839bb-226">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="839bb-227">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="839bb-227">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
