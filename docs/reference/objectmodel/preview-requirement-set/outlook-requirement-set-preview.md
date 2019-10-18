---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: aa3f46c505e8c87508699f6e84194272ee4d13bb
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696458"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="34990-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="34990-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="34990-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="34990-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="34990-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="34990-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="34990-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="34990-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="34990-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="34990-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="34990-107">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="34990-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="34990-108">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="34990-108">Features in preview</span></span>

<span data-ttu-id="34990-109">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="34990-109">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="34990-110">Вложения</span><span class="sxs-lookup"><span data-stu-id="34990-110">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="34990-111">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="34990-111">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="34990-112">Добавлен новый объект, представляющий содержимое вложения.</span><span class="sxs-lookup"><span data-stu-id="34990-112">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="34990-113">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="34990-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="34990-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="34990-115">Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="34990-115">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="34990-116">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-116">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="34990-117">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="34990-117">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="34990-118">Добавлен новый метод, позволяющий получить содержимое определенного вложения.</span><span class="sxs-lookup"><span data-stu-id="34990-118">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="34990-119">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-119">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="34990-120">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="34990-120">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="34990-121">Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="34990-121">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="34990-122">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-122">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="34990-123">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="34990-123">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="34990-124">Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.</span><span class="sxs-lookup"><span data-stu-id="34990-124">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="34990-125">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="34990-126">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="34990-126">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="34990-127">Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.</span><span class="sxs-lookup"><span data-stu-id="34990-127">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="34990-128">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="34990-129">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="34990-129">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="34990-130">Добавлено событие `AttachmentsChanged` для объекта `Item`.</span><span class="sxs-lookup"><span data-stu-id="34990-130">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="34990-131">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="block-on-send"></a><span data-ttu-id="34990-132">Блокировка при отправке</span><span class="sxs-lookup"><span data-stu-id="34990-132">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="34990-133">Event.completed</span><span class="sxs-lookup"><span data-stu-id="34990-133">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="34990-134">Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="34990-134">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="34990-135">Это значение используется для отмены выполнения события.</span><span class="sxs-lookup"><span data-stu-id="34990-135">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="34990-136">**Доступно в** Outlook в Интернете (классическая версия), Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-136">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="categories"></a><span data-ttu-id="34990-137">Категории</span><span class="sxs-lookup"><span data-stu-id="34990-137">Categories</span></span>

<span data-ttu-id="34990-138">В Outlook пользователь может группировать сообщения и встречи, используя категории для выделения их цветом.</span><span class="sxs-lookup"><span data-stu-id="34990-138">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="34990-139">Пользователь определяет категории в главном списке своего почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="34990-139">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="34990-140">Затем он может применить одну или несколько категорий к элементу.</span><span class="sxs-lookup"><span data-stu-id="34990-140">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="34990-141">Эта возможность не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="34990-141">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="34990-142">Categories</span><span class="sxs-lookup"><span data-stu-id="34990-142">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="34990-143">Добавлен новый объект, представляющий категории элемента.</span><span class="sxs-lookup"><span data-stu-id="34990-143">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="34990-144">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="34990-145">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="34990-145">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="34990-146">Добавлен новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).</span><span class="sxs-lookup"><span data-stu-id="34990-146">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="34990-147">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-147">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="34990-148">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="34990-148">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="34990-149">Добавлен новый объект, представляющий главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="34990-149">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="34990-150">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="34990-151">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="34990-151">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="34990-152">Добавлено новое свойство, представляющее главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="34990-152">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="34990-153">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="34990-154">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="34990-154">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="34990-155">Добавлено новое свойство, представляющее набор категорий для элемента.</span><span class="sxs-lookup"><span data-stu-id="34990-155">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="34990-156">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="34990-157">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="34990-157">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="34990-158">Добавлено новое перечисление, указывающее цвета, доступные для сопоставления с категориями.</span><span class="sxs-lookup"><span data-stu-id="34990-158">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="34990-159">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="delegate-access"></a><span data-ttu-id="34990-160">Делегированный доступ</span><span class="sxs-lookup"><span data-stu-id="34990-160">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="34990-161">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="34990-161">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="34990-162">Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="34990-162">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="34990-163">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="34990-164">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="34990-164">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="34990-165">Добавлен новый метод, получающий идентификатор сохраненного элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="34990-165">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="34990-166">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-166">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="34990-167">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="34990-167">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="34990-168">Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="34990-168">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="34990-169">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="34990-170">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="34990-170">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="34990-171">Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.</span><span class="sxs-lookup"><span data-stu-id="34990-171">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="34990-172">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-172">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="34990-173">Элемент манифеста SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="34990-173">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="34990-174">К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент.</span><span class="sxs-lookup"><span data-stu-id="34990-174">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="34990-175">Он определяет, доступна ли надстройка в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="34990-175">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="34990-176">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="enhanced-location"></a><span data-ttu-id="34990-177">Расширенные функции расположения</span><span class="sxs-lookup"><span data-stu-id="34990-177">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="34990-178">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="34990-178">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="34990-179">Добавлен новый объект, представляющий набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="34990-179">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="34990-180">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="34990-181">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="34990-181">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="34990-182">Добавлен новый объект, представляющий расположение.</span><span class="sxs-lookup"><span data-stu-id="34990-182">Added a new object that represents a location.</span></span> <span data-ttu-id="34990-183">Только для чтения.</span><span class="sxs-lookup"><span data-stu-id="34990-183">Read only.</span></span>

<span data-ttu-id="34990-184">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="34990-185">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="34990-185">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="34990-186">Добавлен новый объект, представляющий идентификатор расположения.</span><span class="sxs-lookup"><span data-stu-id="34990-186">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="34990-187">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="34990-188">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="34990-188">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="34990-189">Добавлено новое свойство, представляющее набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="34990-189">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="34990-190">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-190">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="34990-191">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="34990-191">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="34990-192">Добавлено новое перечисление, которое определяет тип расположения встречи.</span><span class="sxs-lookup"><span data-stu-id="34990-192">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="34990-193">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-193">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="34990-194">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="34990-194">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="34990-195">Добавлено событие `EnhancedLocationsChanged` для объекта `Item`.</span><span class="sxs-lookup"><span data-stu-id="34990-195">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="34990-196">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-196">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="34990-197">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="34990-197">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="34990-198">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="34990-198">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="34990-199">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="34990-199">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="34990-200">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="34990-200">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

<br>

---

### <a name="internet-headers"></a><span data-ttu-id="34990-201">Заголовки Интернета</span><span class="sxs-lookup"><span data-stu-id="34990-201">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="34990-202">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="34990-202">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="34990-203">Добавлен новый объект, представляющий пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="34990-203">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="34990-204">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-204">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="34990-205">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="34990-205">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="34990-206">Добавлено новое свойство, представляющее пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="34990-206">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="34990-207">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-207">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="office-theme"></a><span data-ttu-id="34990-208">Тема Office</span><span class="sxs-lookup"><span data-stu-id="34990-208">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="34990-209">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="34990-209">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="34990-210">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="34990-210">Added ability to get Office theme.</span></span>

<span data-ttu-id="34990-211">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-211">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="34990-212">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="34990-212">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="34990-213">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="34990-213">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="34990-214">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="34990-214">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="34990-215">Единый вход</span><span class="sxs-lookup"><span data-stu-id="34990-215">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="34990-216">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="34990-216">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="34990-217">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="34990-217">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="34990-218">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="34990-218">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="34990-219">См. также</span><span class="sxs-lookup"><span data-stu-id="34990-219">See also</span></span>

- [<span data-ttu-id="34990-220">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="34990-220">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="34990-221">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="34990-221">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="34990-222">Начало работы</span><span class="sxs-lookup"><span data-stu-id="34990-222">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="34990-223">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="34990-223">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
