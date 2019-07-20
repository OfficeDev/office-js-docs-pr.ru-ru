---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: f2030ca40a83694c1c1c92ca4b956a9294cb3e0a
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804992"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="e2d02-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="e2d02-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="e2d02-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2d02-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d02-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e2d02-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="e2d02-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="e2d02-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="e2d02-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="e2d02-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="e2d02-107">Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.</span><span class="sxs-lookup"><span data-stu-id="e2d02-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="e2d02-108">Также может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="e2d02-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="e2d02-109">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="e2d02-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="e2d02-110">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="e2d02-110">Features in preview</span></span>

<span data-ttu-id="e2d02-111">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="e2d02-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="e2d02-112">Вложения</span><span class="sxs-lookup"><span data-stu-id="e2d02-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="e2d02-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="e2d02-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="e2d02-114">Добавлен новый объект, представляющий содержимое вложения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="e2d02-115">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-115">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="e2d02-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e2d02-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="e2d02-117">Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="e2d02-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="e2d02-118">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-118">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="e2d02-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="e2d02-120">Добавлен новый метод, позволяющий получить содержимое определенного вложения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="e2d02-121">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-121">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="e2d02-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="e2d02-123">Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="e2d02-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="e2d02-124">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-124">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="e2d02-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="e2d02-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="e2d02-126">Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="e2d02-127">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-127">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="e2d02-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="e2d02-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="e2d02-129">Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.</span><span class="sxs-lookup"><span data-stu-id="e2d02-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="e2d02-130">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-130">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e2d02-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="e2d02-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e2d02-132">Добавлено событие `AttachmentsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="e2d02-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="e2d02-133">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-133">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="e2d02-134">Блокировка при отправке</span><span class="sxs-lookup"><span data-stu-id="e2d02-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="e2d02-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="e2d02-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="e2d02-136">Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="e2d02-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="e2d02-137">Это значение используется для отмены выполнения события.</span><span class="sxs-lookup"><span data-stu-id="e2d02-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="e2d02-138">**Доступно в** Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-138">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="categories"></a><span data-ttu-id="e2d02-139">Категории</span><span class="sxs-lookup"><span data-stu-id="e2d02-139">Categories</span></span>

<span data-ttu-id="e2d02-140">В Outlook пользователь может группировать сообщения и встречи, используя категории для выделения их цветом.</span><span class="sxs-lookup"><span data-stu-id="e2d02-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="e2d02-141">Пользователь определяет категории в главном списке своего почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="e2d02-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="e2d02-142">Затем он может применить одну или несколько категорий к элементу.</span><span class="sxs-lookup"><span data-stu-id="e2d02-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="e2d02-143">Эта возможность не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="e2d02-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="e2d02-144">Categories</span><span class="sxs-lookup"><span data-stu-id="e2d02-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="e2d02-145">Добавлен новый объект, представляющий категории элемента.</span><span class="sxs-lookup"><span data-stu-id="e2d02-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="e2d02-146">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="e2d02-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="e2d02-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="e2d02-148">Добавлен новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).</span><span class="sxs-lookup"><span data-stu-id="e2d02-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="e2d02-149">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="e2d02-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="e2d02-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="e2d02-151">Добавлен новый объект, представляющий главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="e2d02-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="e2d02-152">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="e2d02-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="e2d02-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="e2d02-154">Добавлено новое свойство, представляющее главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="e2d02-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="e2d02-155">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="e2d02-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="e2d02-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="e2d02-157">Добавлено новое свойство, представляющее набор категорий для элемента.</span><span class="sxs-lookup"><span data-stu-id="e2d02-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="e2d02-158">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="e2d02-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="e2d02-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="e2d02-160">Добавлено новое перечисление, указывающее цвета, доступные для сопоставления с категориями.</span><span class="sxs-lookup"><span data-stu-id="e2d02-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="e2d02-161">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="e2d02-162">Делегированный доступ</span><span class="sxs-lookup"><span data-stu-id="e2d02-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="e2d02-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="e2d02-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="e2d02-164">Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="e2d02-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="e2d02-165">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-165">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="e2d02-166">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-166">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="e2d02-167">Добавлен новый метод, получающий идентификатор сохраненного элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="e2d02-168">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-168">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="e2d02-169">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-169">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="e2d02-170">Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-170">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="e2d02-171">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-171">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="e2d02-172">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="e2d02-172">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="e2d02-173">Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.</span><span class="sxs-lookup"><span data-stu-id="e2d02-173">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="e2d02-174">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-174">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="e2d02-175">Элемент манифеста SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e2d02-175">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="e2d02-176">К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент.</span><span class="sxs-lookup"><span data-stu-id="e2d02-176">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="e2d02-177">Он определяет, доступна ли надстройка в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="e2d02-177">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="e2d02-178">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-178">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="e2d02-179">Расширенные функции расположения</span><span class="sxs-lookup"><span data-stu-id="e2d02-179">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="e2d02-180">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e2d02-180">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="e2d02-181">Добавлен новый объект, представляющий набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="e2d02-181">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e2d02-182">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-182">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="e2d02-183">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="e2d02-183">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="e2d02-184">Добавлен новый объект, представляющий расположение.</span><span class="sxs-lookup"><span data-stu-id="e2d02-184">Added a new object that represents a location.</span></span> <span data-ttu-id="e2d02-185">Только для чтения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-185">Read only.</span></span>

<span data-ttu-id="e2d02-186">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-186">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="e2d02-187">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="e2d02-187">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="e2d02-188">Добавлен новый объект, представляющий идентификатор расположения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-188">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="e2d02-189">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-189">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="e2d02-190">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e2d02-190">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="e2d02-191">Добавлено новое свойство, представляющее набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="e2d02-191">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e2d02-192">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-192">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="e2d02-193">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="e2d02-193">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="e2d02-194">Добавлено новое перечисление, которое определяет тип расположения встречи.</span><span class="sxs-lookup"><span data-stu-id="e2d02-194">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="e2d02-195">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-195">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e2d02-196">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="e2d02-196">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e2d02-197">Добавлено событие `EnhancedLocationsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="e2d02-197">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="e2d02-198">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-198">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="e2d02-199">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="e2d02-199">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="e2d02-200">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-200">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="e2d02-201">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="e2d02-201">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="e2d02-202">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-202">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="e2d02-203">Заголовки Интернета</span><span class="sxs-lookup"><span data-stu-id="e2d02-203">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="e2d02-204">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e2d02-204">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="e2d02-205">Добавлен новый объект, представляющий пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-205">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="e2d02-206">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="e2d02-207">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="e2d02-207">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="e2d02-208">Добавлено новое свойство, представляющее пользовательские заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="e2d02-208">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="e2d02-209">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-209">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="e2d02-210">Тема Office</span><span class="sxs-lookup"><span data-stu-id="e2d02-210">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="e2d02-211">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="e2d02-211">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="e2d02-212">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="e2d02-212">Added ability to get Office theme.</span></span>

<span data-ttu-id="e2d02-213">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e2d02-214">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="e2d02-214">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e2d02-215">Добавлено событие `OfficeThemeChanged` в объект `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="e2d02-215">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="e2d02-216">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="e2d02-216">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="e2d02-217">Единый вход</span><span class="sxs-lookup"><span data-stu-id="e2d02-217">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="e2d02-218">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e2d02-218">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="e2d02-219">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e2d02-219">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="e2d02-220">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="e2d02-220">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="e2d02-221">См. также</span><span class="sxs-lookup"><span data-stu-id="e2d02-221">See also</span></span>

- [<span data-ttu-id="e2d02-222">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="e2d02-222">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="e2d02-223">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="e2d02-223">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e2d02-224">Начало работы</span><span class="sxs-lookup"><span data-stu-id="e2d02-224">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="e2d02-225">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="e2d02-225">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
