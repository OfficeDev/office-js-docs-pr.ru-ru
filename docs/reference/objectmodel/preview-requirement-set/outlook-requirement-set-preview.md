---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 05/17/2019
localization_priority: Priority
ms.openlocfilehash: d97efe8bbdfdadb252190458960b4356e0c8a564
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337176"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="b275a-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="b275a-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="b275a-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="b275a-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b275a-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b275a-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="b275a-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="b275a-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="b275a-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="b275a-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="b275a-107">Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.</span><span class="sxs-lookup"><span data-stu-id="b275a-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="b275a-108">Также может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="b275a-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="b275a-109">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="b275a-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="b275a-110">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="b275a-110">Features in preview</span></span>

<span data-ttu-id="b275a-111">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="b275a-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="b275a-112">Вложения</span><span class="sxs-lookup"><span data-stu-id="b275a-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="b275a-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="b275a-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="b275a-114">Добавлен новый объект, представляющий содержимое вложения.</span><span class="sxs-lookup"><span data-stu-id="b275a-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="b275a-115">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-115">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="b275a-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="b275a-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="b275a-117">Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="b275a-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="b275a-118">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-118">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="b275a-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="b275a-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="b275a-120">Добавлен новый метод, позволяющий получить содержимое определенного вложения.</span><span class="sxs-lookup"><span data-stu-id="b275a-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="b275a-121">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-121">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="b275a-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="b275a-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="b275a-123">Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="b275a-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="b275a-124">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-124">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="b275a-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="b275a-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="b275a-126">Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.</span><span class="sxs-lookup"><span data-stu-id="b275a-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="b275a-127">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-127">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="b275a-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="b275a-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="b275a-129">Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.</span><span class="sxs-lookup"><span data-stu-id="b275a-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="b275a-130">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-130">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="b275a-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="b275a-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="b275a-132">Добавлено событие `AttachmentsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="b275a-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="b275a-133">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-133">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="b275a-134">Блокировка при отправке</span><span class="sxs-lookup"><span data-stu-id="b275a-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="b275a-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="b275a-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="b275a-136">Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="b275a-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="b275a-137">Это значение используется для отмены выполнения события.</span><span class="sxs-lookup"><span data-stu-id="b275a-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="b275a-138">**Доступно в** Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="b275a-138">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="categories"></a><span data-ttu-id="b275a-139">Категории</span><span class="sxs-lookup"><span data-stu-id="b275a-139">Categories</span></span>

<span data-ttu-id="b275a-140">В Outlook пользователь может группировать сообщения и встречи, используя категории для выделения их цветом.</span><span class="sxs-lookup"><span data-stu-id="b275a-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="b275a-141">Пользователь определяет категории в главном списке своего почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="b275a-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="b275a-142">Затем он может применить одну или несколько категорий к элементу.</span><span class="sxs-lookup"><span data-stu-id="b275a-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="b275a-143">Эта функция не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="b275a-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="b275a-144">Categories</span><span class="sxs-lookup"><span data-stu-id="b275a-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="b275a-145">Добавлен новый объект, представляющий категории элемента.</span><span class="sxs-lookup"><span data-stu-id="b275a-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="b275a-146">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="b275a-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="b275a-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="b275a-148">Добавлен новый объект, представляющий сведения о категории (ее имя и соответствующий цвет).</span><span class="sxs-lookup"><span data-stu-id="b275a-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="b275a-149">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="b275a-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="b275a-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="b275a-151">Добавлен новый объект, представляющий главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="b275a-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="b275a-152">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="b275a-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="b275a-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="b275a-154">Добавлено новое свойство, представляющее главный список категорий для почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="b275a-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="b275a-155">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="b275a-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="b275a-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="b275a-157">Добавлено новое свойство, представляющее набор категорий для элемента.</span><span class="sxs-lookup"><span data-stu-id="b275a-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="b275a-158">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="b275a-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="b275a-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="b275a-160">Добавлено новое перечисление, указывающее цвета, доступные для сопоставления с категориями.</span><span class="sxs-lookup"><span data-stu-id="b275a-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="b275a-161">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="b275a-162">Делегированный доступ</span><span class="sxs-lookup"><span data-stu-id="b275a-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="b275a-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="b275a-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="b275a-164">Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="b275a-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="b275a-165">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-165">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="b275a-166">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="b275a-166">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="b275a-167">Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="b275a-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="b275a-168">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-168">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="b275a-169">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="b275a-169">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="b275a-170">Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.</span><span class="sxs-lookup"><span data-stu-id="b275a-170">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="b275a-171">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-171">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="b275a-172">Элемент манифеста SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="b275a-172">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="b275a-173">К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент.</span><span class="sxs-lookup"><span data-stu-id="b275a-173">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="b275a-174">Он определяет, доступна ли надстройка в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="b275a-174">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="b275a-175">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-175">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="b275a-176">Расширенные функции расположения</span><span class="sxs-lookup"><span data-stu-id="b275a-176">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="b275a-177">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="b275a-177">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="b275a-178">Добавлен новый объект, представляющий набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="b275a-178">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="b275a-179">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-179">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="b275a-180">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="b275a-180">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="b275a-181">Добавлен новый объект, представляющий расположение.</span><span class="sxs-lookup"><span data-stu-id="b275a-181">Added a new object that represents a location.</span></span> <span data-ttu-id="b275a-182">Только для чтения.</span><span class="sxs-lookup"><span data-stu-id="b275a-182">Read only.</span></span>

<span data-ttu-id="b275a-183">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-183">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="b275a-184">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="b275a-184">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="b275a-185">Добавлен новый объект, представляющий идентификатор расположения.</span><span class="sxs-lookup"><span data-stu-id="b275a-185">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="b275a-186">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-186">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="b275a-187">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="b275a-187">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="b275a-188">Добавлено новое свойство, представляющее набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="b275a-188">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="b275a-189">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-189">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="b275a-190">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="b275a-190">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="b275a-191">Добавлено новое перечисление, которое определяет тип расположения встречи.</span><span class="sxs-lookup"><span data-stu-id="b275a-191">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="b275a-192">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-192">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="b275a-193">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="b275a-193">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="b275a-194">Добавлено событие `EnhancedLocationsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="b275a-194">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="b275a-195">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-195">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="b275a-196">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="b275a-196">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="b275a-197">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="b275a-197">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="b275a-198">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="b275a-198">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="b275a-199">**Доступно в** Outlook для Windows (подключенный к Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="b275a-199">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="b275a-200">Заголовки Интернета</span><span class="sxs-lookup"><span data-stu-id="b275a-200">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="b275a-201">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="b275a-201">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="b275a-202">Добавлен новый объект, представляющий заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="b275a-202">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="b275a-203">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-203">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="b275a-204">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="b275a-204">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="b275a-205">Добавлено новое свойство, представляющее заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="b275a-205">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="b275a-206">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="b275a-207">Тема Office</span><span class="sxs-lookup"><span data-stu-id="b275a-207">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="b275a-208">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="b275a-208">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="b275a-209">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="b275a-209">Added ability to get Office theme.</span></span>

<span data-ttu-id="b275a-210">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-210">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="b275a-211">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="b275a-211">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="b275a-212">Добавлено событие `OfficeThemeChanged` в объект `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="b275a-212">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="b275a-213">**Доступно в** Outlook для Windows (подключенный к Office 365)</span><span class="sxs-lookup"><span data-stu-id="b275a-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="b275a-214">Единый вход</span><span class="sxs-lookup"><span data-stu-id="b275a-214">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="b275a-215">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b275a-215">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="b275a-216">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="b275a-216">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="b275a-217">**Доступно в** Outlook для Windows (подключенный к Office 365), Outlook для Mac (подключенный к Office 365), Outlook в Интернете (Outlook.com и подключенный к Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="b275a-217">**Available in**: Outlook on Windows (connected to Office 365), Outlook for Mac (connected to Office 365), Outlook on the web (Outlook.com and connected to Office 365), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="b275a-218">См. также</span><span class="sxs-lookup"><span data-stu-id="b275a-218">See also</span></span>

- [<span data-ttu-id="b275a-219">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="b275a-219">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="b275a-220">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="b275a-220">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="b275a-221">Начало работы</span><span class="sxs-lookup"><span data-stu-id="b275a-221">Get started</span></span>](/outlook/add-ins/quick-start)
