---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512862"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="ab6e3-102">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="ab6e3-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="ab6e3-103">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="ab6e3-104">Эта документация относится к **предварительной версии** [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ab6e3-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="ab6e3-105">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="ab6e3-106">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="ab6e3-107">Прежде чем использовать методы и свойства, добавленные в этом наборе обязательных элементов, следует отдельно проверять их на доступность.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="ab6e3-108">Также может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="ab6e3-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="ab6e3-109">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="ab6e3-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="ab6e3-110">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="ab6e3-110">Features in preview</span></span>

<span data-ttu-id="ab6e3-111">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="ab6e3-112">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="ab6e3-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="ab6e3-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="ab6e3-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="ab6e3-114">Добавлен новый необязательный параметр `options`, представляющий собой словарь с одним допустимым значением `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-114">Event.completed - A new optional parameter , which is a dictionary with one valid value .</span></span> <span data-ttu-id="ab6e3-115">Это значение используется для отмены выполнения события.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="ab6e3-116">**Доступно в** Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-116">**Available in**: Outlook on the web (Classic)</span></span>

### <a name="attachments"></a><span data-ttu-id="ab6e3-117">Вложения</span><span class="sxs-lookup"><span data-stu-id="ab6e3-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="ab6e3-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="ab6e3-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="ab6e3-119">Добавлен новый объект, представляющий содержимое вложения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-119">AttachmentContent - Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="ab6e3-120">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-120">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="ab6e3-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="ab6e3-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="ab6e3-122">Добавлен новый метод, который позволяет вложить в сообщение или встречу файл, представленный в виде строки в кодировке base64.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-122">Office.context.mailbox.item.addFileAttachmentFromBase64Async - Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="ab6e3-123">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-123">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="ab6e3-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="ab6e3-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

<span data-ttu-id="ab6e3-125">Добавлен новый метод, позволяющий получить содержимое определенного вложения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-125">Office.context.mailbox.item.getAttachmentContentAsync - Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="ab6e3-126">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-126">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[<span data-ttu-id="ab6e3-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="ab6e3-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

<span data-ttu-id="ab6e3-128">Добавлен новый метод, который получает вложенные в элемент объекты в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-128">Office.context.mailbox.item.getAttachmentsAsync - Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="ab6e3-129">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-129">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="ab6e3-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="ab6e3-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="ab6e3-131">Добавлено новое перечисление, в котором указывается форматирование, применяемое к содержимому вложения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-131">Office.MailboxEnums.AttachmentContentFormat - Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="ab6e3-132">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-132">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="ab6e3-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="ab6e3-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="ab6e3-134">Добавлено новое перечисление, в котором указывается, добавлено вложение в элемент или удалено из него.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-134">Office.MailboxEnums.AttachmentStatus - Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="ab6e3-135">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-135">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="ab6e3-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="ab6e3-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="ab6e3-137">Добавлено событие `AttachmentsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="ab6e3-138">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-138">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="delegate-access"></a><span data-ttu-id="ab6e3-139">Делегированный доступ</span><span class="sxs-lookup"><span data-stu-id="ab6e3-139">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="ab6e3-140">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="ab6e3-140">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="ab6e3-141">Добавлен новый объект, который представляет свойства элемента встречи или сообщения в общей папке, календаре или почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-141">SharedProperties - Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="ab6e3-142">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-142">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="ab6e3-143">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="ab6e3-143">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="ab6e3-144">Добавлен новый метод, позволяющий получить объект, который представляет свойства sharedProperties элемента встречи или сообщения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-144">Office.context.mailbox.item.getSharedPropertiesAsync - Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="ab6e3-145">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-145">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="ab6e3-146">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="ab6e3-146">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="ab6e3-147">Добавлено перечисление нового битового флага, в котором указываются разрешения на делегирование.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-147">Office.MailboxEnums.DelegatePermissions - Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="ab6e3-148">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-148">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="ab6e3-149">Элемент манифеста SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="ab6e3-149">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="ab6e3-150">К элементу манифеста [DesktopFormFactor](../../manifest/desktopformfactor.md) добавлен дочерний элемент.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-150">SupportsSharedFolders manifest element - Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="ab6e3-151">Он определяет, доступна ли надстройка в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-151">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="ab6e3-152">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-152">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="enhanced-location"></a><span data-ttu-id="ab6e3-153">Расширенные функции расположения</span><span class="sxs-lookup"><span data-stu-id="ab6e3-153">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="ab6e3-154">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="ab6e3-154">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="ab6e3-155">Добавлен новый объект, представляющий набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-155">EnhancedLocation - Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="ab6e3-156">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-156">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="ab6e3-157">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="ab6e3-157">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="ab6e3-158">Добавлен новый объект, представляющий расположение.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-158">LocationDetails - Added a new object that represents a location.</span></span> <span data-ttu-id="ab6e3-159">Только для чтения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-159">Read only.</span></span>

<span data-ttu-id="ab6e3-160">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-160">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="ab6e3-161">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="ab6e3-161">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="ab6e3-162">Добавлен новый объект, представляющий идентификатор расположения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-162">LocationIdentifier - Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="ab6e3-163">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-163">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="ab6e3-164">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="ab6e3-164">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

<span data-ttu-id="ab6e3-165">Добавлено новое свойство, представляющее набор расположений для встречи.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-165">Office.context.mailbox.item.enhancedLocation - Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="ab6e3-166">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-166">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="ab6e3-167">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="ab6e3-167">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="ab6e3-168">Добавлено новое перечисление, которое определяет тип расположения встречи.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-168">Office.MailboxEnums.LocationType - Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="ab6e3-169">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-169">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="ab6e3-170">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="ab6e3-170">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="ab6e3-171">Добавлено событие `EnhancedLocationsChanged` в объект `Item`.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-171">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="ab6e3-172">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-172">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="ab6e3-173">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="ab6e3-173">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="ab6e3-174">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="ab6e3-174">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="ab6e3-175">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="ab6e3-175">Office.context.mailbox.item.getInitializationContextAsync - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="ab6e3-176">**Доступно в** Office 2019 для Windows (подписка на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-176">**Available in**: Office 2019 for Windows (Office 365 subscription), Outlook on the web (Classic)</span></span>

### <a name="internet-headers"></a><span data-ttu-id="ab6e3-177">Заголовки Интернета</span><span class="sxs-lookup"><span data-stu-id="ab6e3-177">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="ab6e3-178">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="ab6e3-178">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="ab6e3-179">Добавлен новый объект, представляющий заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-179">InternetHeaders - Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="ab6e3-180">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-180">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="ab6e3-181">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="ab6e3-181">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

<span data-ttu-id="ab6e3-182">Добавлено новое свойство, представляющее заголовки Интернета в элементе сообщения.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-182">Office.context.mailbox.item.internetHeaders - Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="ab6e3-183">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-183">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="office-theme"></a><span data-ttu-id="ab6e3-184">Тема Office</span><span class="sxs-lookup"><span data-stu-id="ab6e3-184">Office Theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="ab6e3-185">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="ab6e3-185">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="ab6e3-186">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="ab6e3-187">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-187">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="ab6e3-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="ab6e3-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="ab6e3-189">Добавлено событие `OfficeThemeChanged` в объект `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="ab6e3-190">**Доступно в** Outlook 2019 для Windows (подписка на Office 365)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-190">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="sso"></a><span data-ttu-id="ab6e3-191">Единый вход</span><span class="sxs-lookup"><span data-stu-id="ab6e3-191">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="ab6e3-192">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="ab6e3-192">Office.context.auth.getAccessTokenAsync</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="ab6e3-193">Добавлена возможность доступа к `getAccessTokenAsync`, что позволяет надстройкам [получать маркер доступа](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="ab6e3-193">Office.context.auth.getAccessTokenAsync - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="ab6e3-194">**Доступно в** Outlook 2019 для Windows (подписка на Office 365), Outlook 2019 для Mac, Outlook в Интернете (Office 365 и Outlook.com), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="ab6e3-194">**Available in**: Outlook 2019 for Windows (Office 365 subscription), Outlook 2019 for Mac, Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="ab6e3-195">См. также</span><span class="sxs-lookup"><span data-stu-id="ab6e3-195">See also</span></span>

- [<span data-ttu-id="ab6e3-196">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="ab6e3-196">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="ab6e3-197">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="ab6e3-197">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="ab6e3-198">Начало работы</span><span class="sxs-lookup"><span data-stu-id="ab6e3-198">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
