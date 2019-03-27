---
title: Пространство имен Office — Предварительная версия набора требований
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: e426ea87c14c4ad21ebdbfd3df05988ba848b906
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872237"
---
# <a name="office"></a><span data-ttu-id="ad760-102">Office</span><span class="sxs-lookup"><span data-stu-id="ad760-102">Office</span></span>

<span data-ttu-id="ad760-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ad760-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ad760-105">Требования</span><span class="sxs-lookup"><span data-stu-id="ad760-105">Requirements</span></span>

|<span data-ttu-id="ad760-106">Требование</span><span class="sxs-lookup"><span data-stu-id="ad760-106">Requirement</span></span>| <span data-ttu-id="ad760-107">Значение</span><span class="sxs-lookup"><span data-stu-id="ad760-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad760-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad760-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad760-109">1.0</span><span class="sxs-lookup"><span data-stu-id="ad760-109">1.0</span></span>|
|[<span data-ttu-id="ad760-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad760-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad760-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad760-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ad760-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ad760-112">Members and methods</span></span>

| <span data-ttu-id="ad760-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="ad760-113">Member</span></span> | <span data-ttu-id="ad760-114">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ad760-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ad760-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ad760-116">Member</span><span class="sxs-lookup"><span data-stu-id="ad760-116">Member</span></span> |
| [<span data-ttu-id="ad760-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ad760-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ad760-118">Member</span><span class="sxs-lookup"><span data-stu-id="ad760-118">Member</span></span> |
| [<span data-ttu-id="ad760-119">EventType</span><span class="sxs-lookup"><span data-stu-id="ad760-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ad760-120">Member</span><span class="sxs-lookup"><span data-stu-id="ad760-120">Member</span></span> |
| [<span data-ttu-id="ad760-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ad760-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ad760-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="ad760-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="ad760-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ad760-123">Namespaces</span></span>

<span data-ttu-id="ad760-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="ad760-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="ad760-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="ad760-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="ad760-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="ad760-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="ad760-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="ad760-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="ad760-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="ad760-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ad760-129">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-129">Type</span></span>

*   <span data-ttu-id="ad760-130">String</span><span class="sxs-lookup"><span data-stu-id="ad760-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad760-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad760-131">Properties:</span></span>

|<span data-ttu-id="ad760-132">Имя</span><span class="sxs-lookup"><span data-stu-id="ad760-132">Name</span></span>| <span data-ttu-id="ad760-133">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-133">Type</span></span>| <span data-ttu-id="ad760-134">Описание</span><span class="sxs-lookup"><span data-stu-id="ad760-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ad760-135">String</span><span class="sxs-lookup"><span data-stu-id="ad760-135">String</span></span>|<span data-ttu-id="ad760-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="ad760-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ad760-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="ad760-137">String</span></span>|<span data-ttu-id="ad760-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="ad760-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad760-139">Требования</span><span class="sxs-lookup"><span data-stu-id="ad760-139">Requirements</span></span>

|<span data-ttu-id="ad760-140">Требование</span><span class="sxs-lookup"><span data-stu-id="ad760-140">Requirement</span></span>| <span data-ttu-id="ad760-141">Значение</span><span class="sxs-lookup"><span data-stu-id="ad760-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad760-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad760-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad760-143">1.0</span><span class="sxs-lookup"><span data-stu-id="ad760-143">1.0</span></span>|
|[<span data-ttu-id="ad760-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad760-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad760-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad760-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="ad760-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="ad760-146">CoercionType :String</span></span>

<span data-ttu-id="ad760-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ad760-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad760-148">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-148">Type</span></span>

*   <span data-ttu-id="ad760-149">String</span><span class="sxs-lookup"><span data-stu-id="ad760-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad760-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad760-150">Properties:</span></span>

|<span data-ttu-id="ad760-151">Имя</span><span class="sxs-lookup"><span data-stu-id="ad760-151">Name</span></span>| <span data-ttu-id="ad760-152">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-152">Type</span></span>| <span data-ttu-id="ad760-153">Описание</span><span class="sxs-lookup"><span data-stu-id="ad760-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ad760-154">String</span><span class="sxs-lookup"><span data-stu-id="ad760-154">String</span></span>|<span data-ttu-id="ad760-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="ad760-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ad760-156">String</span><span class="sxs-lookup"><span data-stu-id="ad760-156">String</span></span>|<span data-ttu-id="ad760-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="ad760-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad760-158">Требования</span><span class="sxs-lookup"><span data-stu-id="ad760-158">Requirements</span></span>

|<span data-ttu-id="ad760-159">Требование</span><span class="sxs-lookup"><span data-stu-id="ad760-159">Requirement</span></span>| <span data-ttu-id="ad760-160">Значение</span><span class="sxs-lookup"><span data-stu-id="ad760-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad760-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad760-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad760-162">1.0</span><span class="sxs-lookup"><span data-stu-id="ad760-162">1.0</span></span>|
|[<span data-ttu-id="ad760-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad760-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad760-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad760-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="ad760-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="ad760-165">EventType :String</span></span>

<span data-ttu-id="ad760-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="ad760-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ad760-167">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-167">Type</span></span>

*   <span data-ttu-id="ad760-168">String</span><span class="sxs-lookup"><span data-stu-id="ad760-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad760-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad760-169">Properties:</span></span>

| <span data-ttu-id="ad760-170">Имя</span><span class="sxs-lookup"><span data-stu-id="ad760-170">Name</span></span> | <span data-ttu-id="ad760-171">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-171">Type</span></span> | <span data-ttu-id="ad760-172">Описание</span><span class="sxs-lookup"><span data-stu-id="ad760-172">Description</span></span> | <span data-ttu-id="ad760-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="ad760-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="ad760-174">String</span><span class="sxs-lookup"><span data-stu-id="ad760-174">String</span></span> | <span data-ttu-id="ad760-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="ad760-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ad760-176">1.7</span><span class="sxs-lookup"><span data-stu-id="ad760-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="ad760-177">String</span><span class="sxs-lookup"><span data-stu-id="ad760-177">String</span></span> | <span data-ttu-id="ad760-178">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="ad760-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="ad760-179">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ad760-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="ad760-180">String</span><span class="sxs-lookup"><span data-stu-id="ad760-180">String</span></span> | <span data-ttu-id="ad760-181">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="ad760-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="ad760-182">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ad760-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="ad760-183">String</span><span class="sxs-lookup"><span data-stu-id="ad760-183">String</span></span> | <span data-ttu-id="ad760-184">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="ad760-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ad760-185">1.5</span><span class="sxs-lookup"><span data-stu-id="ad760-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="ad760-186">String</span><span class="sxs-lookup"><span data-stu-id="ad760-186">String</span></span> | <span data-ttu-id="ad760-187">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="ad760-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="ad760-188">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="ad760-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ad760-189">String</span><span class="sxs-lookup"><span data-stu-id="ad760-189">String</span></span> | <span data-ttu-id="ad760-190">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="ad760-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ad760-191">1.7</span><span class="sxs-lookup"><span data-stu-id="ad760-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ad760-192">String</span><span class="sxs-lookup"><span data-stu-id="ad760-192">String</span></span> | <span data-ttu-id="ad760-193">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="ad760-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ad760-194">1.7</span><span class="sxs-lookup"><span data-stu-id="ad760-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ad760-195">Требования</span><span class="sxs-lookup"><span data-stu-id="ad760-195">Requirements</span></span>

|<span data-ttu-id="ad760-196">Требование</span><span class="sxs-lookup"><span data-stu-id="ad760-196">Requirement</span></span>| <span data-ttu-id="ad760-197">Значение</span><span class="sxs-lookup"><span data-stu-id="ad760-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad760-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ad760-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad760-199">1.5</span><span class="sxs-lookup"><span data-stu-id="ad760-199">1.5</span></span> |
|[<span data-ttu-id="ad760-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad760-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad760-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad760-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="ad760-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="ad760-202">SourceProperty :String</span></span>

<span data-ttu-id="ad760-203">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="ad760-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ad760-204">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-204">Type</span></span>

*   <span data-ttu-id="ad760-205">String</span><span class="sxs-lookup"><span data-stu-id="ad760-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ad760-206">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ad760-206">Properties:</span></span>

|<span data-ttu-id="ad760-207">Имя</span><span class="sxs-lookup"><span data-stu-id="ad760-207">Name</span></span>| <span data-ttu-id="ad760-208">Тип</span><span class="sxs-lookup"><span data-stu-id="ad760-208">Type</span></span>| <span data-ttu-id="ad760-209">Описание</span><span class="sxs-lookup"><span data-stu-id="ad760-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ad760-210">String</span><span class="sxs-lookup"><span data-stu-id="ad760-210">String</span></span>|<span data-ttu-id="ad760-211">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad760-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ad760-212">String</span><span class="sxs-lookup"><span data-stu-id="ad760-212">String</span></span>|<span data-ttu-id="ad760-213">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="ad760-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ad760-214">Требования</span><span class="sxs-lookup"><span data-stu-id="ad760-214">Requirements</span></span>

|<span data-ttu-id="ad760-215">Требование</span><span class="sxs-lookup"><span data-stu-id="ad760-215">Requirement</span></span>| <span data-ttu-id="ad760-216">Значение</span><span class="sxs-lookup"><span data-stu-id="ad760-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="ad760-217">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ad760-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ad760-218">1.0</span><span class="sxs-lookup"><span data-stu-id="ad760-218">1.0</span></span>|
|[<span data-ttu-id="ad760-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ad760-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ad760-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ad760-220">Compose or Read</span></span>|
