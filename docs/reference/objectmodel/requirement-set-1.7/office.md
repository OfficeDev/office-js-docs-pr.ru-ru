---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 533e997fc7f8be6eb6d3aefefaf023e8c7666af2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870529"
---
# <a name="office"></a><span data-ttu-id="156a1-102">Office</span><span class="sxs-lookup"><span data-stu-id="156a1-102">Office</span></span>

<span data-ttu-id="156a1-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="156a1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="156a1-105">Требования</span><span class="sxs-lookup"><span data-stu-id="156a1-105">Requirements</span></span>

|<span data-ttu-id="156a1-106">Требование</span><span class="sxs-lookup"><span data-stu-id="156a1-106">Requirement</span></span>| <span data-ttu-id="156a1-107">Значение</span><span class="sxs-lookup"><span data-stu-id="156a1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="156a1-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="156a1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156a1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="156a1-109">1.0</span></span>|
|[<span data-ttu-id="156a1-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="156a1-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156a1-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="156a1-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="156a1-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="156a1-112">Members and methods</span></span>

| <span data-ttu-id="156a1-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="156a1-113">Member</span></span> | <span data-ttu-id="156a1-114">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="156a1-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="156a1-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="156a1-116">Member</span><span class="sxs-lookup"><span data-stu-id="156a1-116">Member</span></span> |
| [<span data-ttu-id="156a1-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="156a1-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="156a1-118">Member</span><span class="sxs-lookup"><span data-stu-id="156a1-118">Member</span></span> |
| [<span data-ttu-id="156a1-119">EventType</span><span class="sxs-lookup"><span data-stu-id="156a1-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="156a1-120">Member</span><span class="sxs-lookup"><span data-stu-id="156a1-120">Member</span></span> |
| [<span data-ttu-id="156a1-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="156a1-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="156a1-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="156a1-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="156a1-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="156a1-123">Namespaces</span></span>

<span data-ttu-id="156a1-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="156a1-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="156a1-125">[MailboxEnums.](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="156a1-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="156a1-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="156a1-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="156a1-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="156a1-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="156a1-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="156a1-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="156a1-129">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-129">Type</span></span>

*   <span data-ttu-id="156a1-130">String</span><span class="sxs-lookup"><span data-stu-id="156a1-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156a1-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="156a1-131">Properties:</span></span>

|<span data-ttu-id="156a1-132">Имя</span><span class="sxs-lookup"><span data-stu-id="156a1-132">Name</span></span>| <span data-ttu-id="156a1-133">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-133">Type</span></span>| <span data-ttu-id="156a1-134">Описание</span><span class="sxs-lookup"><span data-stu-id="156a1-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="156a1-135">String</span><span class="sxs-lookup"><span data-stu-id="156a1-135">String</span></span>|<span data-ttu-id="156a1-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="156a1-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="156a1-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="156a1-137">String</span></span>|<span data-ttu-id="156a1-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="156a1-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156a1-139">Требования</span><span class="sxs-lookup"><span data-stu-id="156a1-139">Requirements</span></span>

|<span data-ttu-id="156a1-140">Требование</span><span class="sxs-lookup"><span data-stu-id="156a1-140">Requirement</span></span>| <span data-ttu-id="156a1-141">Значение</span><span class="sxs-lookup"><span data-stu-id="156a1-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="156a1-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="156a1-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156a1-143">1.0</span><span class="sxs-lookup"><span data-stu-id="156a1-143">1.0</span></span>|
|[<span data-ttu-id="156a1-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="156a1-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156a1-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="156a1-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="156a1-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="156a1-146">CoercionType :String</span></span>

<span data-ttu-id="156a1-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="156a1-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="156a1-148">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-148">Type</span></span>

*   <span data-ttu-id="156a1-149">String</span><span class="sxs-lookup"><span data-stu-id="156a1-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156a1-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="156a1-150">Properties:</span></span>

|<span data-ttu-id="156a1-151">Имя</span><span class="sxs-lookup"><span data-stu-id="156a1-151">Name</span></span>| <span data-ttu-id="156a1-152">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-152">Type</span></span>| <span data-ttu-id="156a1-153">Описание</span><span class="sxs-lookup"><span data-stu-id="156a1-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="156a1-154">String</span><span class="sxs-lookup"><span data-stu-id="156a1-154">String</span></span>|<span data-ttu-id="156a1-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="156a1-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="156a1-156">String</span><span class="sxs-lookup"><span data-stu-id="156a1-156">String</span></span>|<span data-ttu-id="156a1-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="156a1-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156a1-158">Требования</span><span class="sxs-lookup"><span data-stu-id="156a1-158">Requirements</span></span>

|<span data-ttu-id="156a1-159">Требование</span><span class="sxs-lookup"><span data-stu-id="156a1-159">Requirement</span></span>| <span data-ttu-id="156a1-160">Значение</span><span class="sxs-lookup"><span data-stu-id="156a1-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="156a1-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="156a1-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156a1-162">1.0</span><span class="sxs-lookup"><span data-stu-id="156a1-162">1.0</span></span>|
|[<span data-ttu-id="156a1-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="156a1-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156a1-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="156a1-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="156a1-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="156a1-165">EventType :String</span></span>

<span data-ttu-id="156a1-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="156a1-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="156a1-167">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-167">Type</span></span>

*   <span data-ttu-id="156a1-168">String</span><span class="sxs-lookup"><span data-stu-id="156a1-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156a1-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="156a1-169">Properties:</span></span>

| <span data-ttu-id="156a1-170">Имя</span><span class="sxs-lookup"><span data-stu-id="156a1-170">Name</span></span> | <span data-ttu-id="156a1-171">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-171">Type</span></span> | <span data-ttu-id="156a1-172">Описание</span><span class="sxs-lookup"><span data-stu-id="156a1-172">Description</span></span> | <span data-ttu-id="156a1-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="156a1-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="156a1-174">String</span><span class="sxs-lookup"><span data-stu-id="156a1-174">String</span></span> | <span data-ttu-id="156a1-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="156a1-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="156a1-176">1.7</span><span class="sxs-lookup"><span data-stu-id="156a1-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="156a1-177">String</span><span class="sxs-lookup"><span data-stu-id="156a1-177">String</span></span> | <span data-ttu-id="156a1-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="156a1-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="156a1-179">1.5</span><span class="sxs-lookup"><span data-stu-id="156a1-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="156a1-180">String</span><span class="sxs-lookup"><span data-stu-id="156a1-180">String</span></span> | <span data-ttu-id="156a1-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="156a1-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="156a1-182">1.7</span><span class="sxs-lookup"><span data-stu-id="156a1-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="156a1-183">String</span><span class="sxs-lookup"><span data-stu-id="156a1-183">String</span></span> | <span data-ttu-id="156a1-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="156a1-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="156a1-185">1.7</span><span class="sxs-lookup"><span data-stu-id="156a1-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="156a1-186">Требования</span><span class="sxs-lookup"><span data-stu-id="156a1-186">Requirements</span></span>

|<span data-ttu-id="156a1-187">Требование</span><span class="sxs-lookup"><span data-stu-id="156a1-187">Requirement</span></span>| <span data-ttu-id="156a1-188">Значение</span><span class="sxs-lookup"><span data-stu-id="156a1-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="156a1-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="156a1-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156a1-190">1.5</span><span class="sxs-lookup"><span data-stu-id="156a1-190">1.5</span></span> |
|[<span data-ttu-id="156a1-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="156a1-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156a1-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="156a1-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="156a1-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="156a1-193">SourceProperty :String</span></span>

<span data-ttu-id="156a1-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="156a1-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="156a1-195">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-195">Type</span></span>

*   <span data-ttu-id="156a1-196">String</span><span class="sxs-lookup"><span data-stu-id="156a1-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="156a1-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="156a1-197">Properties:</span></span>

|<span data-ttu-id="156a1-198">Имя</span><span class="sxs-lookup"><span data-stu-id="156a1-198">Name</span></span>| <span data-ttu-id="156a1-199">Тип</span><span class="sxs-lookup"><span data-stu-id="156a1-199">Type</span></span>| <span data-ttu-id="156a1-200">Описание</span><span class="sxs-lookup"><span data-stu-id="156a1-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="156a1-201">String</span><span class="sxs-lookup"><span data-stu-id="156a1-201">String</span></span>|<span data-ttu-id="156a1-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="156a1-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="156a1-203">String</span><span class="sxs-lookup"><span data-stu-id="156a1-203">String</span></span>|<span data-ttu-id="156a1-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="156a1-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="156a1-205">Требования</span><span class="sxs-lookup"><span data-stu-id="156a1-205">Requirements</span></span>

|<span data-ttu-id="156a1-206">Требование</span><span class="sxs-lookup"><span data-stu-id="156a1-206">Requirement</span></span>| <span data-ttu-id="156a1-207">Значение</span><span class="sxs-lookup"><span data-stu-id="156a1-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="156a1-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="156a1-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="156a1-209">1.0</span><span class="sxs-lookup"><span data-stu-id="156a1-209">1.0</span></span>|
|[<span data-ttu-id="156a1-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="156a1-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="156a1-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="156a1-211">Compose or Read</span></span>|
