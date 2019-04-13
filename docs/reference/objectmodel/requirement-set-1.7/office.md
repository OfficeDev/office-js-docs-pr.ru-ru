---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838454"
---
# <a name="office"></a><span data-ttu-id="f7a1e-102">Office</span><span class="sxs-lookup"><span data-stu-id="f7a1e-102">Office</span></span>

<span data-ttu-id="f7a1e-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f7a1e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7a1e-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f7a1e-105">Requirements</span></span>

|<span data-ttu-id="f7a1e-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f7a1e-106">Requirement</span></span>| <span data-ttu-id="f7a1e-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7a1e-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7a1e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7a1e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f7a1e-109">1.0</span></span>|
|[<span data-ttu-id="f7a1e-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7a1e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7a1e-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f7a1e-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="f7a1e-112">Members and methods</span></span>

| <span data-ttu-id="f7a1e-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="f7a1e-113">Member</span></span> | <span data-ttu-id="f7a1e-114">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f7a1e-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f7a1e-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f7a1e-116">Member</span><span class="sxs-lookup"><span data-stu-id="f7a1e-116">Member</span></span> |
| [<span data-ttu-id="f7a1e-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f7a1e-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f7a1e-118">Member</span><span class="sxs-lookup"><span data-stu-id="f7a1e-118">Member</span></span> |
| [<span data-ttu-id="f7a1e-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f7a1e-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f7a1e-120">Member</span><span class="sxs-lookup"><span data-stu-id="f7a1e-120">Member</span></span> |
| [<span data-ttu-id="f7a1e-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f7a1e-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f7a1e-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="f7a1e-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f7a1e-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f7a1e-123">Namespaces</span></span>

<span data-ttu-id="f7a1e-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f7a1e-125">[MailboxEnums.](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f7a1e-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="f7a1e-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f7a1e-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="f7a1e-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f7a1e-129">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-129">Type</span></span>

*   <span data-ttu-id="f7a1e-130">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7a1e-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7a1e-131">Properties:</span></span>

|<span data-ttu-id="f7a1e-132">Имя</span><span class="sxs-lookup"><span data-stu-id="f7a1e-132">Name</span></span>| <span data-ttu-id="f7a1e-133">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-133">Type</span></span>| <span data-ttu-id="f7a1e-134">Описание</span><span class="sxs-lookup"><span data-stu-id="f7a1e-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f7a1e-135">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-135">String</span></span>|<span data-ttu-id="f7a1e-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f7a1e-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="f7a1e-137">String</span></span>|<span data-ttu-id="f7a1e-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7a1e-139">Требования</span><span class="sxs-lookup"><span data-stu-id="f7a1e-139">Requirements</span></span>

|<span data-ttu-id="f7a1e-140">Требование</span><span class="sxs-lookup"><span data-stu-id="f7a1e-140">Requirement</span></span>| <span data-ttu-id="f7a1e-141">Значение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7a1e-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7a1e-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7a1e-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f7a1e-143">1.0</span></span>|
|[<span data-ttu-id="f7a1e-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7a1e-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7a1e-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="f7a1e-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-146">CoercionType :String</span></span>

<span data-ttu-id="f7a1e-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7a1e-148">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-148">Type</span></span>

*   <span data-ttu-id="f7a1e-149">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7a1e-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7a1e-150">Properties:</span></span>

|<span data-ttu-id="f7a1e-151">Имя</span><span class="sxs-lookup"><span data-stu-id="f7a1e-151">Name</span></span>| <span data-ttu-id="f7a1e-152">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-152">Type</span></span>| <span data-ttu-id="f7a1e-153">Описание</span><span class="sxs-lookup"><span data-stu-id="f7a1e-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f7a1e-154">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-154">String</span></span>|<span data-ttu-id="f7a1e-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f7a1e-156">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-156">String</span></span>|<span data-ttu-id="f7a1e-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7a1e-158">Требования</span><span class="sxs-lookup"><span data-stu-id="f7a1e-158">Requirements</span></span>

|<span data-ttu-id="f7a1e-159">Требование</span><span class="sxs-lookup"><span data-stu-id="f7a1e-159">Requirement</span></span>| <span data-ttu-id="f7a1e-160">Значение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7a1e-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7a1e-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7a1e-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f7a1e-162">1.0</span></span>|
|[<span data-ttu-id="f7a1e-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7a1e-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7a1e-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="f7a1e-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-165">EventType :String</span></span>

<span data-ttu-id="f7a1e-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f7a1e-167">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-167">Type</span></span>

*   <span data-ttu-id="f7a1e-168">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7a1e-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7a1e-169">Properties:</span></span>

| <span data-ttu-id="f7a1e-170">Имя</span><span class="sxs-lookup"><span data-stu-id="f7a1e-170">Name</span></span> | <span data-ttu-id="f7a1e-171">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-171">Type</span></span> | <span data-ttu-id="f7a1e-172">Описание</span><span class="sxs-lookup"><span data-stu-id="f7a1e-172">Description</span></span> | <span data-ttu-id="f7a1e-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="f7a1e-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="f7a1e-174">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-174">String</span></span> | <span data-ttu-id="f7a1e-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f7a1e-176">1.7</span><span class="sxs-lookup"><span data-stu-id="f7a1e-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="f7a1e-177">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-177">String</span></span> | <span data-ttu-id="f7a1e-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f7a1e-179">1.5</span><span class="sxs-lookup"><span data-stu-id="f7a1e-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f7a1e-180">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-180">String</span></span> | <span data-ttu-id="f7a1e-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f7a1e-182">1.7</span><span class="sxs-lookup"><span data-stu-id="f7a1e-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f7a1e-183">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-183">String</span></span> | <span data-ttu-id="f7a1e-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f7a1e-185">1.7</span><span class="sxs-lookup"><span data-stu-id="f7a1e-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f7a1e-186">Требования</span><span class="sxs-lookup"><span data-stu-id="f7a1e-186">Requirements</span></span>

|<span data-ttu-id="f7a1e-187">Требование</span><span class="sxs-lookup"><span data-stu-id="f7a1e-187">Requirement</span></span>| <span data-ttu-id="f7a1e-188">Значение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7a1e-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f7a1e-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7a1e-190">1.5</span><span class="sxs-lookup"><span data-stu-id="f7a1e-190">1.5</span></span> |
|[<span data-ttu-id="f7a1e-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7a1e-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7a1e-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="f7a1e-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-193">SourceProperty :String</span></span>

<span data-ttu-id="f7a1e-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7a1e-195">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-195">Type</span></span>

*   <span data-ttu-id="f7a1e-196">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7a1e-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f7a1e-197">Properties:</span></span>

|<span data-ttu-id="f7a1e-198">Имя</span><span class="sxs-lookup"><span data-stu-id="f7a1e-198">Name</span></span>| <span data-ttu-id="f7a1e-199">Тип</span><span class="sxs-lookup"><span data-stu-id="f7a1e-199">Type</span></span>| <span data-ttu-id="f7a1e-200">Описание</span><span class="sxs-lookup"><span data-stu-id="f7a1e-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f7a1e-201">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-201">String</span></span>|<span data-ttu-id="f7a1e-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f7a1e-203">String</span><span class="sxs-lookup"><span data-stu-id="f7a1e-203">String</span></span>|<span data-ttu-id="f7a1e-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f7a1e-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7a1e-205">Требования</span><span class="sxs-lookup"><span data-stu-id="f7a1e-205">Requirements</span></span>

|<span data-ttu-id="f7a1e-206">Требование</span><span class="sxs-lookup"><span data-stu-id="f7a1e-206">Requirement</span></span>| <span data-ttu-id="f7a1e-207">Значение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7a1e-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f7a1e-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f7a1e-209">1.0</span><span class="sxs-lookup"><span data-stu-id="f7a1e-209">1.0</span></span>|
|[<span data-ttu-id="f7a1e-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f7a1e-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f7a1e-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f7a1e-211">Compose or Read</span></span>|
