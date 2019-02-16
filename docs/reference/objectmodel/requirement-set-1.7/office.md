---
title: Пространство имен Office — набор обязательных элементов 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: d6422e470864d5a02db37e1fef295e8cbb82a213
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067897"
---
# <a name="office"></a><span data-ttu-id="3339a-102">Office</span><span class="sxs-lookup"><span data-stu-id="3339a-102">Office</span></span>

<span data-ttu-id="3339a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3339a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3339a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="3339a-105">Requirements</span></span>

|<span data-ttu-id="3339a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="3339a-106">Requirement</span></span>| <span data-ttu-id="3339a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="3339a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3339a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3339a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3339a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3339a-109">1.0</span></span>|
|[<span data-ttu-id="3339a-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3339a-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3339a-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3339a-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3339a-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="3339a-112">Members and methods</span></span>

| <span data-ttu-id="3339a-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="3339a-113">Member</span></span> | <span data-ttu-id="3339a-114">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3339a-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3339a-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3339a-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="3339a-116">Member</span></span> |
| [<span data-ttu-id="3339a-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3339a-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3339a-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="3339a-118">Member</span></span> |
| [<span data-ttu-id="3339a-119">EventType</span><span class="sxs-lookup"><span data-stu-id="3339a-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3339a-120">Член</span><span class="sxs-lookup"><span data-stu-id="3339a-120">Member</span></span> |
| [<span data-ttu-id="3339a-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3339a-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3339a-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="3339a-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3339a-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3339a-123">Namespaces</span></span>

<span data-ttu-id="3339a-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="3339a-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="3339a-125">[MailboxEnums.](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="3339a-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="3339a-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="3339a-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="3339a-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="3339a-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="3339a-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="3339a-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3339a-129">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-129">Type</span></span>

*   <span data-ttu-id="3339a-130">String</span><span class="sxs-lookup"><span data-stu-id="3339a-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3339a-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3339a-131">Properties:</span></span>

|<span data-ttu-id="3339a-132">Имя</span><span class="sxs-lookup"><span data-stu-id="3339a-132">Name</span></span>| <span data-ttu-id="3339a-133">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-133">Type</span></span>| <span data-ttu-id="3339a-134">Описание</span><span class="sxs-lookup"><span data-stu-id="3339a-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3339a-135">Строка</span><span class="sxs-lookup"><span data-stu-id="3339a-135">String</span></span>|<span data-ttu-id="3339a-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="3339a-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3339a-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="3339a-137">String</span></span>|<span data-ttu-id="3339a-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="3339a-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3339a-139">Требования</span><span class="sxs-lookup"><span data-stu-id="3339a-139">Requirements</span></span>

|<span data-ttu-id="3339a-140">Требование</span><span class="sxs-lookup"><span data-stu-id="3339a-140">Requirement</span></span>| <span data-ttu-id="3339a-141">Значение</span><span class="sxs-lookup"><span data-stu-id="3339a-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="3339a-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3339a-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3339a-143">1.0</span><span class="sxs-lookup"><span data-stu-id="3339a-143">1.0</span></span>|
|[<span data-ttu-id="3339a-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3339a-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3339a-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3339a-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="3339a-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="3339a-146">CoercionType :String</span></span>

<span data-ttu-id="3339a-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3339a-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3339a-148">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-148">Type</span></span>

*   <span data-ttu-id="3339a-149">String</span><span class="sxs-lookup"><span data-stu-id="3339a-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3339a-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3339a-150">Properties:</span></span>

|<span data-ttu-id="3339a-151">Имя</span><span class="sxs-lookup"><span data-stu-id="3339a-151">Name</span></span>| <span data-ttu-id="3339a-152">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-152">Type</span></span>| <span data-ttu-id="3339a-153">Описание</span><span class="sxs-lookup"><span data-stu-id="3339a-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3339a-154">Строка</span><span class="sxs-lookup"><span data-stu-id="3339a-154">String</span></span>|<span data-ttu-id="3339a-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="3339a-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3339a-156">String</span><span class="sxs-lookup"><span data-stu-id="3339a-156">String</span></span>|<span data-ttu-id="3339a-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="3339a-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3339a-158">Требования</span><span class="sxs-lookup"><span data-stu-id="3339a-158">Requirements</span></span>

|<span data-ttu-id="3339a-159">Требование</span><span class="sxs-lookup"><span data-stu-id="3339a-159">Requirement</span></span>| <span data-ttu-id="3339a-160">Значение</span><span class="sxs-lookup"><span data-stu-id="3339a-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="3339a-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3339a-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3339a-162">1.0</span><span class="sxs-lookup"><span data-stu-id="3339a-162">1.0</span></span>|
|[<span data-ttu-id="3339a-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3339a-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3339a-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3339a-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="3339a-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="3339a-165">EventType :String</span></span>

<span data-ttu-id="3339a-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="3339a-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3339a-167">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-167">Type</span></span>

*   <span data-ttu-id="3339a-168">String</span><span class="sxs-lookup"><span data-stu-id="3339a-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3339a-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3339a-169">Properties:</span></span>

| <span data-ttu-id="3339a-170">Имя</span><span class="sxs-lookup"><span data-stu-id="3339a-170">Name</span></span> | <span data-ttu-id="3339a-171">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-171">Type</span></span> | <span data-ttu-id="3339a-172">Описание</span><span class="sxs-lookup"><span data-stu-id="3339a-172">Description</span></span> | <span data-ttu-id="3339a-173">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="3339a-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="3339a-174">String</span><span class="sxs-lookup"><span data-stu-id="3339a-174">String</span></span> | <span data-ttu-id="3339a-175">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="3339a-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="3339a-176">1.7</span><span class="sxs-lookup"><span data-stu-id="3339a-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="3339a-177">String</span><span class="sxs-lookup"><span data-stu-id="3339a-177">String</span></span> | <span data-ttu-id="3339a-178">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="3339a-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3339a-179">1.5</span><span class="sxs-lookup"><span data-stu-id="3339a-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="3339a-180">String</span><span class="sxs-lookup"><span data-stu-id="3339a-180">String</span></span> | <span data-ttu-id="3339a-181">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="3339a-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="3339a-182">1.7</span><span class="sxs-lookup"><span data-stu-id="3339a-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="3339a-183">String</span><span class="sxs-lookup"><span data-stu-id="3339a-183">String</span></span> | <span data-ttu-id="3339a-184">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="3339a-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="3339a-185">1.7</span><span class="sxs-lookup"><span data-stu-id="3339a-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3339a-186">Требования</span><span class="sxs-lookup"><span data-stu-id="3339a-186">Requirements</span></span>

|<span data-ttu-id="3339a-187">Требование</span><span class="sxs-lookup"><span data-stu-id="3339a-187">Requirement</span></span>| <span data-ttu-id="3339a-188">Значение</span><span class="sxs-lookup"><span data-stu-id="3339a-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="3339a-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3339a-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3339a-190">1.5</span><span class="sxs-lookup"><span data-stu-id="3339a-190">1.5</span></span> |
|[<span data-ttu-id="3339a-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3339a-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3339a-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3339a-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="3339a-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="3339a-193">SourceProperty :String</span></span>

<span data-ttu-id="3339a-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="3339a-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3339a-195">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-195">Type</span></span>

*   <span data-ttu-id="3339a-196">String</span><span class="sxs-lookup"><span data-stu-id="3339a-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3339a-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3339a-197">Properties:</span></span>

|<span data-ttu-id="3339a-198">Имя</span><span class="sxs-lookup"><span data-stu-id="3339a-198">Name</span></span>| <span data-ttu-id="3339a-199">Тип</span><span class="sxs-lookup"><span data-stu-id="3339a-199">Type</span></span>| <span data-ttu-id="3339a-200">Описание</span><span class="sxs-lookup"><span data-stu-id="3339a-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3339a-201">Строка</span><span class="sxs-lookup"><span data-stu-id="3339a-201">String</span></span>|<span data-ttu-id="3339a-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="3339a-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3339a-203">String</span><span class="sxs-lookup"><span data-stu-id="3339a-203">String</span></span>|<span data-ttu-id="3339a-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="3339a-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3339a-205">Требования</span><span class="sxs-lookup"><span data-stu-id="3339a-205">Requirements</span></span>

|<span data-ttu-id="3339a-206">Требование</span><span class="sxs-lookup"><span data-stu-id="3339a-206">Requirement</span></span>| <span data-ttu-id="3339a-207">Значение</span><span class="sxs-lookup"><span data-stu-id="3339a-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="3339a-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3339a-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3339a-209">1.0</span><span class="sxs-lookup"><span data-stu-id="3339a-209">1.0</span></span>|
|[<span data-ttu-id="3339a-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3339a-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3339a-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3339a-211">Compose or Read</span></span>|
