---
title: Пространство имен Office — предварительная версия набора обязательных элементов
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: f4a4f0d7a4ce0de433d4e70b6a4675b5f63f26f0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457931"
---
# <a name="office"></a><span data-ttu-id="e7c84-102">Office</span><span class="sxs-lookup"><span data-stu-id="e7c84-102">Office</span></span>

<span data-ttu-id="e7c84-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e7c84-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e7c84-105">Требования</span><span class="sxs-lookup"><span data-stu-id="e7c84-105">Requirements</span></span>

|<span data-ttu-id="e7c84-106">Требование</span><span class="sxs-lookup"><span data-stu-id="e7c84-106">Requirement</span></span>| <span data-ttu-id="e7c84-107">Значение</span><span class="sxs-lookup"><span data-stu-id="e7c84-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7c84-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7c84-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7c84-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e7c84-109">1.0</span></span>|
|[<span data-ttu-id="e7c84-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7c84-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e7c84-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7c84-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e7c84-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="e7c84-112">Members and methods</span></span>

| <span data-ttu-id="e7c84-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="e7c84-113">Member</span></span> | <span data-ttu-id="e7c84-114">Тип</span><span class="sxs-lookup"><span data-stu-id="e7c84-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e7c84-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e7c84-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e7c84-116">Член</span><span class="sxs-lookup"><span data-stu-id="e7c84-116">Member</span></span> |
| [<span data-ttu-id="e7c84-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e7c84-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e7c84-118">Член</span><span class="sxs-lookup"><span data-stu-id="e7c84-118">Member</span></span> |
| [<span data-ttu-id="e7c84-119">EventType</span><span class="sxs-lookup"><span data-stu-id="e7c84-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e7c84-120">Член</span><span class="sxs-lookup"><span data-stu-id="e7c84-120">Member</span></span> |
| [<span data-ttu-id="e7c84-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e7c84-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e7c84-122">Член</span><span class="sxs-lookup"><span data-stu-id="e7c84-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e7c84-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="e7c84-123">Namespaces</span></span>

<span data-ttu-id="e7c84-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7c84-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e7c84-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="e7c84-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="e7c84-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="e7c84-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="e7c84-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="e7c84-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="e7c84-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="e7c84-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e7c84-129">Тип:</span><span class="sxs-lookup"><span data-stu-id="e7c84-129">Type:</span></span>

*   <span data-ttu-id="e7c84-130">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7c84-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7c84-131">Properties:</span></span>

|<span data-ttu-id="e7c84-132">Имя</span><span class="sxs-lookup"><span data-stu-id="e7c84-132">Name</span></span>| <span data-ttu-id="e7c84-133">Тип</span><span class="sxs-lookup"><span data-stu-id="e7c84-133">Type</span></span>| <span data-ttu-id="e7c84-134">Описание</span><span class="sxs-lookup"><span data-stu-id="e7c84-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e7c84-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="e7c84-135">String</span></span>|<span data-ttu-id="e7c84-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="e7c84-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e7c84-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="e7c84-137">String</span></span>|<span data-ttu-id="e7c84-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="e7c84-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7c84-139">Требования</span><span class="sxs-lookup"><span data-stu-id="e7c84-139">Requirements</span></span>

|<span data-ttu-id="e7c84-140">Требование</span><span class="sxs-lookup"><span data-stu-id="e7c84-140">Requirement</span></span>| <span data-ttu-id="e7c84-141">Значение</span><span class="sxs-lookup"><span data-stu-id="e7c84-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7c84-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7c84-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7c84-143">1.0</span><span class="sxs-lookup"><span data-stu-id="e7c84-143">1.0</span></span>|
|[<span data-ttu-id="e7c84-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7c84-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e7c84-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7c84-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="e7c84-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="e7c84-146">CoercionType :String</span></span>

<span data-ttu-id="e7c84-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e7c84-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7c84-148">Тип:</span><span class="sxs-lookup"><span data-stu-id="e7c84-148">Type:</span></span>

*   <span data-ttu-id="e7c84-149">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7c84-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7c84-150">Properties:</span></span>

|<span data-ttu-id="e7c84-151">Имя</span><span class="sxs-lookup"><span data-stu-id="e7c84-151">Name</span></span>| <span data-ttu-id="e7c84-152">Тип</span><span class="sxs-lookup"><span data-stu-id="e7c84-152">Type</span></span>| <span data-ttu-id="e7c84-153">Описание</span><span class="sxs-lookup"><span data-stu-id="e7c84-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e7c84-154">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-154">String</span></span>|<span data-ttu-id="e7c84-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="e7c84-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e7c84-156">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-156">String</span></span>|<span data-ttu-id="e7c84-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="e7c84-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7c84-158">Требования</span><span class="sxs-lookup"><span data-stu-id="e7c84-158">Requirements</span></span>

|<span data-ttu-id="e7c84-159">Требование</span><span class="sxs-lookup"><span data-stu-id="e7c84-159">Requirement</span></span>| <span data-ttu-id="e7c84-160">Значение</span><span class="sxs-lookup"><span data-stu-id="e7c84-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7c84-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7c84-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7c84-162">1.0</span><span class="sxs-lookup"><span data-stu-id="e7c84-162">1.0</span></span>|
|[<span data-ttu-id="e7c84-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7c84-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e7c84-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7c84-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="e7c84-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="e7c84-165">EventType :String</span></span>

<span data-ttu-id="e7c84-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="e7c84-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e7c84-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="e7c84-167">Type:</span></span>

*   <span data-ttu-id="e7c84-168">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7c84-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7c84-169">Properties:</span></span>

| <span data-ttu-id="e7c84-170">Имя</span><span class="sxs-lookup"><span data-stu-id="e7c84-170">Name</span></span> | <span data-ttu-id="e7c84-171">Тип</span><span class="sxs-lookup"><span data-stu-id="e7c84-171">Type</span></span> | <span data-ttu-id="e7c84-172">Описание</span><span class="sxs-lookup"><span data-stu-id="e7c84-172">Description</span></span> | <span data-ttu-id="e7c84-173">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="e7c84-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="e7c84-174">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-174">String</span></span> | <span data-ttu-id="e7c84-175">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="e7c84-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e7c84-176">1.7</span><span class="sxs-lookup"><span data-stu-id="e7c84-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e7c84-177">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-177">String</span></span> | <span data-ttu-id="e7c84-178">Было добавлено или удалено вложение для элемента.</span><span class="sxs-lookup"><span data-stu-id="e7c84-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e7c84-179">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e7c84-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="e7c84-180">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-180">String</span></span> | <span data-ttu-id="e7c84-181">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="e7c84-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e7c84-182">1.5</span><span class="sxs-lookup"><span data-stu-id="e7c84-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="e7c84-183">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-183">String</span></span> | <span data-ttu-id="e7c84-184">Тема Office в почтовом ящике была изменена.</span><span class="sxs-lookup"><span data-stu-id="e7c84-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="e7c84-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="e7c84-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e7c84-186">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-186">String</span></span> | <span data-ttu-id="e7c84-187">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="e7c84-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e7c84-188">1.7</span><span class="sxs-lookup"><span data-stu-id="e7c84-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e7c84-189">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-189">String</span></span> | <span data-ttu-id="e7c84-190">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="e7c84-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e7c84-191">1.7</span><span class="sxs-lookup"><span data-stu-id="e7c84-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e7c84-192">Требования</span><span class="sxs-lookup"><span data-stu-id="e7c84-192">Requirements</span></span>

|<span data-ttu-id="e7c84-193">Требование</span><span class="sxs-lookup"><span data-stu-id="e7c84-193">Requirement</span></span>| <span data-ttu-id="e7c84-194">Значение</span><span class="sxs-lookup"><span data-stu-id="e7c84-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7c84-195">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e7c84-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7c84-196">1.5</span><span class="sxs-lookup"><span data-stu-id="e7c84-196">1.5</span></span> |
|[<span data-ttu-id="e7c84-197">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7c84-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e7c84-198">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7c84-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="e7c84-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="e7c84-199">SourceProperty :String</span></span>

<span data-ttu-id="e7c84-200">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="e7c84-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e7c84-201">Тип:</span><span class="sxs-lookup"><span data-stu-id="e7c84-201">Type:</span></span>

*   <span data-ttu-id="e7c84-202">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e7c84-203">Свойства:</span><span class="sxs-lookup"><span data-stu-id="e7c84-203">Properties:</span></span>

|<span data-ttu-id="e7c84-204">Имя</span><span class="sxs-lookup"><span data-stu-id="e7c84-204">Name</span></span>| <span data-ttu-id="e7c84-205">Тип</span><span class="sxs-lookup"><span data-stu-id="e7c84-205">Type</span></span>| <span data-ttu-id="e7c84-206">Описание</span><span class="sxs-lookup"><span data-stu-id="e7c84-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e7c84-207">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-207">String</span></span>|<span data-ttu-id="e7c84-208">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7c84-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e7c84-209">String</span><span class="sxs-lookup"><span data-stu-id="e7c84-209">String</span></span>|<span data-ttu-id="e7c84-210">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="e7c84-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e7c84-211">Требования</span><span class="sxs-lookup"><span data-stu-id="e7c84-211">Requirements</span></span>

|<span data-ttu-id="e7c84-212">Требование</span><span class="sxs-lookup"><span data-stu-id="e7c84-212">Requirement</span></span>| <span data-ttu-id="e7c84-213">Значение</span><span class="sxs-lookup"><span data-stu-id="e7c84-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="e7c84-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e7c84-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e7c84-215">1.0</span><span class="sxs-lookup"><span data-stu-id="e7c84-215">1.0</span></span>|
|[<span data-ttu-id="e7c84-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e7c84-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e7c84-217">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e7c84-217">Compose or read</span></span>|