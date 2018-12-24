---
title: Пространство имен Office — набор обязательных элементов 1.7
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 2bf1c31f4dc4156cb4f1d0eb3508193305c860e9
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432804"
---
# <a name="office"></a><span data-ttu-id="62472-102">Office</span><span class="sxs-lookup"><span data-stu-id="62472-102">Office</span></span>

<span data-ttu-id="62472-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="62472-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="62472-105">Требования</span><span class="sxs-lookup"><span data-stu-id="62472-105">Requirements</span></span>

|<span data-ttu-id="62472-106">Requirement</span><span class="sxs-lookup"><span data-stu-id="62472-106">Requirement</span></span>| <span data-ttu-id="62472-107">Значение</span><span class="sxs-lookup"><span data-stu-id="62472-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="62472-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62472-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62472-109">1.0</span><span class="sxs-lookup"><span data-stu-id="62472-109">1.0</span></span>|
|[<span data-ttu-id="62472-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62472-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62472-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62472-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="62472-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="62472-112">Members and methods</span></span>

| <span data-ttu-id="62472-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="62472-113">Member</span></span> | <span data-ttu-id="62472-114">Тип</span><span class="sxs-lookup"><span data-stu-id="62472-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="62472-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="62472-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="62472-116">Член</span><span class="sxs-lookup"><span data-stu-id="62472-116">Member</span></span> |
| [<span data-ttu-id="62472-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="62472-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="62472-118">Член</span><span class="sxs-lookup"><span data-stu-id="62472-118">Member</span></span> |
| [<span data-ttu-id="62472-119">EventType</span><span class="sxs-lookup"><span data-stu-id="62472-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="62472-120">Член</span><span class="sxs-lookup"><span data-stu-id="62472-120">Member</span></span> |
| [<span data-ttu-id="62472-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="62472-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="62472-122">Член</span><span class="sxs-lookup"><span data-stu-id="62472-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="62472-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="62472-123">Namespaces</span></span>

<span data-ttu-id="62472-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="62472-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="62472-125">[MailboxEnums.](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="62472-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="62472-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="62472-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="62472-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="62472-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="62472-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="62472-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="62472-129">Тип:</span><span class="sxs-lookup"><span data-stu-id="62472-129">Type:</span></span>

*   <span data-ttu-id="62472-130">String</span><span class="sxs-lookup"><span data-stu-id="62472-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="62472-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="62472-131">Properties:</span></span>

|<span data-ttu-id="62472-132">Имя</span><span class="sxs-lookup"><span data-stu-id="62472-132">Name</span></span>| <span data-ttu-id="62472-133">Тип</span><span class="sxs-lookup"><span data-stu-id="62472-133">Type</span></span>| <span data-ttu-id="62472-134">Описание</span><span class="sxs-lookup"><span data-stu-id="62472-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="62472-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="62472-135">String</span></span>|<span data-ttu-id="62472-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="62472-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="62472-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="62472-137">String</span></span>|<span data-ttu-id="62472-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="62472-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62472-139">Требования</span><span class="sxs-lookup"><span data-stu-id="62472-139">Requirements</span></span>

|<span data-ttu-id="62472-140">Requirement</span><span class="sxs-lookup"><span data-stu-id="62472-140">Requirement</span></span>| <span data-ttu-id="62472-141">Значение</span><span class="sxs-lookup"><span data-stu-id="62472-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="62472-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62472-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62472-143">1.0</span><span class="sxs-lookup"><span data-stu-id="62472-143">1.0</span></span>|
|[<span data-ttu-id="62472-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62472-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62472-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62472-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="62472-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="62472-146">CoercionType :String</span></span>

<span data-ttu-id="62472-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="62472-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="62472-148">Тип:</span><span class="sxs-lookup"><span data-stu-id="62472-148">Type:</span></span>

*   <span data-ttu-id="62472-149">String</span><span class="sxs-lookup"><span data-stu-id="62472-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="62472-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="62472-150">Properties:</span></span>

|<span data-ttu-id="62472-151">Имя</span><span class="sxs-lookup"><span data-stu-id="62472-151">Name</span></span>| <span data-ttu-id="62472-152">Тип</span><span class="sxs-lookup"><span data-stu-id="62472-152">Type</span></span>| <span data-ttu-id="62472-153">Описание</span><span class="sxs-lookup"><span data-stu-id="62472-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="62472-154">String</span><span class="sxs-lookup"><span data-stu-id="62472-154">String</span></span>|<span data-ttu-id="62472-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="62472-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="62472-156">String</span><span class="sxs-lookup"><span data-stu-id="62472-156">String</span></span>|<span data-ttu-id="62472-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="62472-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62472-158">Требования</span><span class="sxs-lookup"><span data-stu-id="62472-158">Requirements</span></span>

|<span data-ttu-id="62472-159">Requirement</span><span class="sxs-lookup"><span data-stu-id="62472-159">Requirement</span></span>| <span data-ttu-id="62472-160">Значение</span><span class="sxs-lookup"><span data-stu-id="62472-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="62472-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62472-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62472-162">1.0</span><span class="sxs-lookup"><span data-stu-id="62472-162">1.0</span></span>|
|[<span data-ttu-id="62472-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62472-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62472-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62472-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="62472-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="62472-165">EventType :String</span></span>

<span data-ttu-id="62472-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="62472-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="62472-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="62472-167">Type:</span></span>

*   <span data-ttu-id="62472-168">String</span><span class="sxs-lookup"><span data-stu-id="62472-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="62472-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="62472-169">Properties:</span></span>

| <span data-ttu-id="62472-170">Имя</span><span class="sxs-lookup"><span data-stu-id="62472-170">Name</span></span> | <span data-ttu-id="62472-171">Тип</span><span class="sxs-lookup"><span data-stu-id="62472-171">Type</span></span> | <span data-ttu-id="62472-172">Описание</span><span class="sxs-lookup"><span data-stu-id="62472-172">Description</span></span> | <span data-ttu-id="62472-173">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="62472-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="62472-174">String</span><span class="sxs-lookup"><span data-stu-id="62472-174">String</span></span> | <span data-ttu-id="62472-175">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="62472-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="62472-176">1.7</span><span class="sxs-lookup"><span data-stu-id="62472-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="62472-177">String</span><span class="sxs-lookup"><span data-stu-id="62472-177">String</span></span> | <span data-ttu-id="62472-178">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="62472-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="62472-179">1.5</span><span class="sxs-lookup"><span data-stu-id="62472-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="62472-180">String</span><span class="sxs-lookup"><span data-stu-id="62472-180">String</span></span> | <span data-ttu-id="62472-181">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="62472-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="62472-182">1.7</span><span class="sxs-lookup"><span data-stu-id="62472-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="62472-183">String</span><span class="sxs-lookup"><span data-stu-id="62472-183">String</span></span> | <span data-ttu-id="62472-184">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="62472-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="62472-185">1.7</span><span class="sxs-lookup"><span data-stu-id="62472-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="62472-186">Требования</span><span class="sxs-lookup"><span data-stu-id="62472-186">Requirements</span></span>

|<span data-ttu-id="62472-187">Requirement</span><span class="sxs-lookup"><span data-stu-id="62472-187">Requirement</span></span>| <span data-ttu-id="62472-188">Значение</span><span class="sxs-lookup"><span data-stu-id="62472-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="62472-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62472-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62472-190">1.5</span><span class="sxs-lookup"><span data-stu-id="62472-190">1.5</span></span> |
|[<span data-ttu-id="62472-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62472-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62472-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62472-192">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="62472-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="62472-193">SourceProperty :String</span></span>

<span data-ttu-id="62472-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="62472-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="62472-195">Тип:</span><span class="sxs-lookup"><span data-stu-id="62472-195">Type:</span></span>

*   <span data-ttu-id="62472-196">String</span><span class="sxs-lookup"><span data-stu-id="62472-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="62472-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="62472-197">Properties:</span></span>

|<span data-ttu-id="62472-198">Имя</span><span class="sxs-lookup"><span data-stu-id="62472-198">Name</span></span>| <span data-ttu-id="62472-199">Тип</span><span class="sxs-lookup"><span data-stu-id="62472-199">Type</span></span>| <span data-ttu-id="62472-200">Описание</span><span class="sxs-lookup"><span data-stu-id="62472-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="62472-201">String</span><span class="sxs-lookup"><span data-stu-id="62472-201">String</span></span>|<span data-ttu-id="62472-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="62472-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="62472-203">String</span><span class="sxs-lookup"><span data-stu-id="62472-203">String</span></span>|<span data-ttu-id="62472-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="62472-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62472-205">Требования</span><span class="sxs-lookup"><span data-stu-id="62472-205">Requirements</span></span>

|<span data-ttu-id="62472-206">Requirement</span><span class="sxs-lookup"><span data-stu-id="62472-206">Requirement</span></span>| <span data-ttu-id="62472-207">Значение</span><span class="sxs-lookup"><span data-stu-id="62472-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="62472-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62472-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62472-209">1.0</span><span class="sxs-lookup"><span data-stu-id="62472-209">1.0</span></span>|
|[<span data-ttu-id="62472-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62472-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62472-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62472-211">Compose or read</span></span>|