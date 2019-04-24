---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451621"
---
# <a name="office"></a><span data-ttu-id="c149c-102">Office</span><span class="sxs-lookup"><span data-stu-id="c149c-102">Office</span></span>

<span data-ttu-id="c149c-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c149c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c149c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="c149c-105">Requirements</span></span>

|<span data-ttu-id="c149c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="c149c-106">Requirement</span></span>| <span data-ttu-id="c149c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="c149c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c149c-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c149c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c149c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c149c-109">1.0</span></span>|
|[<span data-ttu-id="c149c-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c149c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c149c-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c149c-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c149c-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="c149c-112">Members and methods</span></span>

| <span data-ttu-id="c149c-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="c149c-113">Member</span></span> | <span data-ttu-id="c149c-114">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c149c-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c149c-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c149c-116">Member</span><span class="sxs-lookup"><span data-stu-id="c149c-116">Member</span></span> |
| [<span data-ttu-id="c149c-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c149c-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c149c-118">Member</span><span class="sxs-lookup"><span data-stu-id="c149c-118">Member</span></span> |
| [<span data-ttu-id="c149c-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c149c-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c149c-120">Member</span><span class="sxs-lookup"><span data-stu-id="c149c-120">Member</span></span> |
| [<span data-ttu-id="c149c-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c149c-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c149c-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="c149c-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c149c-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c149c-123">Namespaces</span></span>

<span data-ttu-id="c149c-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="c149c-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c149c-125">[MailboxEnums.](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c149c-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c149c-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="c149c-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c149c-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c149c-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="c149c-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c149c-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c149c-129">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-129">Type</span></span>

*   <span data-ttu-id="c149c-130">String</span><span class="sxs-lookup"><span data-stu-id="c149c-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c149c-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c149c-131">Properties:</span></span>

|<span data-ttu-id="c149c-132">Имя</span><span class="sxs-lookup"><span data-stu-id="c149c-132">Name</span></span>| <span data-ttu-id="c149c-133">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-133">Type</span></span>| <span data-ttu-id="c149c-134">Описание</span><span class="sxs-lookup"><span data-stu-id="c149c-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c149c-135">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-135">String</span></span>|<span data-ttu-id="c149c-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c149c-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c149c-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="c149c-137">String</span></span>|<span data-ttu-id="c149c-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="c149c-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c149c-139">Требования</span><span class="sxs-lookup"><span data-stu-id="c149c-139">Requirements</span></span>

|<span data-ttu-id="c149c-140">Требование</span><span class="sxs-lookup"><span data-stu-id="c149c-140">Requirement</span></span>| <span data-ttu-id="c149c-141">Значение</span><span class="sxs-lookup"><span data-stu-id="c149c-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c149c-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c149c-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c149c-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c149c-143">1.0</span></span>|
|[<span data-ttu-id="c149c-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c149c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c149c-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c149c-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="c149c-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c149c-146">CoercionType :String</span></span>

<span data-ttu-id="c149c-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c149c-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c149c-148">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-148">Type</span></span>

*   <span data-ttu-id="c149c-149">String</span><span class="sxs-lookup"><span data-stu-id="c149c-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c149c-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c149c-150">Properties:</span></span>

|<span data-ttu-id="c149c-151">Имя</span><span class="sxs-lookup"><span data-stu-id="c149c-151">Name</span></span>| <span data-ttu-id="c149c-152">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-152">Type</span></span>| <span data-ttu-id="c149c-153">Описание</span><span class="sxs-lookup"><span data-stu-id="c149c-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c149c-154">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-154">String</span></span>|<span data-ttu-id="c149c-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c149c-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c149c-156">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-156">String</span></span>|<span data-ttu-id="c149c-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c149c-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c149c-158">Требования</span><span class="sxs-lookup"><span data-stu-id="c149c-158">Requirements</span></span>

|<span data-ttu-id="c149c-159">Требование</span><span class="sxs-lookup"><span data-stu-id="c149c-159">Requirement</span></span>| <span data-ttu-id="c149c-160">Значение</span><span class="sxs-lookup"><span data-stu-id="c149c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c149c-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c149c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c149c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c149c-162">1.0</span></span>|
|[<span data-ttu-id="c149c-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c149c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c149c-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c149c-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="c149c-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c149c-165">EventType :String</span></span>

<span data-ttu-id="c149c-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="c149c-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c149c-167">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-167">Type</span></span>

*   <span data-ttu-id="c149c-168">String</span><span class="sxs-lookup"><span data-stu-id="c149c-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c149c-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c149c-169">Properties:</span></span>

| <span data-ttu-id="c149c-170">Имя</span><span class="sxs-lookup"><span data-stu-id="c149c-170">Name</span></span> | <span data-ttu-id="c149c-171">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-171">Type</span></span> | <span data-ttu-id="c149c-172">Описание</span><span class="sxs-lookup"><span data-stu-id="c149c-172">Description</span></span> | <span data-ttu-id="c149c-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="c149c-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c149c-174">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-174">String</span></span> | <span data-ttu-id="c149c-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="c149c-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c149c-176">1.7</span><span class="sxs-lookup"><span data-stu-id="c149c-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="c149c-177">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-177">String</span></span> | <span data-ttu-id="c149c-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="c149c-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="c149c-179">1.5</span><span class="sxs-lookup"><span data-stu-id="c149c-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c149c-180">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-180">String</span></span> | <span data-ttu-id="c149c-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="c149c-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c149c-182">1.7</span><span class="sxs-lookup"><span data-stu-id="c149c-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c149c-183">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-183">String</span></span> | <span data-ttu-id="c149c-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="c149c-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c149c-185">1.7</span><span class="sxs-lookup"><span data-stu-id="c149c-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c149c-186">Требования</span><span class="sxs-lookup"><span data-stu-id="c149c-186">Requirements</span></span>

|<span data-ttu-id="c149c-187">Требование</span><span class="sxs-lookup"><span data-stu-id="c149c-187">Requirement</span></span>| <span data-ttu-id="c149c-188">Значение</span><span class="sxs-lookup"><span data-stu-id="c149c-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="c149c-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c149c-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c149c-190">1.5</span><span class="sxs-lookup"><span data-stu-id="c149c-190">1.5</span></span> |
|[<span data-ttu-id="c149c-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c149c-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c149c-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c149c-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c149c-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c149c-193">SourceProperty :String</span></span>

<span data-ttu-id="c149c-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c149c-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c149c-195">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-195">Type</span></span>

*   <span data-ttu-id="c149c-196">String</span><span class="sxs-lookup"><span data-stu-id="c149c-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c149c-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c149c-197">Properties:</span></span>

|<span data-ttu-id="c149c-198">Имя</span><span class="sxs-lookup"><span data-stu-id="c149c-198">Name</span></span>| <span data-ttu-id="c149c-199">Тип</span><span class="sxs-lookup"><span data-stu-id="c149c-199">Type</span></span>| <span data-ttu-id="c149c-200">Описание</span><span class="sxs-lookup"><span data-stu-id="c149c-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c149c-201">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-201">String</span></span>|<span data-ttu-id="c149c-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c149c-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c149c-203">Строка</span><span class="sxs-lookup"><span data-stu-id="c149c-203">String</span></span>|<span data-ttu-id="c149c-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c149c-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c149c-205">Требования</span><span class="sxs-lookup"><span data-stu-id="c149c-205">Requirements</span></span>

|<span data-ttu-id="c149c-206">Требование</span><span class="sxs-lookup"><span data-stu-id="c149c-206">Requirement</span></span>| <span data-ttu-id="c149c-207">Значение</span><span class="sxs-lookup"><span data-stu-id="c149c-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="c149c-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c149c-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c149c-209">1.0</span><span class="sxs-lookup"><span data-stu-id="c149c-209">1.0</span></span>|
|[<span data-ttu-id="c149c-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c149c-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c149c-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c149c-211">Compose or Read</span></span>|
