---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: b65a9b0dd4523423a52e08a725e652e1740a779b
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064461"
---
# <a name="office"></a><span data-ttu-id="835f0-102">Office</span><span class="sxs-lookup"><span data-stu-id="835f0-102">Office</span></span>

<span data-ttu-id="835f0-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="835f0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="835f0-105">Требования</span><span class="sxs-lookup"><span data-stu-id="835f0-105">Requirements</span></span>

|<span data-ttu-id="835f0-106">Требование</span><span class="sxs-lookup"><span data-stu-id="835f0-106">Requirement</span></span>| <span data-ttu-id="835f0-107">Значение</span><span class="sxs-lookup"><span data-stu-id="835f0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="835f0-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="835f0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="835f0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="835f0-109">1.0</span></span>|
|[<span data-ttu-id="835f0-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="835f0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="835f0-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="835f0-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="835f0-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="835f0-112">Members and methods</span></span>

| <span data-ttu-id="835f0-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="835f0-113">Member</span></span> | <span data-ttu-id="835f0-114">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="835f0-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="835f0-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="835f0-116">Member</span><span class="sxs-lookup"><span data-stu-id="835f0-116">Member</span></span> |
| [<span data-ttu-id="835f0-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="835f0-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="835f0-118">Member</span><span class="sxs-lookup"><span data-stu-id="835f0-118">Member</span></span> |
| [<span data-ttu-id="835f0-119">EventType</span><span class="sxs-lookup"><span data-stu-id="835f0-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="835f0-120">Member</span><span class="sxs-lookup"><span data-stu-id="835f0-120">Member</span></span> |
| [<span data-ttu-id="835f0-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="835f0-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="835f0-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="835f0-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="835f0-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="835f0-123">Namespaces</span></span>

<span data-ttu-id="835f0-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="835f0-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="835f0-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="835f0-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="835f0-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="835f0-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="835f0-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="835f0-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="835f0-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="835f0-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="835f0-129">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-129">Type</span></span>

*   <span data-ttu-id="835f0-130">String</span><span class="sxs-lookup"><span data-stu-id="835f0-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="835f0-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="835f0-131">Properties:</span></span>

|<span data-ttu-id="835f0-132">Имя</span><span class="sxs-lookup"><span data-stu-id="835f0-132">Name</span></span>| <span data-ttu-id="835f0-133">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-133">Type</span></span>| <span data-ttu-id="835f0-134">Описание</span><span class="sxs-lookup"><span data-stu-id="835f0-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="835f0-135">String</span><span class="sxs-lookup"><span data-stu-id="835f0-135">String</span></span>|<span data-ttu-id="835f0-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="835f0-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="835f0-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="835f0-137">String</span></span>|<span data-ttu-id="835f0-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="835f0-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="835f0-139">Требования</span><span class="sxs-lookup"><span data-stu-id="835f0-139">Requirements</span></span>

|<span data-ttu-id="835f0-140">Требование</span><span class="sxs-lookup"><span data-stu-id="835f0-140">Requirement</span></span>| <span data-ttu-id="835f0-141">Значение</span><span class="sxs-lookup"><span data-stu-id="835f0-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="835f0-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="835f0-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="835f0-143">1.0</span><span class="sxs-lookup"><span data-stu-id="835f0-143">1.0</span></span>|
|[<span data-ttu-id="835f0-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="835f0-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="835f0-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="835f0-145">Compose or Read</span></span>|

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="835f0-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="835f0-146">CoercionType: String</span></span>

<span data-ttu-id="835f0-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="835f0-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="835f0-148">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-148">Type</span></span>

*   <span data-ttu-id="835f0-149">String</span><span class="sxs-lookup"><span data-stu-id="835f0-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="835f0-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="835f0-150">Properties:</span></span>

|<span data-ttu-id="835f0-151">Имя</span><span class="sxs-lookup"><span data-stu-id="835f0-151">Name</span></span>| <span data-ttu-id="835f0-152">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-152">Type</span></span>| <span data-ttu-id="835f0-153">Описание</span><span class="sxs-lookup"><span data-stu-id="835f0-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="835f0-154">String</span><span class="sxs-lookup"><span data-stu-id="835f0-154">String</span></span>|<span data-ttu-id="835f0-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="835f0-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="835f0-156">String</span><span class="sxs-lookup"><span data-stu-id="835f0-156">String</span></span>|<span data-ttu-id="835f0-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="835f0-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="835f0-158">Требования</span><span class="sxs-lookup"><span data-stu-id="835f0-158">Requirements</span></span>

|<span data-ttu-id="835f0-159">Требование</span><span class="sxs-lookup"><span data-stu-id="835f0-159">Requirement</span></span>| <span data-ttu-id="835f0-160">Значение</span><span class="sxs-lookup"><span data-stu-id="835f0-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="835f0-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="835f0-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="835f0-162">1.0</span><span class="sxs-lookup"><span data-stu-id="835f0-162">1.0</span></span>|
|[<span data-ttu-id="835f0-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="835f0-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="835f0-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="835f0-164">Compose or Read</span></span>|

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="835f0-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="835f0-165">EventType: String</span></span>

<span data-ttu-id="835f0-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="835f0-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="835f0-167">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-167">Type</span></span>

*   <span data-ttu-id="835f0-168">String</span><span class="sxs-lookup"><span data-stu-id="835f0-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="835f0-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="835f0-169">Properties:</span></span>

| <span data-ttu-id="835f0-170">Имя</span><span class="sxs-lookup"><span data-stu-id="835f0-170">Name</span></span> | <span data-ttu-id="835f0-171">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-171">Type</span></span> | <span data-ttu-id="835f0-172">Описание</span><span class="sxs-lookup"><span data-stu-id="835f0-172">Description</span></span> | <span data-ttu-id="835f0-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="835f0-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="835f0-174">String</span><span class="sxs-lookup"><span data-stu-id="835f0-174">String</span></span> | <span data-ttu-id="835f0-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="835f0-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="835f0-176">1.7</span><span class="sxs-lookup"><span data-stu-id="835f0-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="835f0-177">String</span><span class="sxs-lookup"><span data-stu-id="835f0-177">String</span></span> | <span data-ttu-id="835f0-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="835f0-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="835f0-179">1.5</span><span class="sxs-lookup"><span data-stu-id="835f0-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="835f0-180">String</span><span class="sxs-lookup"><span data-stu-id="835f0-180">String</span></span> | <span data-ttu-id="835f0-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="835f0-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="835f0-182">1.7</span><span class="sxs-lookup"><span data-stu-id="835f0-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="835f0-183">String</span><span class="sxs-lookup"><span data-stu-id="835f0-183">String</span></span> | <span data-ttu-id="835f0-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="835f0-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="835f0-185">1.7</span><span class="sxs-lookup"><span data-stu-id="835f0-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="835f0-186">Требования</span><span class="sxs-lookup"><span data-stu-id="835f0-186">Requirements</span></span>

|<span data-ttu-id="835f0-187">Требование</span><span class="sxs-lookup"><span data-stu-id="835f0-187">Requirement</span></span>| <span data-ttu-id="835f0-188">Значение</span><span class="sxs-lookup"><span data-stu-id="835f0-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="835f0-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="835f0-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="835f0-190">1.5</span><span class="sxs-lookup"><span data-stu-id="835f0-190">1.5</span></span> |
|[<span data-ttu-id="835f0-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="835f0-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="835f0-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="835f0-192">Compose or Read</span></span> |

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="835f0-193">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="835f0-193">SourceProperty: String</span></span>

<span data-ttu-id="835f0-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="835f0-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="835f0-195">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-195">Type</span></span>

*   <span data-ttu-id="835f0-196">String</span><span class="sxs-lookup"><span data-stu-id="835f0-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="835f0-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="835f0-197">Properties:</span></span>

|<span data-ttu-id="835f0-198">Имя</span><span class="sxs-lookup"><span data-stu-id="835f0-198">Name</span></span>| <span data-ttu-id="835f0-199">Тип</span><span class="sxs-lookup"><span data-stu-id="835f0-199">Type</span></span>| <span data-ttu-id="835f0-200">Описание</span><span class="sxs-lookup"><span data-stu-id="835f0-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="835f0-201">String</span><span class="sxs-lookup"><span data-stu-id="835f0-201">String</span></span>|<span data-ttu-id="835f0-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="835f0-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="835f0-203">String</span><span class="sxs-lookup"><span data-stu-id="835f0-203">String</span></span>|<span data-ttu-id="835f0-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="835f0-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="835f0-205">Требования</span><span class="sxs-lookup"><span data-stu-id="835f0-205">Requirements</span></span>

|<span data-ttu-id="835f0-206">Требование</span><span class="sxs-lookup"><span data-stu-id="835f0-206">Requirement</span></span>| <span data-ttu-id="835f0-207">Значение</span><span class="sxs-lookup"><span data-stu-id="835f0-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="835f0-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="835f0-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="835f0-209">1.0</span><span class="sxs-lookup"><span data-stu-id="835f0-209">1.0</span></span>|
|[<span data-ttu-id="835f0-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="835f0-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="835f0-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="835f0-211">Compose or Read</span></span>|
