---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 8d22ce8400916dffe12a15bba35f70ceca4db510
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695870"
---
# <a name="office"></a><span data-ttu-id="d7005-102">Office</span><span class="sxs-lookup"><span data-stu-id="d7005-102">Office</span></span>

<span data-ttu-id="d7005-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d7005-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7005-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d7005-105">Requirements</span></span>

|<span data-ttu-id="d7005-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d7005-106">Requirement</span></span>| <span data-ttu-id="d7005-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d7005-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7005-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d7005-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7005-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d7005-109">1.0</span></span>|
|[<span data-ttu-id="d7005-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7005-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7005-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7005-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d7005-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="d7005-112">Members and methods</span></span>

| <span data-ttu-id="d7005-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="d7005-113">Member</span></span> | <span data-ttu-id="d7005-114">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d7005-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d7005-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d7005-116">Member</span><span class="sxs-lookup"><span data-stu-id="d7005-116">Member</span></span> |
| [<span data-ttu-id="d7005-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7005-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d7005-118">Member</span><span class="sxs-lookup"><span data-stu-id="d7005-118">Member</span></span> |
| [<span data-ttu-id="d7005-119">EventType</span><span class="sxs-lookup"><span data-stu-id="d7005-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d7005-120">Member</span><span class="sxs-lookup"><span data-stu-id="d7005-120">Member</span></span> |
| [<span data-ttu-id="d7005-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d7005-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d7005-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="d7005-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d7005-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d7005-123">Namespaces</span></span>

<span data-ttu-id="d7005-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7005-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d7005-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="d7005-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="d7005-126">Members</span><span class="sxs-lookup"><span data-stu-id="d7005-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d7005-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="d7005-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="d7005-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="d7005-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d7005-129">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-129">Type</span></span>

*   <span data-ttu-id="d7005-130">String</span><span class="sxs-lookup"><span data-stu-id="d7005-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7005-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7005-131">Properties:</span></span>

|<span data-ttu-id="d7005-132">Имя</span><span class="sxs-lookup"><span data-stu-id="d7005-132">Name</span></span>| <span data-ttu-id="d7005-133">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-133">Type</span></span>| <span data-ttu-id="d7005-134">Описание</span><span class="sxs-lookup"><span data-stu-id="d7005-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d7005-135">String</span><span class="sxs-lookup"><span data-stu-id="d7005-135">String</span></span>|<span data-ttu-id="d7005-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="d7005-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d7005-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="d7005-137">String</span></span>|<span data-ttu-id="d7005-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="d7005-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7005-139">Требования</span><span class="sxs-lookup"><span data-stu-id="d7005-139">Requirements</span></span>

|<span data-ttu-id="d7005-140">Требование</span><span class="sxs-lookup"><span data-stu-id="d7005-140">Requirement</span></span>| <span data-ttu-id="d7005-141">Значение</span><span class="sxs-lookup"><span data-stu-id="d7005-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7005-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d7005-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7005-143">1.0</span><span class="sxs-lookup"><span data-stu-id="d7005-143">1.0</span></span>|
|[<span data-ttu-id="d7005-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7005-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7005-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7005-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d7005-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="d7005-146">CoercionType: String</span></span>

<span data-ttu-id="d7005-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d7005-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7005-148">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-148">Type</span></span>

*   <span data-ttu-id="d7005-149">String</span><span class="sxs-lookup"><span data-stu-id="d7005-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7005-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7005-150">Properties:</span></span>

|<span data-ttu-id="d7005-151">Имя</span><span class="sxs-lookup"><span data-stu-id="d7005-151">Name</span></span>| <span data-ttu-id="d7005-152">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-152">Type</span></span>| <span data-ttu-id="d7005-153">Описание</span><span class="sxs-lookup"><span data-stu-id="d7005-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d7005-154">String</span><span class="sxs-lookup"><span data-stu-id="d7005-154">String</span></span>|<span data-ttu-id="d7005-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d7005-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d7005-156">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-156">String</span></span>|<span data-ttu-id="d7005-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="d7005-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7005-158">Требования</span><span class="sxs-lookup"><span data-stu-id="d7005-158">Requirements</span></span>

|<span data-ttu-id="d7005-159">Требование</span><span class="sxs-lookup"><span data-stu-id="d7005-159">Requirement</span></span>| <span data-ttu-id="d7005-160">Значение</span><span class="sxs-lookup"><span data-stu-id="d7005-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7005-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d7005-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7005-162">1.0</span><span class="sxs-lookup"><span data-stu-id="d7005-162">1.0</span></span>|
|[<span data-ttu-id="d7005-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7005-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7005-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7005-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="d7005-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="d7005-165">EventType: String</span></span>

<span data-ttu-id="d7005-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="d7005-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d7005-167">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-167">Type</span></span>

*   <span data-ttu-id="d7005-168">String</span><span class="sxs-lookup"><span data-stu-id="d7005-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7005-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7005-169">Properties:</span></span>

| <span data-ttu-id="d7005-170">Имя</span><span class="sxs-lookup"><span data-stu-id="d7005-170">Name</span></span> | <span data-ttu-id="d7005-171">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-171">Type</span></span> | <span data-ttu-id="d7005-172">Описание</span><span class="sxs-lookup"><span data-stu-id="d7005-172">Description</span></span> | <span data-ttu-id="d7005-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="d7005-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="d7005-174">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-174">String</span></span> | <span data-ttu-id="d7005-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="d7005-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="d7005-176">1.7</span><span class="sxs-lookup"><span data-stu-id="d7005-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="d7005-177">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-177">String</span></span> | <span data-ttu-id="d7005-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="d7005-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d7005-179">1.5</span><span class="sxs-lookup"><span data-stu-id="d7005-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="d7005-180">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-180">String</span></span> | <span data-ttu-id="d7005-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="d7005-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="d7005-182">1.7</span><span class="sxs-lookup"><span data-stu-id="d7005-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="d7005-183">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-183">String</span></span> | <span data-ttu-id="d7005-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="d7005-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="d7005-185">1.7</span><span class="sxs-lookup"><span data-stu-id="d7005-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7005-186">Требования</span><span class="sxs-lookup"><span data-stu-id="d7005-186">Requirements</span></span>

|<span data-ttu-id="d7005-187">Требование</span><span class="sxs-lookup"><span data-stu-id="d7005-187">Requirement</span></span>| <span data-ttu-id="d7005-188">Значение</span><span class="sxs-lookup"><span data-stu-id="d7005-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7005-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d7005-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7005-190">1.5</span><span class="sxs-lookup"><span data-stu-id="d7005-190">1.5</span></span> |
|[<span data-ttu-id="d7005-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7005-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7005-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7005-192">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d7005-193">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="d7005-193">SourceProperty: String</span></span>

<span data-ttu-id="d7005-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d7005-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7005-195">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-195">Type</span></span>

*   <span data-ttu-id="d7005-196">String</span><span class="sxs-lookup"><span data-stu-id="d7005-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7005-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7005-197">Properties:</span></span>

|<span data-ttu-id="d7005-198">Имя</span><span class="sxs-lookup"><span data-stu-id="d7005-198">Name</span></span>| <span data-ttu-id="d7005-199">Тип</span><span class="sxs-lookup"><span data-stu-id="d7005-199">Type</span></span>| <span data-ttu-id="d7005-200">Описание</span><span class="sxs-lookup"><span data-stu-id="d7005-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d7005-201">String</span><span class="sxs-lookup"><span data-stu-id="d7005-201">String</span></span>|<span data-ttu-id="d7005-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="d7005-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d7005-203">String.</span><span class="sxs-lookup"><span data-stu-id="d7005-203">String</span></span>|<span data-ttu-id="d7005-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="d7005-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7005-205">Требования</span><span class="sxs-lookup"><span data-stu-id="d7005-205">Requirements</span></span>

|<span data-ttu-id="d7005-206">Требование</span><span class="sxs-lookup"><span data-stu-id="d7005-206">Requirement</span></span>| <span data-ttu-id="d7005-207">Значение</span><span class="sxs-lookup"><span data-stu-id="d7005-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7005-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d7005-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7005-209">1.0</span><span class="sxs-lookup"><span data-stu-id="d7005-209">1.0</span></span>|
|[<span data-ttu-id="d7005-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7005-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7005-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7005-211">Compose or Read</span></span>|
