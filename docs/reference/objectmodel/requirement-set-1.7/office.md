---
title: Пространство имен Office — набор обязательных элементов 1,7
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: be0223e7ed274abf0e742be13f258c14f6dccf91
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395696"
---
# <a name="office"></a><span data-ttu-id="c7b63-102">Office</span><span class="sxs-lookup"><span data-stu-id="c7b63-102">Office</span></span>

<span data-ttu-id="c7b63-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c7b63-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7b63-105">Требования</span><span class="sxs-lookup"><span data-stu-id="c7b63-105">Requirements</span></span>

|<span data-ttu-id="c7b63-106">Требование</span><span class="sxs-lookup"><span data-stu-id="c7b63-106">Requirement</span></span>| <span data-ttu-id="c7b63-107">Значение</span><span class="sxs-lookup"><span data-stu-id="c7b63-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7b63-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c7b63-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7b63-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c7b63-109">1.0</span></span>|
|[<span data-ttu-id="c7b63-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c7b63-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7b63-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c7b63-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c7b63-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="c7b63-112">Members and methods</span></span>

| <span data-ttu-id="c7b63-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="c7b63-113">Member</span></span> | <span data-ttu-id="c7b63-114">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c7b63-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c7b63-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c7b63-116">Member</span><span class="sxs-lookup"><span data-stu-id="c7b63-116">Member</span></span> |
| [<span data-ttu-id="c7b63-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c7b63-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c7b63-118">Member</span><span class="sxs-lookup"><span data-stu-id="c7b63-118">Member</span></span> |
| [<span data-ttu-id="c7b63-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c7b63-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c7b63-120">Member</span><span class="sxs-lookup"><span data-stu-id="c7b63-120">Member</span></span> |
| [<span data-ttu-id="c7b63-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c7b63-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c7b63-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="c7b63-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c7b63-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c7b63-123">Namespaces</span></span>

<span data-ttu-id="c7b63-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="c7b63-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c7b63-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="c7b63-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.7): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="c7b63-126">Members</span><span class="sxs-lookup"><span data-stu-id="c7b63-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c7b63-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="c7b63-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="c7b63-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c7b63-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c7b63-129">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-129">Type</span></span>

*   <span data-ttu-id="c7b63-130">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7b63-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c7b63-131">Properties:</span></span>

|<span data-ttu-id="c7b63-132">Имя</span><span class="sxs-lookup"><span data-stu-id="c7b63-132">Name</span></span>| <span data-ttu-id="c7b63-133">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-133">Type</span></span>| <span data-ttu-id="c7b63-134">Описание</span><span class="sxs-lookup"><span data-stu-id="c7b63-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c7b63-135">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-135">String</span></span>|<span data-ttu-id="c7b63-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c7b63-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c7b63-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="c7b63-137">String</span></span>|<span data-ttu-id="c7b63-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="c7b63-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7b63-139">Требования</span><span class="sxs-lookup"><span data-stu-id="c7b63-139">Requirements</span></span>

|<span data-ttu-id="c7b63-140">Требование</span><span class="sxs-lookup"><span data-stu-id="c7b63-140">Requirement</span></span>| <span data-ttu-id="c7b63-141">Значение</span><span class="sxs-lookup"><span data-stu-id="c7b63-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7b63-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c7b63-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7b63-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c7b63-143">1.0</span></span>|
|[<span data-ttu-id="c7b63-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c7b63-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7b63-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c7b63-145">Compose or Read</span></span>|

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="c7b63-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="c7b63-146">CoercionType: String</span></span>

<span data-ttu-id="c7b63-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c7b63-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7b63-148">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-148">Type</span></span>

*   <span data-ttu-id="c7b63-149">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7b63-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c7b63-150">Properties:</span></span>

|<span data-ttu-id="c7b63-151">Имя</span><span class="sxs-lookup"><span data-stu-id="c7b63-151">Name</span></span>| <span data-ttu-id="c7b63-152">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-152">Type</span></span>| <span data-ttu-id="c7b63-153">Описание</span><span class="sxs-lookup"><span data-stu-id="c7b63-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c7b63-154">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-154">String</span></span>|<span data-ttu-id="c7b63-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c7b63-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c7b63-156">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-156">String</span></span>|<span data-ttu-id="c7b63-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c7b63-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7b63-158">Требования</span><span class="sxs-lookup"><span data-stu-id="c7b63-158">Requirements</span></span>

|<span data-ttu-id="c7b63-159">Требование</span><span class="sxs-lookup"><span data-stu-id="c7b63-159">Requirement</span></span>| <span data-ttu-id="c7b63-160">Значение</span><span class="sxs-lookup"><span data-stu-id="c7b63-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7b63-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c7b63-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7b63-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c7b63-162">1.0</span></span>|
|[<span data-ttu-id="c7b63-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c7b63-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7b63-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c7b63-164">Compose or Read</span></span>|

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="c7b63-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="c7b63-165">EventType: String</span></span>

<span data-ttu-id="c7b63-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="c7b63-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c7b63-167">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-167">Type</span></span>

*   <span data-ttu-id="c7b63-168">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7b63-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c7b63-169">Properties:</span></span>

| <span data-ttu-id="c7b63-170">Имя</span><span class="sxs-lookup"><span data-stu-id="c7b63-170">Name</span></span> | <span data-ttu-id="c7b63-171">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-171">Type</span></span> | <span data-ttu-id="c7b63-172">Описание</span><span class="sxs-lookup"><span data-stu-id="c7b63-172">Description</span></span> | <span data-ttu-id="c7b63-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="c7b63-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="c7b63-174">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-174">String</span></span> | <span data-ttu-id="c7b63-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="c7b63-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="c7b63-176">1.7</span><span class="sxs-lookup"><span data-stu-id="c7b63-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="c7b63-177">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-177">String</span></span> | <span data-ttu-id="c7b63-178">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="c7b63-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="c7b63-179">1.5</span><span class="sxs-lookup"><span data-stu-id="c7b63-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="c7b63-180">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-180">String</span></span> | <span data-ttu-id="c7b63-181">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="c7b63-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="c7b63-182">1.7</span><span class="sxs-lookup"><span data-stu-id="c7b63-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="c7b63-183">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-183">String</span></span> | <span data-ttu-id="c7b63-184">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="c7b63-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="c7b63-185">1.7</span><span class="sxs-lookup"><span data-stu-id="c7b63-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c7b63-186">Требования</span><span class="sxs-lookup"><span data-stu-id="c7b63-186">Requirements</span></span>

|<span data-ttu-id="c7b63-187">Требование</span><span class="sxs-lookup"><span data-stu-id="c7b63-187">Requirement</span></span>| <span data-ttu-id="c7b63-188">Значение</span><span class="sxs-lookup"><span data-stu-id="c7b63-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7b63-189">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c7b63-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7b63-190">1.5</span><span class="sxs-lookup"><span data-stu-id="c7b63-190">1.5</span></span> |
|[<span data-ttu-id="c7b63-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c7b63-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7b63-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c7b63-192">Compose or Read</span></span> |

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="c7b63-193">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="c7b63-193">SourceProperty: String</span></span>

<span data-ttu-id="c7b63-194">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c7b63-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c7b63-195">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-195">Type</span></span>

*   <span data-ttu-id="c7b63-196">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c7b63-197">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c7b63-197">Properties:</span></span>

|<span data-ttu-id="c7b63-198">Имя</span><span class="sxs-lookup"><span data-stu-id="c7b63-198">Name</span></span>| <span data-ttu-id="c7b63-199">Тип</span><span class="sxs-lookup"><span data-stu-id="c7b63-199">Type</span></span>| <span data-ttu-id="c7b63-200">Описание</span><span class="sxs-lookup"><span data-stu-id="c7b63-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c7b63-201">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-201">String</span></span>|<span data-ttu-id="c7b63-202">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c7b63-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c7b63-203">String</span><span class="sxs-lookup"><span data-stu-id="c7b63-203">String</span></span>|<span data-ttu-id="c7b63-204">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c7b63-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7b63-205">Требования</span><span class="sxs-lookup"><span data-stu-id="c7b63-205">Requirements</span></span>

|<span data-ttu-id="c7b63-206">Требование</span><span class="sxs-lookup"><span data-stu-id="c7b63-206">Requirement</span></span>| <span data-ttu-id="c7b63-207">Значение</span><span class="sxs-lookup"><span data-stu-id="c7b63-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7b63-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c7b63-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7b63-209">1.0</span><span class="sxs-lookup"><span data-stu-id="c7b63-209">1.0</span></span>|
|[<span data-ttu-id="c7b63-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c7b63-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c7b63-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c7b63-211">Compose or Read</span></span>|
