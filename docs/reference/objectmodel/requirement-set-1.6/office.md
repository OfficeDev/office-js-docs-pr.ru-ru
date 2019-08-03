---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: e211a3a2983567b79b73a791914f8d4ed1501ab1
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064665"
---
# <a name="office"></a><span data-ttu-id="67a8c-102">Office</span><span class="sxs-lookup"><span data-stu-id="67a8c-102">Office</span></span>

<span data-ttu-id="67a8c-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="67a8c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="67a8c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="67a8c-105">Requirements</span></span>

|<span data-ttu-id="67a8c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="67a8c-106">Requirement</span></span>| <span data-ttu-id="67a8c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="67a8c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="67a8c-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67a8c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67a8c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="67a8c-109">1.0</span></span>|
|[<span data-ttu-id="67a8c-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67a8c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67a8c-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67a8c-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="67a8c-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="67a8c-112">Members and methods</span></span>

| <span data-ttu-id="67a8c-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="67a8c-113">Member</span></span> | <span data-ttu-id="67a8c-114">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="67a8c-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="67a8c-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="67a8c-116">Member</span><span class="sxs-lookup"><span data-stu-id="67a8c-116">Member</span></span> |
| [<span data-ttu-id="67a8c-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="67a8c-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="67a8c-118">Member</span><span class="sxs-lookup"><span data-stu-id="67a8c-118">Member</span></span> |
| [<span data-ttu-id="67a8c-119">EventType</span><span class="sxs-lookup"><span data-stu-id="67a8c-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="67a8c-120">Member</span><span class="sxs-lookup"><span data-stu-id="67a8c-120">Member</span></span> |
| [<span data-ttu-id="67a8c-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="67a8c-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="67a8c-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="67a8c-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="67a8c-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="67a8c-123">Namespaces</span></span>

<span data-ttu-id="67a8c-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="67a8c-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="67a8c-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="67a8c-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="67a8c-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="67a8c-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="67a8c-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="67a8c-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="67a8c-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="67a8c-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="67a8c-129">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-129">Type</span></span>

*   <span data-ttu-id="67a8c-130">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="67a8c-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="67a8c-131">Properties:</span></span>

|<span data-ttu-id="67a8c-132">Имя</span><span class="sxs-lookup"><span data-stu-id="67a8c-132">Name</span></span>| <span data-ttu-id="67a8c-133">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-133">Type</span></span>| <span data-ttu-id="67a8c-134">Описание</span><span class="sxs-lookup"><span data-stu-id="67a8c-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="67a8c-135">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-135">String</span></span>|<span data-ttu-id="67a8c-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="67a8c-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="67a8c-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="67a8c-137">String</span></span>|<span data-ttu-id="67a8c-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="67a8c-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67a8c-139">Требования</span><span class="sxs-lookup"><span data-stu-id="67a8c-139">Requirements</span></span>

|<span data-ttu-id="67a8c-140">Требование</span><span class="sxs-lookup"><span data-stu-id="67a8c-140">Requirement</span></span>| <span data-ttu-id="67a8c-141">Значение</span><span class="sxs-lookup"><span data-stu-id="67a8c-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="67a8c-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67a8c-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67a8c-143">1.0</span><span class="sxs-lookup"><span data-stu-id="67a8c-143">1.0</span></span>|
|[<span data-ttu-id="67a8c-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67a8c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67a8c-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67a8c-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="67a8c-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="67a8c-146">CoercionType: String</span></span>

<span data-ttu-id="67a8c-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="67a8c-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="67a8c-148">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-148">Type</span></span>

*   <span data-ttu-id="67a8c-149">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="67a8c-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="67a8c-150">Properties:</span></span>

|<span data-ttu-id="67a8c-151">Имя</span><span class="sxs-lookup"><span data-stu-id="67a8c-151">Name</span></span>| <span data-ttu-id="67a8c-152">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-152">Type</span></span>| <span data-ttu-id="67a8c-153">Описание</span><span class="sxs-lookup"><span data-stu-id="67a8c-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="67a8c-154">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-154">String</span></span>|<span data-ttu-id="67a8c-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="67a8c-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="67a8c-156">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-156">String</span></span>|<span data-ttu-id="67a8c-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="67a8c-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67a8c-158">Требования</span><span class="sxs-lookup"><span data-stu-id="67a8c-158">Requirements</span></span>

|<span data-ttu-id="67a8c-159">Требование</span><span class="sxs-lookup"><span data-stu-id="67a8c-159">Requirement</span></span>| <span data-ttu-id="67a8c-160">Значение</span><span class="sxs-lookup"><span data-stu-id="67a8c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="67a8c-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67a8c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67a8c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="67a8c-162">1.0</span></span>|
|[<span data-ttu-id="67a8c-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67a8c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67a8c-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67a8c-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="67a8c-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="67a8c-165">EventType: String</span></span>

<span data-ttu-id="67a8c-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="67a8c-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="67a8c-167">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-167">Type</span></span>

*   <span data-ttu-id="67a8c-168">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="67a8c-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="67a8c-169">Properties:</span></span>

| <span data-ttu-id="67a8c-170">Имя</span><span class="sxs-lookup"><span data-stu-id="67a8c-170">Name</span></span> | <span data-ttu-id="67a8c-171">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-171">Type</span></span> | <span data-ttu-id="67a8c-172">Описание</span><span class="sxs-lookup"><span data-stu-id="67a8c-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="67a8c-173">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-173">String</span></span> | <span data-ttu-id="67a8c-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="67a8c-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="67a8c-175">Требования</span><span class="sxs-lookup"><span data-stu-id="67a8c-175">Requirements</span></span>

|<span data-ttu-id="67a8c-176">Требование</span><span class="sxs-lookup"><span data-stu-id="67a8c-176">Requirement</span></span>| <span data-ttu-id="67a8c-177">Значение</span><span class="sxs-lookup"><span data-stu-id="67a8c-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="67a8c-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="67a8c-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67a8c-179">1.5</span><span class="sxs-lookup"><span data-stu-id="67a8c-179">1.5</span></span> |
|[<span data-ttu-id="67a8c-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67a8c-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67a8c-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67a8c-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="67a8c-182">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="67a8c-182">SourceProperty: String</span></span>

<span data-ttu-id="67a8c-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="67a8c-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="67a8c-184">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-184">Type</span></span>

*   <span data-ttu-id="67a8c-185">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="67a8c-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="67a8c-186">Properties:</span></span>

|<span data-ttu-id="67a8c-187">Имя</span><span class="sxs-lookup"><span data-stu-id="67a8c-187">Name</span></span>| <span data-ttu-id="67a8c-188">Тип</span><span class="sxs-lookup"><span data-stu-id="67a8c-188">Type</span></span>| <span data-ttu-id="67a8c-189">Описание</span><span class="sxs-lookup"><span data-stu-id="67a8c-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="67a8c-190">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-190">String</span></span>|<span data-ttu-id="67a8c-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="67a8c-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="67a8c-192">String</span><span class="sxs-lookup"><span data-stu-id="67a8c-192">String</span></span>|<span data-ttu-id="67a8c-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="67a8c-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="67a8c-194">Требования</span><span class="sxs-lookup"><span data-stu-id="67a8c-194">Requirements</span></span>

|<span data-ttu-id="67a8c-195">Требование</span><span class="sxs-lookup"><span data-stu-id="67a8c-195">Requirement</span></span>| <span data-ttu-id="67a8c-196">Значение</span><span class="sxs-lookup"><span data-stu-id="67a8c-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="67a8c-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67a8c-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67a8c-198">1.0</span><span class="sxs-lookup"><span data-stu-id="67a8c-198">1.0</span></span>|
|[<span data-ttu-id="67a8c-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67a8c-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67a8c-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67a8c-200">Compose or Read</span></span>|
