---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 36faf4569ac58693dcc1218c42a19347816d9abd
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064680"
---
# <a name="office"></a><span data-ttu-id="49ebf-102">Office</span><span class="sxs-lookup"><span data-stu-id="49ebf-102">Office</span></span>

<span data-ttu-id="49ebf-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="49ebf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="49ebf-105">Требования</span><span class="sxs-lookup"><span data-stu-id="49ebf-105">Requirements</span></span>

|<span data-ttu-id="49ebf-106">Требование</span><span class="sxs-lookup"><span data-stu-id="49ebf-106">Requirement</span></span>| <span data-ttu-id="49ebf-107">Значение</span><span class="sxs-lookup"><span data-stu-id="49ebf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="49ebf-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="49ebf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="49ebf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="49ebf-109">1.0</span></span>|
|[<span data-ttu-id="49ebf-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="49ebf-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="49ebf-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="49ebf-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="49ebf-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="49ebf-112">Members and methods</span></span>

| <span data-ttu-id="49ebf-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="49ebf-113">Member</span></span> | <span data-ttu-id="49ebf-114">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="49ebf-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="49ebf-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="49ebf-116">Member</span><span class="sxs-lookup"><span data-stu-id="49ebf-116">Member</span></span> |
| [<span data-ttu-id="49ebf-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="49ebf-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="49ebf-118">Member</span><span class="sxs-lookup"><span data-stu-id="49ebf-118">Member</span></span> |
| [<span data-ttu-id="49ebf-119">EventType</span><span class="sxs-lookup"><span data-stu-id="49ebf-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="49ebf-120">Member</span><span class="sxs-lookup"><span data-stu-id="49ebf-120">Member</span></span> |
| [<span data-ttu-id="49ebf-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="49ebf-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="49ebf-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="49ebf-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="49ebf-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="49ebf-123">Namespaces</span></span>

<span data-ttu-id="49ebf-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="49ebf-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="49ebf-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="49ebf-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="49ebf-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="49ebf-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="49ebf-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="49ebf-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="49ebf-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="49ebf-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="49ebf-129">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-129">Type</span></span>

*   <span data-ttu-id="49ebf-130">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="49ebf-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="49ebf-131">Properties:</span></span>

|<span data-ttu-id="49ebf-132">Имя</span><span class="sxs-lookup"><span data-stu-id="49ebf-132">Name</span></span>| <span data-ttu-id="49ebf-133">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-133">Type</span></span>| <span data-ttu-id="49ebf-134">Описание</span><span class="sxs-lookup"><span data-stu-id="49ebf-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="49ebf-135">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-135">String</span></span>|<span data-ttu-id="49ebf-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="49ebf-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="49ebf-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="49ebf-137">String</span></span>|<span data-ttu-id="49ebf-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="49ebf-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="49ebf-139">Требования</span><span class="sxs-lookup"><span data-stu-id="49ebf-139">Requirements</span></span>

|<span data-ttu-id="49ebf-140">Требование</span><span class="sxs-lookup"><span data-stu-id="49ebf-140">Requirement</span></span>| <span data-ttu-id="49ebf-141">Значение</span><span class="sxs-lookup"><span data-stu-id="49ebf-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="49ebf-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="49ebf-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="49ebf-143">1.0</span><span class="sxs-lookup"><span data-stu-id="49ebf-143">1.0</span></span>|
|[<span data-ttu-id="49ebf-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="49ebf-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="49ebf-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="49ebf-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="49ebf-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="49ebf-146">CoercionType: String</span></span>

<span data-ttu-id="49ebf-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="49ebf-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="49ebf-148">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-148">Type</span></span>

*   <span data-ttu-id="49ebf-149">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="49ebf-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="49ebf-150">Properties:</span></span>

|<span data-ttu-id="49ebf-151">Имя</span><span class="sxs-lookup"><span data-stu-id="49ebf-151">Name</span></span>| <span data-ttu-id="49ebf-152">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-152">Type</span></span>| <span data-ttu-id="49ebf-153">Описание</span><span class="sxs-lookup"><span data-stu-id="49ebf-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="49ebf-154">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-154">String</span></span>|<span data-ttu-id="49ebf-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="49ebf-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="49ebf-156">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-156">String</span></span>|<span data-ttu-id="49ebf-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="49ebf-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="49ebf-158">Требования</span><span class="sxs-lookup"><span data-stu-id="49ebf-158">Requirements</span></span>

|<span data-ttu-id="49ebf-159">Требование</span><span class="sxs-lookup"><span data-stu-id="49ebf-159">Requirement</span></span>| <span data-ttu-id="49ebf-160">Значение</span><span class="sxs-lookup"><span data-stu-id="49ebf-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="49ebf-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="49ebf-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="49ebf-162">1.0</span><span class="sxs-lookup"><span data-stu-id="49ebf-162">1.0</span></span>|
|[<span data-ttu-id="49ebf-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="49ebf-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="49ebf-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="49ebf-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="49ebf-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="49ebf-165">EventType: String</span></span>

<span data-ttu-id="49ebf-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="49ebf-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="49ebf-167">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-167">Type</span></span>

*   <span data-ttu-id="49ebf-168">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="49ebf-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="49ebf-169">Properties:</span></span>

| <span data-ttu-id="49ebf-170">Имя</span><span class="sxs-lookup"><span data-stu-id="49ebf-170">Name</span></span> | <span data-ttu-id="49ebf-171">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-171">Type</span></span> | <span data-ttu-id="49ebf-172">Описание</span><span class="sxs-lookup"><span data-stu-id="49ebf-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="49ebf-173">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-173">String</span></span> | <span data-ttu-id="49ebf-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="49ebf-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="49ebf-175">Требования</span><span class="sxs-lookup"><span data-stu-id="49ebf-175">Requirements</span></span>

|<span data-ttu-id="49ebf-176">Требование</span><span class="sxs-lookup"><span data-stu-id="49ebf-176">Requirement</span></span>| <span data-ttu-id="49ebf-177">Значение</span><span class="sxs-lookup"><span data-stu-id="49ebf-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="49ebf-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="49ebf-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="49ebf-179">1.5</span><span class="sxs-lookup"><span data-stu-id="49ebf-179">1.5</span></span> |
|[<span data-ttu-id="49ebf-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="49ebf-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="49ebf-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="49ebf-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="49ebf-182">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="49ebf-182">SourceProperty: String</span></span>

<span data-ttu-id="49ebf-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="49ebf-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="49ebf-184">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-184">Type</span></span>

*   <span data-ttu-id="49ebf-185">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="49ebf-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="49ebf-186">Properties:</span></span>

|<span data-ttu-id="49ebf-187">Имя</span><span class="sxs-lookup"><span data-stu-id="49ebf-187">Name</span></span>| <span data-ttu-id="49ebf-188">Тип</span><span class="sxs-lookup"><span data-stu-id="49ebf-188">Type</span></span>| <span data-ttu-id="49ebf-189">Описание</span><span class="sxs-lookup"><span data-stu-id="49ebf-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="49ebf-190">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-190">String</span></span>|<span data-ttu-id="49ebf-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="49ebf-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="49ebf-192">String</span><span class="sxs-lookup"><span data-stu-id="49ebf-192">String</span></span>|<span data-ttu-id="49ebf-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="49ebf-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="49ebf-194">Требования</span><span class="sxs-lookup"><span data-stu-id="49ebf-194">Requirements</span></span>

|<span data-ttu-id="49ebf-195">Требование</span><span class="sxs-lookup"><span data-stu-id="49ebf-195">Requirement</span></span>| <span data-ttu-id="49ebf-196">Значение</span><span class="sxs-lookup"><span data-stu-id="49ebf-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="49ebf-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="49ebf-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="49ebf-198">1.0</span><span class="sxs-lookup"><span data-stu-id="49ebf-198">1.0</span></span>|
|[<span data-ttu-id="49ebf-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="49ebf-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="49ebf-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="49ebf-200">Compose or Read</span></span>|
