---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871418"
---
# <a name="office"></a><span data-ttu-id="53cdf-102">Office</span><span class="sxs-lookup"><span data-stu-id="53cdf-102">Office</span></span>

<span data-ttu-id="53cdf-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="53cdf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="53cdf-105">Требования</span><span class="sxs-lookup"><span data-stu-id="53cdf-105">Requirements</span></span>

|<span data-ttu-id="53cdf-106">Требование</span><span class="sxs-lookup"><span data-stu-id="53cdf-106">Requirement</span></span>| <span data-ttu-id="53cdf-107">Значение</span><span class="sxs-lookup"><span data-stu-id="53cdf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="53cdf-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53cdf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53cdf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="53cdf-109">1.0</span></span>|
|[<span data-ttu-id="53cdf-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53cdf-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53cdf-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53cdf-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="53cdf-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="53cdf-112">Members and methods</span></span>

| <span data-ttu-id="53cdf-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="53cdf-113">Member</span></span> | <span data-ttu-id="53cdf-114">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="53cdf-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="53cdf-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="53cdf-116">Member</span><span class="sxs-lookup"><span data-stu-id="53cdf-116">Member</span></span> |
| [<span data-ttu-id="53cdf-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="53cdf-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="53cdf-118">Member</span><span class="sxs-lookup"><span data-stu-id="53cdf-118">Member</span></span> |
| [<span data-ttu-id="53cdf-119">EventType</span><span class="sxs-lookup"><span data-stu-id="53cdf-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="53cdf-120">Member</span><span class="sxs-lookup"><span data-stu-id="53cdf-120">Member</span></span> |
| [<span data-ttu-id="53cdf-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="53cdf-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="53cdf-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="53cdf-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="53cdf-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="53cdf-123">Namespaces</span></span>

<span data-ttu-id="53cdf-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="53cdf-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="53cdf-125">[MailboxEnums.](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="53cdf-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="53cdf-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="53cdf-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="53cdf-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="53cdf-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="53cdf-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="53cdf-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="53cdf-129">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-129">Type</span></span>

*   <span data-ttu-id="53cdf-130">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53cdf-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53cdf-131">Properties:</span></span>

|<span data-ttu-id="53cdf-132">Имя</span><span class="sxs-lookup"><span data-stu-id="53cdf-132">Name</span></span>| <span data-ttu-id="53cdf-133">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-133">Type</span></span>| <span data-ttu-id="53cdf-134">Описание</span><span class="sxs-lookup"><span data-stu-id="53cdf-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="53cdf-135">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-135">String</span></span>|<span data-ttu-id="53cdf-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="53cdf-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="53cdf-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="53cdf-137">String</span></span>|<span data-ttu-id="53cdf-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="53cdf-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53cdf-139">Требования</span><span class="sxs-lookup"><span data-stu-id="53cdf-139">Requirements</span></span>

|<span data-ttu-id="53cdf-140">Требование</span><span class="sxs-lookup"><span data-stu-id="53cdf-140">Requirement</span></span>| <span data-ttu-id="53cdf-141">Значение</span><span class="sxs-lookup"><span data-stu-id="53cdf-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="53cdf-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53cdf-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53cdf-143">1.0</span><span class="sxs-lookup"><span data-stu-id="53cdf-143">1.0</span></span>|
|[<span data-ttu-id="53cdf-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53cdf-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53cdf-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53cdf-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="53cdf-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="53cdf-146">CoercionType :String</span></span>

<span data-ttu-id="53cdf-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="53cdf-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53cdf-148">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-148">Type</span></span>

*   <span data-ttu-id="53cdf-149">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53cdf-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53cdf-150">Properties:</span></span>

|<span data-ttu-id="53cdf-151">Имя</span><span class="sxs-lookup"><span data-stu-id="53cdf-151">Name</span></span>| <span data-ttu-id="53cdf-152">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-152">Type</span></span>| <span data-ttu-id="53cdf-153">Описание</span><span class="sxs-lookup"><span data-stu-id="53cdf-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="53cdf-154">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-154">String</span></span>|<span data-ttu-id="53cdf-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="53cdf-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="53cdf-156">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-156">String</span></span>|<span data-ttu-id="53cdf-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="53cdf-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53cdf-158">Требования</span><span class="sxs-lookup"><span data-stu-id="53cdf-158">Requirements</span></span>

|<span data-ttu-id="53cdf-159">Требование</span><span class="sxs-lookup"><span data-stu-id="53cdf-159">Requirement</span></span>| <span data-ttu-id="53cdf-160">Значение</span><span class="sxs-lookup"><span data-stu-id="53cdf-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="53cdf-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53cdf-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53cdf-162">1.0</span><span class="sxs-lookup"><span data-stu-id="53cdf-162">1.0</span></span>|
|[<span data-ttu-id="53cdf-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53cdf-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53cdf-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53cdf-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="53cdf-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="53cdf-165">EventType :String</span></span>

<span data-ttu-id="53cdf-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="53cdf-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="53cdf-167">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-167">Type</span></span>

*   <span data-ttu-id="53cdf-168">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53cdf-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53cdf-169">Properties:</span></span>

| <span data-ttu-id="53cdf-170">Имя</span><span class="sxs-lookup"><span data-stu-id="53cdf-170">Name</span></span> | <span data-ttu-id="53cdf-171">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-171">Type</span></span> | <span data-ttu-id="53cdf-172">Описание</span><span class="sxs-lookup"><span data-stu-id="53cdf-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="53cdf-173">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-173">String</span></span> | <span data-ttu-id="53cdf-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="53cdf-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="53cdf-175">Требования</span><span class="sxs-lookup"><span data-stu-id="53cdf-175">Requirements</span></span>

|<span data-ttu-id="53cdf-176">Требование</span><span class="sxs-lookup"><span data-stu-id="53cdf-176">Requirement</span></span>| <span data-ttu-id="53cdf-177">Значение</span><span class="sxs-lookup"><span data-stu-id="53cdf-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="53cdf-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="53cdf-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53cdf-179">1.5</span><span class="sxs-lookup"><span data-stu-id="53cdf-179">1.5</span></span> |
|[<span data-ttu-id="53cdf-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53cdf-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53cdf-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53cdf-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="53cdf-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="53cdf-182">SourceProperty :String</span></span>

<span data-ttu-id="53cdf-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="53cdf-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53cdf-184">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-184">Type</span></span>

*   <span data-ttu-id="53cdf-185">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53cdf-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="53cdf-186">Properties:</span></span>

|<span data-ttu-id="53cdf-187">Имя</span><span class="sxs-lookup"><span data-stu-id="53cdf-187">Name</span></span>| <span data-ttu-id="53cdf-188">Тип</span><span class="sxs-lookup"><span data-stu-id="53cdf-188">Type</span></span>| <span data-ttu-id="53cdf-189">Описание</span><span class="sxs-lookup"><span data-stu-id="53cdf-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="53cdf-190">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-190">String</span></span>|<span data-ttu-id="53cdf-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="53cdf-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="53cdf-192">String</span><span class="sxs-lookup"><span data-stu-id="53cdf-192">String</span></span>|<span data-ttu-id="53cdf-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="53cdf-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53cdf-194">Требования</span><span class="sxs-lookup"><span data-stu-id="53cdf-194">Requirements</span></span>

|<span data-ttu-id="53cdf-195">Требование</span><span class="sxs-lookup"><span data-stu-id="53cdf-195">Requirement</span></span>| <span data-ttu-id="53cdf-196">Значение</span><span class="sxs-lookup"><span data-stu-id="53cdf-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="53cdf-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53cdf-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53cdf-198">1.0</span><span class="sxs-lookup"><span data-stu-id="53cdf-198">1.0</span></span>|
|[<span data-ttu-id="53cdf-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53cdf-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53cdf-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53cdf-200">Compose or Read</span></span>|
