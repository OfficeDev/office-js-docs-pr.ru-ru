---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dde96f48863459da5072d6b4864169f198264133
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870809"
---
# <a name="office"></a><span data-ttu-id="94382-102">Office</span><span class="sxs-lookup"><span data-stu-id="94382-102">Office</span></span>

<span data-ttu-id="94382-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="94382-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="94382-105">Требования</span><span class="sxs-lookup"><span data-stu-id="94382-105">Requirements</span></span>

|<span data-ttu-id="94382-106">Требование</span><span class="sxs-lookup"><span data-stu-id="94382-106">Requirement</span></span>| <span data-ttu-id="94382-107">Значение</span><span class="sxs-lookup"><span data-stu-id="94382-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="94382-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="94382-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94382-109">1.0</span><span class="sxs-lookup"><span data-stu-id="94382-109">1.0</span></span>|
|[<span data-ttu-id="94382-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="94382-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94382-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="94382-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="94382-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="94382-112">Members and methods</span></span>

| <span data-ttu-id="94382-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="94382-113">Member</span></span> | <span data-ttu-id="94382-114">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="94382-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="94382-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="94382-116">Member</span><span class="sxs-lookup"><span data-stu-id="94382-116">Member</span></span> |
| [<span data-ttu-id="94382-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="94382-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="94382-118">Member</span><span class="sxs-lookup"><span data-stu-id="94382-118">Member</span></span> |
| [<span data-ttu-id="94382-119">EventType</span><span class="sxs-lookup"><span data-stu-id="94382-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="94382-120">Member</span><span class="sxs-lookup"><span data-stu-id="94382-120">Member</span></span> |
| [<span data-ttu-id="94382-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="94382-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="94382-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="94382-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="94382-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="94382-123">Namespaces</span></span>

<span data-ttu-id="94382-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="94382-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="94382-125">[MailboxEnums.](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="94382-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="94382-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="94382-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="94382-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="94382-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="94382-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="94382-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="94382-129">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-129">Type</span></span>

*   <span data-ttu-id="94382-130">String</span><span class="sxs-lookup"><span data-stu-id="94382-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94382-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="94382-131">Properties:</span></span>

|<span data-ttu-id="94382-132">Имя</span><span class="sxs-lookup"><span data-stu-id="94382-132">Name</span></span>| <span data-ttu-id="94382-133">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-133">Type</span></span>| <span data-ttu-id="94382-134">Описание</span><span class="sxs-lookup"><span data-stu-id="94382-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="94382-135">String</span><span class="sxs-lookup"><span data-stu-id="94382-135">String</span></span>|<span data-ttu-id="94382-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="94382-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="94382-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="94382-137">String</span></span>|<span data-ttu-id="94382-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="94382-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94382-139">Требования</span><span class="sxs-lookup"><span data-stu-id="94382-139">Requirements</span></span>

|<span data-ttu-id="94382-140">Требование</span><span class="sxs-lookup"><span data-stu-id="94382-140">Requirement</span></span>| <span data-ttu-id="94382-141">Значение</span><span class="sxs-lookup"><span data-stu-id="94382-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="94382-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="94382-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94382-143">1.0</span><span class="sxs-lookup"><span data-stu-id="94382-143">1.0</span></span>|
|[<span data-ttu-id="94382-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="94382-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94382-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="94382-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="94382-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="94382-146">CoercionType :String</span></span>

<span data-ttu-id="94382-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="94382-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="94382-148">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-148">Type</span></span>

*   <span data-ttu-id="94382-149">String</span><span class="sxs-lookup"><span data-stu-id="94382-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94382-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="94382-150">Properties:</span></span>

|<span data-ttu-id="94382-151">Имя</span><span class="sxs-lookup"><span data-stu-id="94382-151">Name</span></span>| <span data-ttu-id="94382-152">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-152">Type</span></span>| <span data-ttu-id="94382-153">Описание</span><span class="sxs-lookup"><span data-stu-id="94382-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="94382-154">String</span><span class="sxs-lookup"><span data-stu-id="94382-154">String</span></span>|<span data-ttu-id="94382-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="94382-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="94382-156">String</span><span class="sxs-lookup"><span data-stu-id="94382-156">String</span></span>|<span data-ttu-id="94382-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="94382-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94382-158">Требования</span><span class="sxs-lookup"><span data-stu-id="94382-158">Requirements</span></span>

|<span data-ttu-id="94382-159">Требование</span><span class="sxs-lookup"><span data-stu-id="94382-159">Requirement</span></span>| <span data-ttu-id="94382-160">Значение</span><span class="sxs-lookup"><span data-stu-id="94382-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="94382-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="94382-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94382-162">1.0</span><span class="sxs-lookup"><span data-stu-id="94382-162">1.0</span></span>|
|[<span data-ttu-id="94382-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="94382-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94382-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="94382-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="94382-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="94382-165">EventType :String</span></span>

<span data-ttu-id="94382-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="94382-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="94382-167">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-167">Type</span></span>

*   <span data-ttu-id="94382-168">String</span><span class="sxs-lookup"><span data-stu-id="94382-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94382-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="94382-169">Properties:</span></span>

| <span data-ttu-id="94382-170">Имя</span><span class="sxs-lookup"><span data-stu-id="94382-170">Name</span></span> | <span data-ttu-id="94382-171">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-171">Type</span></span> | <span data-ttu-id="94382-172">Описание</span><span class="sxs-lookup"><span data-stu-id="94382-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="94382-173">String</span><span class="sxs-lookup"><span data-stu-id="94382-173">String</span></span> | <span data-ttu-id="94382-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="94382-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="94382-175">Требования</span><span class="sxs-lookup"><span data-stu-id="94382-175">Requirements</span></span>

|<span data-ttu-id="94382-176">Требование</span><span class="sxs-lookup"><span data-stu-id="94382-176">Requirement</span></span>| <span data-ttu-id="94382-177">Значение</span><span class="sxs-lookup"><span data-stu-id="94382-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="94382-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="94382-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94382-179">1.5</span><span class="sxs-lookup"><span data-stu-id="94382-179">1.5</span></span> |
|[<span data-ttu-id="94382-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="94382-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94382-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="94382-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="94382-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="94382-182">SourceProperty :String</span></span>

<span data-ttu-id="94382-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="94382-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="94382-184">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-184">Type</span></span>

*   <span data-ttu-id="94382-185">String</span><span class="sxs-lookup"><span data-stu-id="94382-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="94382-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="94382-186">Properties:</span></span>

|<span data-ttu-id="94382-187">Имя</span><span class="sxs-lookup"><span data-stu-id="94382-187">Name</span></span>| <span data-ttu-id="94382-188">Тип</span><span class="sxs-lookup"><span data-stu-id="94382-188">Type</span></span>| <span data-ttu-id="94382-189">Описание</span><span class="sxs-lookup"><span data-stu-id="94382-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="94382-190">String</span><span class="sxs-lookup"><span data-stu-id="94382-190">String</span></span>|<span data-ttu-id="94382-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="94382-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="94382-192">String</span><span class="sxs-lookup"><span data-stu-id="94382-192">String</span></span>|<span data-ttu-id="94382-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="94382-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="94382-194">Требования</span><span class="sxs-lookup"><span data-stu-id="94382-194">Requirements</span></span>

|<span data-ttu-id="94382-195">Требование</span><span class="sxs-lookup"><span data-stu-id="94382-195">Requirement</span></span>| <span data-ttu-id="94382-196">Значение</span><span class="sxs-lookup"><span data-stu-id="94382-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="94382-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="94382-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="94382-198">1.0</span><span class="sxs-lookup"><span data-stu-id="94382-198">1.0</span></span>|
|[<span data-ttu-id="94382-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="94382-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="94382-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="94382-200">Compose or Read</span></span>|
