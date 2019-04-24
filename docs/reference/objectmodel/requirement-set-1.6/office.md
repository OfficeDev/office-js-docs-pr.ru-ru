---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dde96f48863459da5072d6b4864169f198264133
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450375"
---
# <a name="office"></a><span data-ttu-id="b4467-102">Office</span><span class="sxs-lookup"><span data-stu-id="b4467-102">Office</span></span>

<span data-ttu-id="b4467-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b4467-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4467-105">Требования</span><span class="sxs-lookup"><span data-stu-id="b4467-105">Requirements</span></span>

|<span data-ttu-id="b4467-106">Требование</span><span class="sxs-lookup"><span data-stu-id="b4467-106">Requirement</span></span>| <span data-ttu-id="b4467-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b4467-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4467-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b4467-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4467-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b4467-109">1.0</span></span>|
|[<span data-ttu-id="b4467-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b4467-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b4467-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b4467-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b4467-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="b4467-112">Members and methods</span></span>

| <span data-ttu-id="b4467-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="b4467-113">Member</span></span> | <span data-ttu-id="b4467-114">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b4467-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b4467-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b4467-116">Member</span><span class="sxs-lookup"><span data-stu-id="b4467-116">Member</span></span> |
| [<span data-ttu-id="b4467-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b4467-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b4467-118">Member</span><span class="sxs-lookup"><span data-stu-id="b4467-118">Member</span></span> |
| [<span data-ttu-id="b4467-119">EventType</span><span class="sxs-lookup"><span data-stu-id="b4467-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b4467-120">Member</span><span class="sxs-lookup"><span data-stu-id="b4467-120">Member</span></span> |
| [<span data-ttu-id="b4467-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b4467-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b4467-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="b4467-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b4467-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="b4467-123">Namespaces</span></span>

<span data-ttu-id="b4467-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="b4467-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b4467-125">[MailboxEnums.](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b4467-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b4467-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="b4467-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b4467-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b4467-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="b4467-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="b4467-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b4467-129">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-129">Type</span></span>

*   <span data-ttu-id="b4467-130">String</span><span class="sxs-lookup"><span data-stu-id="b4467-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b4467-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b4467-131">Properties:</span></span>

|<span data-ttu-id="b4467-132">Имя</span><span class="sxs-lookup"><span data-stu-id="b4467-132">Name</span></span>| <span data-ttu-id="b4467-133">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-133">Type</span></span>| <span data-ttu-id="b4467-134">Описание</span><span class="sxs-lookup"><span data-stu-id="b4467-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b4467-135">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-135">String</span></span>|<span data-ttu-id="b4467-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="b4467-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b4467-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="b4467-137">String</span></span>|<span data-ttu-id="b4467-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="b4467-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4467-139">Требования</span><span class="sxs-lookup"><span data-stu-id="b4467-139">Requirements</span></span>

|<span data-ttu-id="b4467-140">Требование</span><span class="sxs-lookup"><span data-stu-id="b4467-140">Requirement</span></span>| <span data-ttu-id="b4467-141">Значение</span><span class="sxs-lookup"><span data-stu-id="b4467-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4467-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b4467-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4467-143">1.0</span><span class="sxs-lookup"><span data-stu-id="b4467-143">1.0</span></span>|
|[<span data-ttu-id="b4467-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b4467-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b4467-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b4467-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="b4467-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b4467-146">CoercionType :String</span></span>

<span data-ttu-id="b4467-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b4467-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b4467-148">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-148">Type</span></span>

*   <span data-ttu-id="b4467-149">String</span><span class="sxs-lookup"><span data-stu-id="b4467-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b4467-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b4467-150">Properties:</span></span>

|<span data-ttu-id="b4467-151">Имя</span><span class="sxs-lookup"><span data-stu-id="b4467-151">Name</span></span>| <span data-ttu-id="b4467-152">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-152">Type</span></span>| <span data-ttu-id="b4467-153">Описание</span><span class="sxs-lookup"><span data-stu-id="b4467-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b4467-154">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-154">String</span></span>|<span data-ttu-id="b4467-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="b4467-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b4467-156">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-156">String</span></span>|<span data-ttu-id="b4467-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="b4467-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4467-158">Требования</span><span class="sxs-lookup"><span data-stu-id="b4467-158">Requirements</span></span>

|<span data-ttu-id="b4467-159">Требование</span><span class="sxs-lookup"><span data-stu-id="b4467-159">Requirement</span></span>| <span data-ttu-id="b4467-160">Значение</span><span class="sxs-lookup"><span data-stu-id="b4467-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4467-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b4467-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4467-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b4467-162">1.0</span></span>|
|[<span data-ttu-id="b4467-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b4467-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b4467-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b4467-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="b4467-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b4467-165">EventType :String</span></span>

<span data-ttu-id="b4467-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="b4467-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b4467-167">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-167">Type</span></span>

*   <span data-ttu-id="b4467-168">String</span><span class="sxs-lookup"><span data-stu-id="b4467-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b4467-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b4467-169">Properties:</span></span>

| <span data-ttu-id="b4467-170">Имя</span><span class="sxs-lookup"><span data-stu-id="b4467-170">Name</span></span> | <span data-ttu-id="b4467-171">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-171">Type</span></span> | <span data-ttu-id="b4467-172">Описание</span><span class="sxs-lookup"><span data-stu-id="b4467-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="b4467-173">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-173">String</span></span> | <span data-ttu-id="b4467-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="b4467-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4467-175">Требования</span><span class="sxs-lookup"><span data-stu-id="b4467-175">Requirements</span></span>

|<span data-ttu-id="b4467-176">Требование</span><span class="sxs-lookup"><span data-stu-id="b4467-176">Requirement</span></span>| <span data-ttu-id="b4467-177">Значение</span><span class="sxs-lookup"><span data-stu-id="b4467-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4467-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b4467-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4467-179">1.5</span><span class="sxs-lookup"><span data-stu-id="b4467-179">1.5</span></span> |
|[<span data-ttu-id="b4467-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b4467-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b4467-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b4467-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b4467-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b4467-182">SourceProperty :String</span></span>

<span data-ttu-id="b4467-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b4467-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b4467-184">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-184">Type</span></span>

*   <span data-ttu-id="b4467-185">String</span><span class="sxs-lookup"><span data-stu-id="b4467-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b4467-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b4467-186">Properties:</span></span>

|<span data-ttu-id="b4467-187">Имя</span><span class="sxs-lookup"><span data-stu-id="b4467-187">Name</span></span>| <span data-ttu-id="b4467-188">Тип</span><span class="sxs-lookup"><span data-stu-id="b4467-188">Type</span></span>| <span data-ttu-id="b4467-189">Описание</span><span class="sxs-lookup"><span data-stu-id="b4467-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b4467-190">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-190">String</span></span>|<span data-ttu-id="b4467-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="b4467-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b4467-192">Строка</span><span class="sxs-lookup"><span data-stu-id="b4467-192">String</span></span>|<span data-ttu-id="b4467-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="b4467-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4467-194">Требования</span><span class="sxs-lookup"><span data-stu-id="b4467-194">Requirements</span></span>

|<span data-ttu-id="b4467-195">Требование</span><span class="sxs-lookup"><span data-stu-id="b4467-195">Requirement</span></span>| <span data-ttu-id="b4467-196">Значение</span><span class="sxs-lookup"><span data-stu-id="b4467-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4467-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b4467-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4467-198">1.0</span><span class="sxs-lookup"><span data-stu-id="b4467-198">1.0</span></span>|
|[<span data-ttu-id="b4467-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b4467-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b4467-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b4467-200">Compose or Read</span></span>|
