---
title: Пространство имен Office — набор обязательных элементов 1.5
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 46b21df77456d2392fbc543e45513246a4ad9a10
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433651"
---
# <a name="office"></a><span data-ttu-id="b3bad-102">Office</span><span class="sxs-lookup"><span data-stu-id="b3bad-102">Office</span></span>

<span data-ttu-id="b3bad-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b3bad-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b3bad-105">Требования</span><span class="sxs-lookup"><span data-stu-id="b3bad-105">Requirements</span></span>

|<span data-ttu-id="b3bad-106">Требование</span><span class="sxs-lookup"><span data-stu-id="b3bad-106">Requirement</span></span>| <span data-ttu-id="b3bad-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b3bad-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b3bad-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b3bad-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b3bad-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b3bad-109">1.0</span></span>|
|[<span data-ttu-id="b3bad-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b3bad-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b3bad-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b3bad-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b3bad-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="b3bad-112">Members and methods</span></span>

| <span data-ttu-id="b3bad-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="b3bad-113">Member</span></span> | <span data-ttu-id="b3bad-114">Тип</span><span class="sxs-lookup"><span data-stu-id="b3bad-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b3bad-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b3bad-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b3bad-116">Член</span><span class="sxs-lookup"><span data-stu-id="b3bad-116">Member</span></span> |
| [<span data-ttu-id="b3bad-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b3bad-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b3bad-118">Член</span><span class="sxs-lookup"><span data-stu-id="b3bad-118">Member</span></span> |
| [<span data-ttu-id="b3bad-119">EventType</span><span class="sxs-lookup"><span data-stu-id="b3bad-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b3bad-120">Член</span><span class="sxs-lookup"><span data-stu-id="b3bad-120">Member</span></span> |
| [<span data-ttu-id="b3bad-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b3bad-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b3bad-122">Член</span><span class="sxs-lookup"><span data-stu-id="b3bad-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b3bad-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="b3bad-123">Namespaces</span></span>

<span data-ttu-id="b3bad-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="b3bad-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b3bad-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b3bad-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b3bad-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="b3bad-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b3bad-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b3bad-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="b3bad-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="b3bad-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b3bad-129">Тип:</span><span class="sxs-lookup"><span data-stu-id="b3bad-129">Type:</span></span>

*   <span data-ttu-id="b3bad-130">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b3bad-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b3bad-131">Properties:</span></span>

|<span data-ttu-id="b3bad-132">Имя</span><span class="sxs-lookup"><span data-stu-id="b3bad-132">Name</span></span>| <span data-ttu-id="b3bad-133">Тип</span><span class="sxs-lookup"><span data-stu-id="b3bad-133">Type</span></span>| <span data-ttu-id="b3bad-134">Описание</span><span class="sxs-lookup"><span data-stu-id="b3bad-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b3bad-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="b3bad-135">String</span></span>|<span data-ttu-id="b3bad-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="b3bad-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b3bad-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="b3bad-137">String</span></span>|<span data-ttu-id="b3bad-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="b3bad-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b3bad-139">Требования</span><span class="sxs-lookup"><span data-stu-id="b3bad-139">Requirements</span></span>

|<span data-ttu-id="b3bad-140">Требование</span><span class="sxs-lookup"><span data-stu-id="b3bad-140">Requirement</span></span>| <span data-ttu-id="b3bad-141">Значение</span><span class="sxs-lookup"><span data-stu-id="b3bad-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="b3bad-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b3bad-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b3bad-143">1.0</span><span class="sxs-lookup"><span data-stu-id="b3bad-143">1.0</span></span>|
|[<span data-ttu-id="b3bad-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b3bad-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b3bad-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b3bad-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="b3bad-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b3bad-146">CoercionType :String</span></span>

<span data-ttu-id="b3bad-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b3bad-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b3bad-148">Тип:</span><span class="sxs-lookup"><span data-stu-id="b3bad-148">Type:</span></span>

*   <span data-ttu-id="b3bad-149">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b3bad-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b3bad-150">Properties:</span></span>

|<span data-ttu-id="b3bad-151">Имя</span><span class="sxs-lookup"><span data-stu-id="b3bad-151">Name</span></span>| <span data-ttu-id="b3bad-152">Тип</span><span class="sxs-lookup"><span data-stu-id="b3bad-152">Type</span></span>| <span data-ttu-id="b3bad-153">Описание</span><span class="sxs-lookup"><span data-stu-id="b3bad-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b3bad-154">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-154">String</span></span>|<span data-ttu-id="b3bad-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="b3bad-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b3bad-156">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-156">String</span></span>|<span data-ttu-id="b3bad-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="b3bad-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b3bad-158">Требования</span><span class="sxs-lookup"><span data-stu-id="b3bad-158">Requirements</span></span>

|<span data-ttu-id="b3bad-159">Требование</span><span class="sxs-lookup"><span data-stu-id="b3bad-159">Requirement</span></span>| <span data-ttu-id="b3bad-160">Значение</span><span class="sxs-lookup"><span data-stu-id="b3bad-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b3bad-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b3bad-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b3bad-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b3bad-162">1.0</span></span>|
|[<span data-ttu-id="b3bad-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b3bad-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b3bad-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b3bad-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="b3bad-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b3bad-165">EventType :String</span></span>

<span data-ttu-id="b3bad-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="b3bad-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b3bad-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="b3bad-167">Type:</span></span>

*   <span data-ttu-id="b3bad-168">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b3bad-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b3bad-169">Properties:</span></span>

| <span data-ttu-id="b3bad-170">Имя</span><span class="sxs-lookup"><span data-stu-id="b3bad-170">Name</span></span> | <span data-ttu-id="b3bad-171">Тип</span><span class="sxs-lookup"><span data-stu-id="b3bad-171">Type</span></span> | <span data-ttu-id="b3bad-172">Описание</span><span class="sxs-lookup"><span data-stu-id="b3bad-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="b3bad-173">Строка</span><span class="sxs-lookup"><span data-stu-id="b3bad-173">String</span></span> | <span data-ttu-id="b3bad-174">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="b3bad-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b3bad-175">Требования</span><span class="sxs-lookup"><span data-stu-id="b3bad-175">Requirements</span></span>

|<span data-ttu-id="b3bad-176">Требование</span><span class="sxs-lookup"><span data-stu-id="b3bad-176">Requirement</span></span>| <span data-ttu-id="b3bad-177">Значение</span><span class="sxs-lookup"><span data-stu-id="b3bad-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="b3bad-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b3bad-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b3bad-179">1.5</span><span class="sxs-lookup"><span data-stu-id="b3bad-179">1.5</span></span> |
|[<span data-ttu-id="b3bad-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b3bad-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b3bad-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b3bad-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b3bad-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b3bad-182">SourceProperty :String</span></span>

<span data-ttu-id="b3bad-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b3bad-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b3bad-184">Тип:</span><span class="sxs-lookup"><span data-stu-id="b3bad-184">Type:</span></span>

*   <span data-ttu-id="b3bad-185">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b3bad-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b3bad-186">Properties:</span></span>

|<span data-ttu-id="b3bad-187">Имя</span><span class="sxs-lookup"><span data-stu-id="b3bad-187">Name</span></span>| <span data-ttu-id="b3bad-188">Тип</span><span class="sxs-lookup"><span data-stu-id="b3bad-188">Type</span></span>| <span data-ttu-id="b3bad-189">Описание</span><span class="sxs-lookup"><span data-stu-id="b3bad-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b3bad-190">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-190">String</span></span>|<span data-ttu-id="b3bad-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="b3bad-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b3bad-192">String</span><span class="sxs-lookup"><span data-stu-id="b3bad-192">String</span></span>|<span data-ttu-id="b3bad-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="b3bad-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b3bad-194">Требования</span><span class="sxs-lookup"><span data-stu-id="b3bad-194">Requirements</span></span>

|<span data-ttu-id="b3bad-195">Требование</span><span class="sxs-lookup"><span data-stu-id="b3bad-195">Requirement</span></span>| <span data-ttu-id="b3bad-196">Значение</span><span class="sxs-lookup"><span data-stu-id="b3bad-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="b3bad-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b3bad-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b3bad-198">1.0</span><span class="sxs-lookup"><span data-stu-id="b3bad-198">1.0</span></span>|
|[<span data-ttu-id="b3bad-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b3bad-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b3bad-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b3bad-200">Compose or read</span></span>|