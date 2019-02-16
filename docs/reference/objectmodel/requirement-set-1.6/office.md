---
title: Пространство имен Office — набор обязательных элементов 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 29b3d58a4cd9dad631c2b23cabc84ade45260451
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068156"
---
# <a name="office"></a><span data-ttu-id="46bee-102">Office</span><span class="sxs-lookup"><span data-stu-id="46bee-102">Office</span></span>

<span data-ttu-id="46bee-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="46bee-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="46bee-105">Требования</span><span class="sxs-lookup"><span data-stu-id="46bee-105">Requirements</span></span>

|<span data-ttu-id="46bee-106">Требование</span><span class="sxs-lookup"><span data-stu-id="46bee-106">Requirement</span></span>| <span data-ttu-id="46bee-107">Значение</span><span class="sxs-lookup"><span data-stu-id="46bee-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="46bee-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="46bee-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46bee-109">1.0</span><span class="sxs-lookup"><span data-stu-id="46bee-109">1.0</span></span>|
|[<span data-ttu-id="46bee-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="46bee-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46bee-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="46bee-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="46bee-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="46bee-112">Members and methods</span></span>

| <span data-ttu-id="46bee-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="46bee-113">Member</span></span> | <span data-ttu-id="46bee-114">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="46bee-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="46bee-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="46bee-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="46bee-116">Member</span></span> |
| [<span data-ttu-id="46bee-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="46bee-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="46bee-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="46bee-118">Member</span></span> |
| [<span data-ttu-id="46bee-119">EventType</span><span class="sxs-lookup"><span data-stu-id="46bee-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="46bee-120">Член</span><span class="sxs-lookup"><span data-stu-id="46bee-120">Member</span></span> |
| [<span data-ttu-id="46bee-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="46bee-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="46bee-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="46bee-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="46bee-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="46bee-123">Namespaces</span></span>

<span data-ttu-id="46bee-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="46bee-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="46bee-125">[MailboxEnums.](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="46bee-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="46bee-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="46bee-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="46bee-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="46bee-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="46bee-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="46bee-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="46bee-129">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-129">Type</span></span>

*   <span data-ttu-id="46bee-130">String</span><span class="sxs-lookup"><span data-stu-id="46bee-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="46bee-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="46bee-131">Properties:</span></span>

|<span data-ttu-id="46bee-132">Имя</span><span class="sxs-lookup"><span data-stu-id="46bee-132">Name</span></span>| <span data-ttu-id="46bee-133">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-133">Type</span></span>| <span data-ttu-id="46bee-134">Описание</span><span class="sxs-lookup"><span data-stu-id="46bee-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="46bee-135">Строка</span><span class="sxs-lookup"><span data-stu-id="46bee-135">String</span></span>|<span data-ttu-id="46bee-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="46bee-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="46bee-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="46bee-137">String</span></span>|<span data-ttu-id="46bee-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="46bee-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46bee-139">Требования</span><span class="sxs-lookup"><span data-stu-id="46bee-139">Requirements</span></span>

|<span data-ttu-id="46bee-140">Требование</span><span class="sxs-lookup"><span data-stu-id="46bee-140">Requirement</span></span>| <span data-ttu-id="46bee-141">Значение</span><span class="sxs-lookup"><span data-stu-id="46bee-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="46bee-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="46bee-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46bee-143">1.0</span><span class="sxs-lookup"><span data-stu-id="46bee-143">1.0</span></span>|
|[<span data-ttu-id="46bee-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="46bee-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46bee-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="46bee-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="46bee-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="46bee-146">CoercionType :String</span></span>

<span data-ttu-id="46bee-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="46bee-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="46bee-148">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-148">Type</span></span>

*   <span data-ttu-id="46bee-149">String</span><span class="sxs-lookup"><span data-stu-id="46bee-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="46bee-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="46bee-150">Properties:</span></span>

|<span data-ttu-id="46bee-151">Имя</span><span class="sxs-lookup"><span data-stu-id="46bee-151">Name</span></span>| <span data-ttu-id="46bee-152">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-152">Type</span></span>| <span data-ttu-id="46bee-153">Описание</span><span class="sxs-lookup"><span data-stu-id="46bee-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="46bee-154">Строка</span><span class="sxs-lookup"><span data-stu-id="46bee-154">String</span></span>|<span data-ttu-id="46bee-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="46bee-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="46bee-156">String</span><span class="sxs-lookup"><span data-stu-id="46bee-156">String</span></span>|<span data-ttu-id="46bee-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="46bee-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46bee-158">Требования</span><span class="sxs-lookup"><span data-stu-id="46bee-158">Requirements</span></span>

|<span data-ttu-id="46bee-159">Требование</span><span class="sxs-lookup"><span data-stu-id="46bee-159">Requirement</span></span>| <span data-ttu-id="46bee-160">Значение</span><span class="sxs-lookup"><span data-stu-id="46bee-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="46bee-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="46bee-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46bee-162">1.0</span><span class="sxs-lookup"><span data-stu-id="46bee-162">1.0</span></span>|
|[<span data-ttu-id="46bee-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="46bee-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46bee-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="46bee-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="46bee-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="46bee-165">EventType :String</span></span>

<span data-ttu-id="46bee-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="46bee-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="46bee-167">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-167">Type</span></span>

*   <span data-ttu-id="46bee-168">String</span><span class="sxs-lookup"><span data-stu-id="46bee-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="46bee-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="46bee-169">Properties:</span></span>

| <span data-ttu-id="46bee-170">Имя</span><span class="sxs-lookup"><span data-stu-id="46bee-170">Name</span></span> | <span data-ttu-id="46bee-171">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-171">Type</span></span> | <span data-ttu-id="46bee-172">Описание</span><span class="sxs-lookup"><span data-stu-id="46bee-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="46bee-173">Строка</span><span class="sxs-lookup"><span data-stu-id="46bee-173">String</span></span> | <span data-ttu-id="46bee-174">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="46bee-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="46bee-175">Требования</span><span class="sxs-lookup"><span data-stu-id="46bee-175">Requirements</span></span>

|<span data-ttu-id="46bee-176">Требование</span><span class="sxs-lookup"><span data-stu-id="46bee-176">Requirement</span></span>| <span data-ttu-id="46bee-177">Значение</span><span class="sxs-lookup"><span data-stu-id="46bee-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="46bee-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="46bee-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46bee-179">1.5</span><span class="sxs-lookup"><span data-stu-id="46bee-179">1.5</span></span> |
|[<span data-ttu-id="46bee-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="46bee-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46bee-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="46bee-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="46bee-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="46bee-182">SourceProperty :String</span></span>

<span data-ttu-id="46bee-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="46bee-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="46bee-184">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-184">Type</span></span>

*   <span data-ttu-id="46bee-185">String</span><span class="sxs-lookup"><span data-stu-id="46bee-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="46bee-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="46bee-186">Properties:</span></span>

|<span data-ttu-id="46bee-187">Имя</span><span class="sxs-lookup"><span data-stu-id="46bee-187">Name</span></span>| <span data-ttu-id="46bee-188">Тип</span><span class="sxs-lookup"><span data-stu-id="46bee-188">Type</span></span>| <span data-ttu-id="46bee-189">Описание</span><span class="sxs-lookup"><span data-stu-id="46bee-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="46bee-190">Строка</span><span class="sxs-lookup"><span data-stu-id="46bee-190">String</span></span>|<span data-ttu-id="46bee-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="46bee-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="46bee-192">String</span><span class="sxs-lookup"><span data-stu-id="46bee-192">String</span></span>|<span data-ttu-id="46bee-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="46bee-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="46bee-194">Требования</span><span class="sxs-lookup"><span data-stu-id="46bee-194">Requirements</span></span>

|<span data-ttu-id="46bee-195">Требование</span><span class="sxs-lookup"><span data-stu-id="46bee-195">Requirement</span></span>| <span data-ttu-id="46bee-196">Значение</span><span class="sxs-lookup"><span data-stu-id="46bee-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="46bee-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="46bee-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="46bee-198">1.0</span><span class="sxs-lookup"><span data-stu-id="46bee-198">1.0</span></span>|
|[<span data-ttu-id="46bee-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="46bee-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="46bee-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="46bee-200">Compose or Read</span></span>|
