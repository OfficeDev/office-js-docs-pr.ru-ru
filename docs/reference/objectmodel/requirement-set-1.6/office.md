---
title: Пространство имен Office — набор обязательных элементов 1.6
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: bf6304515c511eea580a3f37d898b7e80adffaee
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457896"
---
# <a name="office"></a><span data-ttu-id="fcc45-102">Office</span><span class="sxs-lookup"><span data-stu-id="fcc45-102">Office</span></span>

<span data-ttu-id="fcc45-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="fcc45-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fcc45-105">Требования</span><span class="sxs-lookup"><span data-stu-id="fcc45-105">Requirements</span></span>

|<span data-ttu-id="fcc45-106">Требование</span><span class="sxs-lookup"><span data-stu-id="fcc45-106">Requirement</span></span>| <span data-ttu-id="fcc45-107">Значение</span><span class="sxs-lookup"><span data-stu-id="fcc45-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fcc45-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fcc45-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fcc45-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fcc45-109">1.0</span></span>|
|[<span data-ttu-id="fcc45-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fcc45-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fcc45-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fcc45-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fcc45-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="fcc45-112">Members and methods</span></span>

| <span data-ttu-id="fcc45-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="fcc45-113">Member</span></span> | <span data-ttu-id="fcc45-114">Тип</span><span class="sxs-lookup"><span data-stu-id="fcc45-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fcc45-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="fcc45-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="fcc45-116">Член</span><span class="sxs-lookup"><span data-stu-id="fcc45-116">Member</span></span> |
| [<span data-ttu-id="fcc45-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="fcc45-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="fcc45-118">Член</span><span class="sxs-lookup"><span data-stu-id="fcc45-118">Member</span></span> |
| [<span data-ttu-id="fcc45-119">EventType</span><span class="sxs-lookup"><span data-stu-id="fcc45-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="fcc45-120">Член</span><span class="sxs-lookup"><span data-stu-id="fcc45-120">Member</span></span> |
| [<span data-ttu-id="fcc45-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="fcc45-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="fcc45-122">Член</span><span class="sxs-lookup"><span data-stu-id="fcc45-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fcc45-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="fcc45-123">Namespaces</span></span>

<span data-ttu-id="fcc45-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="fcc45-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="fcc45-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="fcc45-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="fcc45-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="fcc45-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="fcc45-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="fcc45-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="fcc45-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="fcc45-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="fcc45-129">Тип:</span><span class="sxs-lookup"><span data-stu-id="fcc45-129">Type:</span></span>

*   <span data-ttu-id="fcc45-130">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fcc45-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="fcc45-131">Properties:</span></span>

|<span data-ttu-id="fcc45-132">Имя</span><span class="sxs-lookup"><span data-stu-id="fcc45-132">Name</span></span>| <span data-ttu-id="fcc45-133">Тип</span><span class="sxs-lookup"><span data-stu-id="fcc45-133">Type</span></span>| <span data-ttu-id="fcc45-134">Описание</span><span class="sxs-lookup"><span data-stu-id="fcc45-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="fcc45-135">Для указания</span><span class="sxs-lookup"><span data-stu-id="fcc45-135">String</span></span>|<span data-ttu-id="fcc45-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="fcc45-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="fcc45-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="fcc45-137">String</span></span>|<span data-ttu-id="fcc45-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="fcc45-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fcc45-139">Требования</span><span class="sxs-lookup"><span data-stu-id="fcc45-139">Requirements</span></span>

|<span data-ttu-id="fcc45-140">Требование</span><span class="sxs-lookup"><span data-stu-id="fcc45-140">Requirement</span></span>| <span data-ttu-id="fcc45-141">Значение</span><span class="sxs-lookup"><span data-stu-id="fcc45-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="fcc45-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fcc45-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fcc45-143">1.0</span><span class="sxs-lookup"><span data-stu-id="fcc45-143">1.0</span></span>|
|[<span data-ttu-id="fcc45-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fcc45-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fcc45-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fcc45-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="fcc45-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="fcc45-146">CoercionType :String</span></span>

<span data-ttu-id="fcc45-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="fcc45-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fcc45-148">Тип:</span><span class="sxs-lookup"><span data-stu-id="fcc45-148">Type:</span></span>

*   <span data-ttu-id="fcc45-149">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fcc45-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="fcc45-150">Properties:</span></span>

|<span data-ttu-id="fcc45-151">Имя</span><span class="sxs-lookup"><span data-stu-id="fcc45-151">Name</span></span>| <span data-ttu-id="fcc45-152">Тип</span><span class="sxs-lookup"><span data-stu-id="fcc45-152">Type</span></span>| <span data-ttu-id="fcc45-153">Описание</span><span class="sxs-lookup"><span data-stu-id="fcc45-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="fcc45-154">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-154">String</span></span>|<span data-ttu-id="fcc45-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="fcc45-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="fcc45-156">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-156">String</span></span>|<span data-ttu-id="fcc45-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="fcc45-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fcc45-158">Требования</span><span class="sxs-lookup"><span data-stu-id="fcc45-158">Requirements</span></span>

|<span data-ttu-id="fcc45-159">Требование</span><span class="sxs-lookup"><span data-stu-id="fcc45-159">Requirement</span></span>| <span data-ttu-id="fcc45-160">Значение</span><span class="sxs-lookup"><span data-stu-id="fcc45-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="fcc45-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fcc45-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fcc45-162">1.0</span><span class="sxs-lookup"><span data-stu-id="fcc45-162">1.0</span></span>|
|[<span data-ttu-id="fcc45-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fcc45-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fcc45-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fcc45-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="fcc45-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="fcc45-165">EventType :String</span></span>

<span data-ttu-id="fcc45-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="fcc45-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="fcc45-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="fcc45-167">Type:</span></span>

*   <span data-ttu-id="fcc45-168">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fcc45-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="fcc45-169">Properties:</span></span>

| <span data-ttu-id="fcc45-170">Имя</span><span class="sxs-lookup"><span data-stu-id="fcc45-170">Name</span></span> | <span data-ttu-id="fcc45-171">Тип</span><span class="sxs-lookup"><span data-stu-id="fcc45-171">Type</span></span> | <span data-ttu-id="fcc45-172">Описание</span><span class="sxs-lookup"><span data-stu-id="fcc45-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="fcc45-173">Строка</span><span class="sxs-lookup"><span data-stu-id="fcc45-173">String</span></span> | <span data-ttu-id="fcc45-174">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="fcc45-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fcc45-175">Требования</span><span class="sxs-lookup"><span data-stu-id="fcc45-175">Requirements</span></span>

|<span data-ttu-id="fcc45-176">Требование</span><span class="sxs-lookup"><span data-stu-id="fcc45-176">Requirement</span></span>| <span data-ttu-id="fcc45-177">Значение</span><span class="sxs-lookup"><span data-stu-id="fcc45-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="fcc45-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="fcc45-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fcc45-179">1.5</span><span class="sxs-lookup"><span data-stu-id="fcc45-179">1.5</span></span> |
|[<span data-ttu-id="fcc45-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fcc45-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fcc45-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fcc45-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="fcc45-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="fcc45-182">SourceProperty :String</span></span>

<span data-ttu-id="fcc45-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="fcc45-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="fcc45-184">Тип:</span><span class="sxs-lookup"><span data-stu-id="fcc45-184">Type:</span></span>

*   <span data-ttu-id="fcc45-185">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="fcc45-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="fcc45-186">Properties:</span></span>

|<span data-ttu-id="fcc45-187">Имя</span><span class="sxs-lookup"><span data-stu-id="fcc45-187">Name</span></span>| <span data-ttu-id="fcc45-188">Тип</span><span class="sxs-lookup"><span data-stu-id="fcc45-188">Type</span></span>| <span data-ttu-id="fcc45-189">Описание</span><span class="sxs-lookup"><span data-stu-id="fcc45-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="fcc45-190">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-190">String</span></span>|<span data-ttu-id="fcc45-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="fcc45-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="fcc45-192">String</span><span class="sxs-lookup"><span data-stu-id="fcc45-192">String</span></span>|<span data-ttu-id="fcc45-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="fcc45-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fcc45-194">Требования</span><span class="sxs-lookup"><span data-stu-id="fcc45-194">Requirements</span></span>

|<span data-ttu-id="fcc45-195">Требование</span><span class="sxs-lookup"><span data-stu-id="fcc45-195">Requirement</span></span>| <span data-ttu-id="fcc45-196">Значение</span><span class="sxs-lookup"><span data-stu-id="fcc45-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="fcc45-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="fcc45-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fcc45-198">1.0</span><span class="sxs-lookup"><span data-stu-id="fcc45-198">1.0</span></span>|
|[<span data-ttu-id="fcc45-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fcc45-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fcc45-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fcc45-200">Compose or read</span></span>|