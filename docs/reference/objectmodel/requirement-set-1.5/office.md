---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 2236dae5421090a571c8cc658cb6f67f2a08d54a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696080"
---
# <a name="office"></a><span data-ttu-id="f4c7a-102">Office</span><span class="sxs-lookup"><span data-stu-id="f4c7a-102">Office</span></span>

<span data-ttu-id="f4c7a-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f4c7a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4c7a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f4c7a-105">Requirements</span></span>

|<span data-ttu-id="f4c7a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f4c7a-106">Requirement</span></span>| <span data-ttu-id="f4c7a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4c7a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4c7a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4c7a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f4c7a-109">1.0</span></span>|
|[<span data-ttu-id="f4c7a-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4c7a-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4c7a-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f4c7a-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="f4c7a-112">Members and methods</span></span>

| <span data-ttu-id="f4c7a-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="f4c7a-113">Member</span></span> | <span data-ttu-id="f4c7a-114">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f4c7a-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f4c7a-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f4c7a-116">Member</span><span class="sxs-lookup"><span data-stu-id="f4c7a-116">Member</span></span> |
| [<span data-ttu-id="f4c7a-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f4c7a-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f4c7a-118">Member</span><span class="sxs-lookup"><span data-stu-id="f4c7a-118">Member</span></span> |
| [<span data-ttu-id="f4c7a-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f4c7a-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f4c7a-120">Member</span><span class="sxs-lookup"><span data-stu-id="f4c7a-120">Member</span></span> |
| [<span data-ttu-id="f4c7a-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f4c7a-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f4c7a-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="f4c7a-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f4c7a-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f4c7a-123">Namespaces</span></span>

<span data-ttu-id="f4c7a-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f4c7a-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="f4c7a-126">Members</span><span class="sxs-lookup"><span data-stu-id="f4c7a-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f4c7a-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="f4c7a-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="f4c7a-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f4c7a-129">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-129">Type</span></span>

*   <span data-ttu-id="f4c7a-130">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4c7a-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f4c7a-131">Properties:</span></span>

|<span data-ttu-id="f4c7a-132">Имя</span><span class="sxs-lookup"><span data-stu-id="f4c7a-132">Name</span></span>| <span data-ttu-id="f4c7a-133">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-133">Type</span></span>| <span data-ttu-id="f4c7a-134">Описание</span><span class="sxs-lookup"><span data-stu-id="f4c7a-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f4c7a-135">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-135">String</span></span>|<span data-ttu-id="f4c7a-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f4c7a-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="f4c7a-137">String</span></span>|<span data-ttu-id="f4c7a-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4c7a-139">Требования</span><span class="sxs-lookup"><span data-stu-id="f4c7a-139">Requirements</span></span>

|<span data-ttu-id="f4c7a-140">Требование</span><span class="sxs-lookup"><span data-stu-id="f4c7a-140">Requirement</span></span>| <span data-ttu-id="f4c7a-141">Значение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4c7a-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4c7a-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4c7a-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f4c7a-143">1.0</span></span>|
|[<span data-ttu-id="f4c7a-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4c7a-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4c7a-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f4c7a-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="f4c7a-146">CoercionType: String</span></span>

<span data-ttu-id="f4c7a-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f4c7a-148">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-148">Type</span></span>

*   <span data-ttu-id="f4c7a-149">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4c7a-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f4c7a-150">Properties:</span></span>

|<span data-ttu-id="f4c7a-151">Имя</span><span class="sxs-lookup"><span data-stu-id="f4c7a-151">Name</span></span>| <span data-ttu-id="f4c7a-152">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-152">Type</span></span>| <span data-ttu-id="f4c7a-153">Описание</span><span class="sxs-lookup"><span data-stu-id="f4c7a-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f4c7a-154">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-154">String</span></span>|<span data-ttu-id="f4c7a-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f4c7a-156">String.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-156">String</span></span>|<span data-ttu-id="f4c7a-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4c7a-158">Требования</span><span class="sxs-lookup"><span data-stu-id="f4c7a-158">Requirements</span></span>

|<span data-ttu-id="f4c7a-159">Требование</span><span class="sxs-lookup"><span data-stu-id="f4c7a-159">Requirement</span></span>| <span data-ttu-id="f4c7a-160">Значение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4c7a-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4c7a-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4c7a-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f4c7a-162">1.0</span></span>|
|[<span data-ttu-id="f4c7a-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4c7a-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4c7a-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f4c7a-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="f4c7a-165">EventType: String</span></span>

<span data-ttu-id="f4c7a-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f4c7a-167">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-167">Type</span></span>

*   <span data-ttu-id="f4c7a-168">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4c7a-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f4c7a-169">Properties:</span></span>

| <span data-ttu-id="f4c7a-170">Имя</span><span class="sxs-lookup"><span data-stu-id="f4c7a-170">Name</span></span> | <span data-ttu-id="f4c7a-171">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-171">Type</span></span> | <span data-ttu-id="f4c7a-172">Описание</span><span class="sxs-lookup"><span data-stu-id="f4c7a-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="f4c7a-173">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-173">String</span></span> | <span data-ttu-id="f4c7a-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f4c7a-175">Требования</span><span class="sxs-lookup"><span data-stu-id="f4c7a-175">Requirements</span></span>

|<span data-ttu-id="f4c7a-176">Требование</span><span class="sxs-lookup"><span data-stu-id="f4c7a-176">Requirement</span></span>| <span data-ttu-id="f4c7a-177">Значение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4c7a-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f4c7a-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4c7a-179">1.5</span><span class="sxs-lookup"><span data-stu-id="f4c7a-179">1.5</span></span> |
|[<span data-ttu-id="f4c7a-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4c7a-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4c7a-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f4c7a-182">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="f4c7a-182">SourceProperty: String</span></span>

<span data-ttu-id="f4c7a-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f4c7a-184">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-184">Type</span></span>

*   <span data-ttu-id="f4c7a-185">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f4c7a-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f4c7a-186">Properties:</span></span>

|<span data-ttu-id="f4c7a-187">Имя</span><span class="sxs-lookup"><span data-stu-id="f4c7a-187">Name</span></span>| <span data-ttu-id="f4c7a-188">Тип</span><span class="sxs-lookup"><span data-stu-id="f4c7a-188">Type</span></span>| <span data-ttu-id="f4c7a-189">Описание</span><span class="sxs-lookup"><span data-stu-id="f4c7a-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f4c7a-190">String</span><span class="sxs-lookup"><span data-stu-id="f4c7a-190">String</span></span>|<span data-ttu-id="f4c7a-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f4c7a-192">String.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-192">String</span></span>|<span data-ttu-id="f4c7a-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f4c7a-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f4c7a-194">Требования</span><span class="sxs-lookup"><span data-stu-id="f4c7a-194">Requirements</span></span>

|<span data-ttu-id="f4c7a-195">Требование</span><span class="sxs-lookup"><span data-stu-id="f4c7a-195">Requirement</span></span>| <span data-ttu-id="f4c7a-196">Значение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4c7a-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4c7a-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4c7a-198">1.0</span><span class="sxs-lookup"><span data-stu-id="f4c7a-198">1.0</span></span>|
|[<span data-ttu-id="f4c7a-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4c7a-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4c7a-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4c7a-200">Compose or Read</span></span>|
