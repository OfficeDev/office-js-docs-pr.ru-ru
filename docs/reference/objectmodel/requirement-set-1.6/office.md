---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: ae764e8cda2b3f14e33b883d054379db7b37a687
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696003"
---
# <a name="office"></a><span data-ttu-id="f0710-102">Office</span><span class="sxs-lookup"><span data-stu-id="f0710-102">Office</span></span>

<span data-ttu-id="f0710-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f0710-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f0710-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f0710-105">Requirements</span></span>

|<span data-ttu-id="f0710-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f0710-106">Requirement</span></span>| <span data-ttu-id="f0710-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f0710-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0710-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f0710-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0710-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f0710-109">1.0</span></span>|
|[<span data-ttu-id="f0710-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f0710-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0710-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f0710-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f0710-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="f0710-112">Members and methods</span></span>

| <span data-ttu-id="f0710-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="f0710-113">Member</span></span> | <span data-ttu-id="f0710-114">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f0710-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f0710-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f0710-116">Member</span><span class="sxs-lookup"><span data-stu-id="f0710-116">Member</span></span> |
| [<span data-ttu-id="f0710-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f0710-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f0710-118">Member</span><span class="sxs-lookup"><span data-stu-id="f0710-118">Member</span></span> |
| [<span data-ttu-id="f0710-119">EventType</span><span class="sxs-lookup"><span data-stu-id="f0710-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f0710-120">Member</span><span class="sxs-lookup"><span data-stu-id="f0710-120">Member</span></span> |
| [<span data-ttu-id="f0710-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f0710-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f0710-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="f0710-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f0710-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f0710-123">Namespaces</span></span>

<span data-ttu-id="f0710-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f0710-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f0710-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f0710-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="f0710-126">Members</span><span class="sxs-lookup"><span data-stu-id="f0710-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f0710-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="f0710-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="f0710-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f0710-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f0710-129">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-129">Type</span></span>

*   <span data-ttu-id="f0710-130">String</span><span class="sxs-lookup"><span data-stu-id="f0710-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0710-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f0710-131">Properties:</span></span>

|<span data-ttu-id="f0710-132">Имя</span><span class="sxs-lookup"><span data-stu-id="f0710-132">Name</span></span>| <span data-ttu-id="f0710-133">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-133">Type</span></span>| <span data-ttu-id="f0710-134">Описание</span><span class="sxs-lookup"><span data-stu-id="f0710-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f0710-135">String</span><span class="sxs-lookup"><span data-stu-id="f0710-135">String</span></span>|<span data-ttu-id="f0710-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f0710-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f0710-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="f0710-137">String</span></span>|<span data-ttu-id="f0710-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="f0710-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0710-139">Требования</span><span class="sxs-lookup"><span data-stu-id="f0710-139">Requirements</span></span>

|<span data-ttu-id="f0710-140">Требование</span><span class="sxs-lookup"><span data-stu-id="f0710-140">Requirement</span></span>| <span data-ttu-id="f0710-141">Значение</span><span class="sxs-lookup"><span data-stu-id="f0710-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0710-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f0710-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0710-143">1.0</span><span class="sxs-lookup"><span data-stu-id="f0710-143">1.0</span></span>|
|[<span data-ttu-id="f0710-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f0710-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0710-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f0710-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f0710-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="f0710-146">CoercionType: String</span></span>

<span data-ttu-id="f0710-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f0710-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0710-148">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-148">Type</span></span>

*   <span data-ttu-id="f0710-149">String</span><span class="sxs-lookup"><span data-stu-id="f0710-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0710-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f0710-150">Properties:</span></span>

|<span data-ttu-id="f0710-151">Имя</span><span class="sxs-lookup"><span data-stu-id="f0710-151">Name</span></span>| <span data-ttu-id="f0710-152">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-152">Type</span></span>| <span data-ttu-id="f0710-153">Описание</span><span class="sxs-lookup"><span data-stu-id="f0710-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f0710-154">String</span><span class="sxs-lookup"><span data-stu-id="f0710-154">String</span></span>|<span data-ttu-id="f0710-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f0710-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f0710-156">String.</span><span class="sxs-lookup"><span data-stu-id="f0710-156">String</span></span>|<span data-ttu-id="f0710-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f0710-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0710-158">Требования</span><span class="sxs-lookup"><span data-stu-id="f0710-158">Requirements</span></span>

|<span data-ttu-id="f0710-159">Требование</span><span class="sxs-lookup"><span data-stu-id="f0710-159">Requirement</span></span>| <span data-ttu-id="f0710-160">Значение</span><span class="sxs-lookup"><span data-stu-id="f0710-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0710-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f0710-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0710-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f0710-162">1.0</span></span>|
|[<span data-ttu-id="f0710-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f0710-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0710-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f0710-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f0710-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="f0710-165">EventType: String</span></span>

<span data-ttu-id="f0710-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="f0710-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f0710-167">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-167">Type</span></span>

*   <span data-ttu-id="f0710-168">String</span><span class="sxs-lookup"><span data-stu-id="f0710-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0710-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f0710-169">Properties:</span></span>

| <span data-ttu-id="f0710-170">Имя</span><span class="sxs-lookup"><span data-stu-id="f0710-170">Name</span></span> | <span data-ttu-id="f0710-171">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-171">Type</span></span> | <span data-ttu-id="f0710-172">Описание</span><span class="sxs-lookup"><span data-stu-id="f0710-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="f0710-173">String</span><span class="sxs-lookup"><span data-stu-id="f0710-173">String</span></span> | <span data-ttu-id="f0710-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="f0710-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f0710-175">Требования</span><span class="sxs-lookup"><span data-stu-id="f0710-175">Requirements</span></span>

|<span data-ttu-id="f0710-176">Требование</span><span class="sxs-lookup"><span data-stu-id="f0710-176">Requirement</span></span>| <span data-ttu-id="f0710-177">Значение</span><span class="sxs-lookup"><span data-stu-id="f0710-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0710-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f0710-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0710-179">1.5</span><span class="sxs-lookup"><span data-stu-id="f0710-179">1.5</span></span> |
|[<span data-ttu-id="f0710-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f0710-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0710-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f0710-181">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f0710-182">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="f0710-182">SourceProperty: String</span></span>

<span data-ttu-id="f0710-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f0710-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f0710-184">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-184">Type</span></span>

*   <span data-ttu-id="f0710-185">String</span><span class="sxs-lookup"><span data-stu-id="f0710-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f0710-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f0710-186">Properties:</span></span>

|<span data-ttu-id="f0710-187">Имя</span><span class="sxs-lookup"><span data-stu-id="f0710-187">Name</span></span>| <span data-ttu-id="f0710-188">Тип</span><span class="sxs-lookup"><span data-stu-id="f0710-188">Type</span></span>| <span data-ttu-id="f0710-189">Описание</span><span class="sxs-lookup"><span data-stu-id="f0710-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f0710-190">String</span><span class="sxs-lookup"><span data-stu-id="f0710-190">String</span></span>|<span data-ttu-id="f0710-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f0710-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f0710-192">String.</span><span class="sxs-lookup"><span data-stu-id="f0710-192">String</span></span>|<span data-ttu-id="f0710-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f0710-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f0710-194">Требования</span><span class="sxs-lookup"><span data-stu-id="f0710-194">Requirements</span></span>

|<span data-ttu-id="f0710-195">Требование</span><span class="sxs-lookup"><span data-stu-id="f0710-195">Requirement</span></span>| <span data-ttu-id="f0710-196">Значение</span><span class="sxs-lookup"><span data-stu-id="f0710-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="f0710-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f0710-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f0710-198">1.0</span><span class="sxs-lookup"><span data-stu-id="f0710-198">1.0</span></span>|
|[<span data-ttu-id="f0710-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f0710-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f0710-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f0710-200">Compose or Read</span></span>|
