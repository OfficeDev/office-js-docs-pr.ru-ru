---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 84e8fa49e1d4dce4239525badafaa051325bb3ec
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395640"
---
# <a name="office"></a><span data-ttu-id="4a671-102">Office</span><span class="sxs-lookup"><span data-stu-id="4a671-102">Office</span></span>

<span data-ttu-id="4a671-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4a671-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a671-105">Требования</span><span class="sxs-lookup"><span data-stu-id="4a671-105">Requirements</span></span>

|<span data-ttu-id="4a671-106">Требование</span><span class="sxs-lookup"><span data-stu-id="4a671-106">Requirement</span></span>| <span data-ttu-id="4a671-107">Значение</span><span class="sxs-lookup"><span data-stu-id="4a671-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a671-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a671-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a671-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4a671-109">1.0</span></span>|
|[<span data-ttu-id="4a671-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a671-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a671-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a671-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4a671-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="4a671-112">Members and methods</span></span>

| <span data-ttu-id="4a671-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="4a671-113">Member</span></span> | <span data-ttu-id="4a671-114">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4a671-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4a671-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4a671-116">Member</span><span class="sxs-lookup"><span data-stu-id="4a671-116">Member</span></span> |
| [<span data-ttu-id="4a671-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4a671-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4a671-118">Member</span><span class="sxs-lookup"><span data-stu-id="4a671-118">Member</span></span> |
| [<span data-ttu-id="4a671-119">EventType</span><span class="sxs-lookup"><span data-stu-id="4a671-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4a671-120">Member</span><span class="sxs-lookup"><span data-stu-id="4a671-120">Member</span></span> |
| [<span data-ttu-id="4a671-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4a671-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4a671-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="4a671-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4a671-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="4a671-123">Namespaces</span></span>

<span data-ttu-id="4a671-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="4a671-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4a671-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="4a671-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.6): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="4a671-126">Members</span><span class="sxs-lookup"><span data-stu-id="4a671-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4a671-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="4a671-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="4a671-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="4a671-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4a671-129">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-129">Type</span></span>

*   <span data-ttu-id="4a671-130">String</span><span class="sxs-lookup"><span data-stu-id="4a671-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4a671-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4a671-131">Properties:</span></span>

|<span data-ttu-id="4a671-132">Имя</span><span class="sxs-lookup"><span data-stu-id="4a671-132">Name</span></span>| <span data-ttu-id="4a671-133">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-133">Type</span></span>| <span data-ttu-id="4a671-134">Описание</span><span class="sxs-lookup"><span data-stu-id="4a671-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4a671-135">String</span><span class="sxs-lookup"><span data-stu-id="4a671-135">String</span></span>|<span data-ttu-id="4a671-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="4a671-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4a671-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="4a671-137">String</span></span>|<span data-ttu-id="4a671-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="4a671-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a671-139">Требования</span><span class="sxs-lookup"><span data-stu-id="4a671-139">Requirements</span></span>

|<span data-ttu-id="4a671-140">Требование</span><span class="sxs-lookup"><span data-stu-id="4a671-140">Requirement</span></span>| <span data-ttu-id="4a671-141">Значение</span><span class="sxs-lookup"><span data-stu-id="4a671-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a671-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a671-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a671-143">1.0</span><span class="sxs-lookup"><span data-stu-id="4a671-143">1.0</span></span>|
|[<span data-ttu-id="4a671-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a671-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a671-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a671-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="4a671-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="4a671-146">CoercionType: String</span></span>

<span data-ttu-id="4a671-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="4a671-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4a671-148">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-148">Type</span></span>

*   <span data-ttu-id="4a671-149">String</span><span class="sxs-lookup"><span data-stu-id="4a671-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4a671-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4a671-150">Properties:</span></span>

|<span data-ttu-id="4a671-151">Имя</span><span class="sxs-lookup"><span data-stu-id="4a671-151">Name</span></span>| <span data-ttu-id="4a671-152">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-152">Type</span></span>| <span data-ttu-id="4a671-153">Описание</span><span class="sxs-lookup"><span data-stu-id="4a671-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4a671-154">String</span><span class="sxs-lookup"><span data-stu-id="4a671-154">String</span></span>|<span data-ttu-id="4a671-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="4a671-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4a671-156">String</span><span class="sxs-lookup"><span data-stu-id="4a671-156">String</span></span>|<span data-ttu-id="4a671-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="4a671-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a671-158">Требования</span><span class="sxs-lookup"><span data-stu-id="4a671-158">Requirements</span></span>

|<span data-ttu-id="4a671-159">Требование</span><span class="sxs-lookup"><span data-stu-id="4a671-159">Requirement</span></span>| <span data-ttu-id="4a671-160">Значение</span><span class="sxs-lookup"><span data-stu-id="4a671-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a671-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a671-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a671-162">1.0</span><span class="sxs-lookup"><span data-stu-id="4a671-162">1.0</span></span>|
|[<span data-ttu-id="4a671-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a671-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a671-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a671-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="4a671-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="4a671-165">EventType: String</span></span>

<span data-ttu-id="4a671-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="4a671-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4a671-167">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-167">Type</span></span>

*   <span data-ttu-id="4a671-168">String</span><span class="sxs-lookup"><span data-stu-id="4a671-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4a671-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4a671-169">Properties:</span></span>

| <span data-ttu-id="4a671-170">Имя</span><span class="sxs-lookup"><span data-stu-id="4a671-170">Name</span></span> | <span data-ttu-id="4a671-171">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-171">Type</span></span> | <span data-ttu-id="4a671-172">Описание</span><span class="sxs-lookup"><span data-stu-id="4a671-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="4a671-173">String</span><span class="sxs-lookup"><span data-stu-id="4a671-173">String</span></span> | <span data-ttu-id="4a671-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="4a671-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a671-175">Требования</span><span class="sxs-lookup"><span data-stu-id="4a671-175">Requirements</span></span>

|<span data-ttu-id="4a671-176">Требование</span><span class="sxs-lookup"><span data-stu-id="4a671-176">Requirement</span></span>| <span data-ttu-id="4a671-177">Значение</span><span class="sxs-lookup"><span data-stu-id="4a671-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a671-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4a671-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a671-179">1.5</span><span class="sxs-lookup"><span data-stu-id="4a671-179">1.5</span></span> |
|[<span data-ttu-id="4a671-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a671-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a671-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a671-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4a671-182">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="4a671-182">SourceProperty: String</span></span>

<span data-ttu-id="4a671-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="4a671-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4a671-184">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-184">Type</span></span>

*   <span data-ttu-id="4a671-185">String</span><span class="sxs-lookup"><span data-stu-id="4a671-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4a671-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4a671-186">Properties:</span></span>

|<span data-ttu-id="4a671-187">Имя</span><span class="sxs-lookup"><span data-stu-id="4a671-187">Name</span></span>| <span data-ttu-id="4a671-188">Тип</span><span class="sxs-lookup"><span data-stu-id="4a671-188">Type</span></span>| <span data-ttu-id="4a671-189">Описание</span><span class="sxs-lookup"><span data-stu-id="4a671-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4a671-190">String</span><span class="sxs-lookup"><span data-stu-id="4a671-190">String</span></span>|<span data-ttu-id="4a671-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="4a671-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4a671-192">String</span><span class="sxs-lookup"><span data-stu-id="4a671-192">String</span></span>|<span data-ttu-id="4a671-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="4a671-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a671-194">Требования</span><span class="sxs-lookup"><span data-stu-id="4a671-194">Requirements</span></span>

|<span data-ttu-id="4a671-195">Требование</span><span class="sxs-lookup"><span data-stu-id="4a671-195">Requirement</span></span>| <span data-ttu-id="4a671-196">Значение</span><span class="sxs-lookup"><span data-stu-id="4a671-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a671-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a671-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a671-198">1.0</span><span class="sxs-lookup"><span data-stu-id="4a671-198">1.0</span></span>|
|[<span data-ttu-id="4a671-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a671-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4a671-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a671-200">Compose or Read</span></span>|
