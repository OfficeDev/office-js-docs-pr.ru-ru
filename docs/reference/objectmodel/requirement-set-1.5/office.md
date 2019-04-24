---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450263"
---
# <a name="office"></a><span data-ttu-id="a2961-102">Office</span><span class="sxs-lookup"><span data-stu-id="a2961-102">Office</span></span>

<span data-ttu-id="a2961-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a2961-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a2961-105">Требования</span><span class="sxs-lookup"><span data-stu-id="a2961-105">Requirements</span></span>

|<span data-ttu-id="a2961-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a2961-106">Requirement</span></span>| <span data-ttu-id="a2961-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a2961-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2961-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a2961-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2961-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a2961-109">1.0</span></span>|
|[<span data-ttu-id="a2961-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a2961-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2961-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a2961-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a2961-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a2961-112">Members and methods</span></span>

| <span data-ttu-id="a2961-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="a2961-113">Member</span></span> | <span data-ttu-id="a2961-114">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a2961-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a2961-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a2961-116">Member</span><span class="sxs-lookup"><span data-stu-id="a2961-116">Member</span></span> |
| [<span data-ttu-id="a2961-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a2961-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a2961-118">Member</span><span class="sxs-lookup"><span data-stu-id="a2961-118">Member</span></span> |
| [<span data-ttu-id="a2961-119">EventType</span><span class="sxs-lookup"><span data-stu-id="a2961-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a2961-120">Member</span><span class="sxs-lookup"><span data-stu-id="a2961-120">Member</span></span> |
| [<span data-ttu-id="a2961-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a2961-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a2961-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="a2961-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a2961-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="a2961-123">Namespaces</span></span>

<span data-ttu-id="a2961-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="a2961-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a2961-125">[MailboxEnums.](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a2961-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a2961-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="a2961-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a2961-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a2961-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="a2961-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="a2961-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a2961-129">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-129">Type</span></span>

*   <span data-ttu-id="a2961-130">String</span><span class="sxs-lookup"><span data-stu-id="a2961-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2961-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a2961-131">Properties:</span></span>

|<span data-ttu-id="a2961-132">Имя</span><span class="sxs-lookup"><span data-stu-id="a2961-132">Name</span></span>| <span data-ttu-id="a2961-133">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-133">Type</span></span>| <span data-ttu-id="a2961-134">Описание</span><span class="sxs-lookup"><span data-stu-id="a2961-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a2961-135">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-135">String</span></span>|<span data-ttu-id="a2961-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="a2961-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a2961-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="a2961-137">String</span></span>|<span data-ttu-id="a2961-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="a2961-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2961-139">Требования</span><span class="sxs-lookup"><span data-stu-id="a2961-139">Requirements</span></span>

|<span data-ttu-id="a2961-140">Требование</span><span class="sxs-lookup"><span data-stu-id="a2961-140">Requirement</span></span>| <span data-ttu-id="a2961-141">Значение</span><span class="sxs-lookup"><span data-stu-id="a2961-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2961-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a2961-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2961-143">1.0</span><span class="sxs-lookup"><span data-stu-id="a2961-143">1.0</span></span>|
|[<span data-ttu-id="a2961-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a2961-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2961-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a2961-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="a2961-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a2961-146">CoercionType :String</span></span>

<span data-ttu-id="a2961-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a2961-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2961-148">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-148">Type</span></span>

*   <span data-ttu-id="a2961-149">String</span><span class="sxs-lookup"><span data-stu-id="a2961-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2961-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a2961-150">Properties:</span></span>

|<span data-ttu-id="a2961-151">Имя</span><span class="sxs-lookup"><span data-stu-id="a2961-151">Name</span></span>| <span data-ttu-id="a2961-152">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-152">Type</span></span>| <span data-ttu-id="a2961-153">Описание</span><span class="sxs-lookup"><span data-stu-id="a2961-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a2961-154">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-154">String</span></span>|<span data-ttu-id="a2961-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="a2961-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a2961-156">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-156">String</span></span>|<span data-ttu-id="a2961-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="a2961-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2961-158">Требования</span><span class="sxs-lookup"><span data-stu-id="a2961-158">Requirements</span></span>

|<span data-ttu-id="a2961-159">Требование</span><span class="sxs-lookup"><span data-stu-id="a2961-159">Requirement</span></span>| <span data-ttu-id="a2961-160">Значение</span><span class="sxs-lookup"><span data-stu-id="a2961-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2961-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a2961-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2961-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a2961-162">1.0</span></span>|
|[<span data-ttu-id="a2961-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a2961-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2961-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a2961-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="a2961-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="a2961-165">EventType :String</span></span>

<span data-ttu-id="a2961-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="a2961-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a2961-167">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-167">Type</span></span>

*   <span data-ttu-id="a2961-168">String</span><span class="sxs-lookup"><span data-stu-id="a2961-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2961-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a2961-169">Properties:</span></span>

| <span data-ttu-id="a2961-170">Имя</span><span class="sxs-lookup"><span data-stu-id="a2961-170">Name</span></span> | <span data-ttu-id="a2961-171">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-171">Type</span></span> | <span data-ttu-id="a2961-172">Описание</span><span class="sxs-lookup"><span data-stu-id="a2961-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="a2961-173">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-173">String</span></span> | <span data-ttu-id="a2961-174">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="a2961-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a2961-175">Требования</span><span class="sxs-lookup"><span data-stu-id="a2961-175">Requirements</span></span>

|<span data-ttu-id="a2961-176">Требование</span><span class="sxs-lookup"><span data-stu-id="a2961-176">Requirement</span></span>| <span data-ttu-id="a2961-177">Значение</span><span class="sxs-lookup"><span data-stu-id="a2961-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2961-178">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a2961-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2961-179">1.5</span><span class="sxs-lookup"><span data-stu-id="a2961-179">1.5</span></span> |
|[<span data-ttu-id="a2961-180">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a2961-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2961-181">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a2961-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="a2961-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a2961-182">SourceProperty :String</span></span>

<span data-ttu-id="a2961-183">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a2961-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a2961-184">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-184">Type</span></span>

*   <span data-ttu-id="a2961-185">String</span><span class="sxs-lookup"><span data-stu-id="a2961-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a2961-186">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a2961-186">Properties:</span></span>

|<span data-ttu-id="a2961-187">Имя</span><span class="sxs-lookup"><span data-stu-id="a2961-187">Name</span></span>| <span data-ttu-id="a2961-188">Тип</span><span class="sxs-lookup"><span data-stu-id="a2961-188">Type</span></span>| <span data-ttu-id="a2961-189">Описание</span><span class="sxs-lookup"><span data-stu-id="a2961-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a2961-190">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-190">String</span></span>|<span data-ttu-id="a2961-191">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a2961-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a2961-192">Строка</span><span class="sxs-lookup"><span data-stu-id="a2961-192">String</span></span>|<span data-ttu-id="a2961-193">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="a2961-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a2961-194">Требования</span><span class="sxs-lookup"><span data-stu-id="a2961-194">Requirements</span></span>

|<span data-ttu-id="a2961-195">Требование</span><span class="sxs-lookup"><span data-stu-id="a2961-195">Requirement</span></span>| <span data-ttu-id="a2961-196">Значение</span><span class="sxs-lookup"><span data-stu-id="a2961-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="a2961-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a2961-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a2961-198">1.0</span><span class="sxs-lookup"><span data-stu-id="a2961-198">1.0</span></span>|
|[<span data-ttu-id="a2961-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a2961-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a2961-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a2961-200">Compose or Read</span></span>|
