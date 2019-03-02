---
title: Пространство имен Office — предварительная версия набора обязательных элементов
description: ''
ms.date: 02/26/2019
localization_priority: Normal
ms.openlocfilehash: 7b27963a85f1dcdaa6f269fce242c45bf1bdd146
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359235"
---
# <a name="office"></a><span data-ttu-id="a3d21-102">Office</span><span class="sxs-lookup"><span data-stu-id="a3d21-102">Office</span></span>

<span data-ttu-id="a3d21-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a3d21-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a3d21-105">Требования</span><span class="sxs-lookup"><span data-stu-id="a3d21-105">Requirements</span></span>

|<span data-ttu-id="a3d21-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a3d21-106">Requirement</span></span>| <span data-ttu-id="a3d21-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a3d21-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3d21-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a3d21-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3d21-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a3d21-109">1.0</span></span>|
|[<span data-ttu-id="a3d21-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a3d21-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3d21-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a3d21-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a3d21-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a3d21-112">Members and methods</span></span>

| <span data-ttu-id="a3d21-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3d21-113">Member</span></span> | <span data-ttu-id="a3d21-114">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a3d21-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a3d21-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a3d21-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3d21-116">Member</span></span> |
| [<span data-ttu-id="a3d21-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a3d21-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a3d21-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3d21-118">Member</span></span> |
| [<span data-ttu-id="a3d21-119">EventType</span><span class="sxs-lookup"><span data-stu-id="a3d21-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a3d21-120">Член</span><span class="sxs-lookup"><span data-stu-id="a3d21-120">Member</span></span> |
| [<span data-ttu-id="a3d21-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a3d21-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a3d21-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="a3d21-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a3d21-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="a3d21-123">Namespaces</span></span>

<span data-ttu-id="a3d21-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="a3d21-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a3d21-125">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a3d21-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a3d21-126">Элементы</span><span class="sxs-lookup"><span data-stu-id="a3d21-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a3d21-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a3d21-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="a3d21-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="a3d21-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a3d21-129">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-129">Type</span></span>

*   <span data-ttu-id="a3d21-130">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3d21-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a3d21-131">Properties:</span></span>

|<span data-ttu-id="a3d21-132">Имя</span><span class="sxs-lookup"><span data-stu-id="a3d21-132">Name</span></span>| <span data-ttu-id="a3d21-133">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-133">Type</span></span>| <span data-ttu-id="a3d21-134">Описание</span><span class="sxs-lookup"><span data-stu-id="a3d21-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a3d21-135">Строка</span><span class="sxs-lookup"><span data-stu-id="a3d21-135">String</span></span>|<span data-ttu-id="a3d21-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="a3d21-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a3d21-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="a3d21-137">String</span></span>|<span data-ttu-id="a3d21-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="a3d21-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3d21-139">Требования</span><span class="sxs-lookup"><span data-stu-id="a3d21-139">Requirements</span></span>

|<span data-ttu-id="a3d21-140">Требование</span><span class="sxs-lookup"><span data-stu-id="a3d21-140">Requirement</span></span>| <span data-ttu-id="a3d21-141">Значение</span><span class="sxs-lookup"><span data-stu-id="a3d21-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3d21-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a3d21-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3d21-143">1.0</span><span class="sxs-lookup"><span data-stu-id="a3d21-143">1.0</span></span>|
|[<span data-ttu-id="a3d21-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a3d21-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3d21-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a3d21-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="a3d21-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a3d21-146">CoercionType :String</span></span>

<span data-ttu-id="a3d21-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a3d21-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a3d21-148">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-148">Type</span></span>

*   <span data-ttu-id="a3d21-149">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3d21-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a3d21-150">Properties:</span></span>

|<span data-ttu-id="a3d21-151">Имя</span><span class="sxs-lookup"><span data-stu-id="a3d21-151">Name</span></span>| <span data-ttu-id="a3d21-152">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-152">Type</span></span>| <span data-ttu-id="a3d21-153">Описание</span><span class="sxs-lookup"><span data-stu-id="a3d21-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a3d21-154">Строка</span><span class="sxs-lookup"><span data-stu-id="a3d21-154">String</span></span>|<span data-ttu-id="a3d21-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="a3d21-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a3d21-156">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-156">String</span></span>|<span data-ttu-id="a3d21-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="a3d21-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3d21-158">Требования</span><span class="sxs-lookup"><span data-stu-id="a3d21-158">Requirements</span></span>

|<span data-ttu-id="a3d21-159">Требование</span><span class="sxs-lookup"><span data-stu-id="a3d21-159">Requirement</span></span>| <span data-ttu-id="a3d21-160">Значение</span><span class="sxs-lookup"><span data-stu-id="a3d21-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3d21-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a3d21-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3d21-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a3d21-162">1.0</span></span>|
|[<span data-ttu-id="a3d21-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a3d21-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3d21-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a3d21-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="a3d21-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="a3d21-165">EventType :String</span></span>

<span data-ttu-id="a3d21-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="a3d21-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a3d21-167">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-167">Type</span></span>

*   <span data-ttu-id="a3d21-168">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3d21-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a3d21-169">Properties:</span></span>

| <span data-ttu-id="a3d21-170">Имя</span><span class="sxs-lookup"><span data-stu-id="a3d21-170">Name</span></span> | <span data-ttu-id="a3d21-171">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-171">Type</span></span> | <span data-ttu-id="a3d21-172">Описание</span><span class="sxs-lookup"><span data-stu-id="a3d21-172">Description</span></span> | <span data-ttu-id="a3d21-173">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="a3d21-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="a3d21-174">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-174">String</span></span> | <span data-ttu-id="a3d21-175">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="a3d21-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a3d21-176">1.7</span><span class="sxs-lookup"><span data-stu-id="a3d21-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a3d21-177">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-177">String</span></span> | <span data-ttu-id="a3d21-178">Было добавлено или удалено вложение для элемента.</span><span class="sxs-lookup"><span data-stu-id="a3d21-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a3d21-179">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="a3d21-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="a3d21-180">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-180">String</span></span> | <span data-ttu-id="a3d21-181">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="a3d21-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="a3d21-182">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="a3d21-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="a3d21-183">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-183">String</span></span> | <span data-ttu-id="a3d21-184">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="a3d21-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a3d21-185">1.5</span><span class="sxs-lookup"><span data-stu-id="a3d21-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a3d21-186">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-186">String</span></span> | <span data-ttu-id="a3d21-187">Тема Office в почтовом ящике была изменена.</span><span class="sxs-lookup"><span data-stu-id="a3d21-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a3d21-188">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="a3d21-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a3d21-189">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-189">String</span></span> | <span data-ttu-id="a3d21-190">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="a3d21-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a3d21-191">1.7</span><span class="sxs-lookup"><span data-stu-id="a3d21-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a3d21-192">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-192">String</span></span> | <span data-ttu-id="a3d21-193">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="a3d21-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a3d21-194">1.7</span><span class="sxs-lookup"><span data-stu-id="a3d21-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a3d21-195">Требования</span><span class="sxs-lookup"><span data-stu-id="a3d21-195">Requirements</span></span>

|<span data-ttu-id="a3d21-196">Требование</span><span class="sxs-lookup"><span data-stu-id="a3d21-196">Requirement</span></span>| <span data-ttu-id="a3d21-197">Значение</span><span class="sxs-lookup"><span data-stu-id="a3d21-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3d21-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a3d21-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3d21-199">1.5</span><span class="sxs-lookup"><span data-stu-id="a3d21-199">1.5</span></span> |
|[<span data-ttu-id="a3d21-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a3d21-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3d21-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a3d21-201">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="a3d21-202">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a3d21-202">SourceProperty :String</span></span>

<span data-ttu-id="a3d21-203">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a3d21-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a3d21-204">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-204">Type</span></span>

*   <span data-ttu-id="a3d21-205">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a3d21-206">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a3d21-206">Properties:</span></span>

|<span data-ttu-id="a3d21-207">Имя</span><span class="sxs-lookup"><span data-stu-id="a3d21-207">Name</span></span>| <span data-ttu-id="a3d21-208">Тип</span><span class="sxs-lookup"><span data-stu-id="a3d21-208">Type</span></span>| <span data-ttu-id="a3d21-209">Описание</span><span class="sxs-lookup"><span data-stu-id="a3d21-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a3d21-210">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-210">String</span></span>|<span data-ttu-id="a3d21-211">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a3d21-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a3d21-212">String</span><span class="sxs-lookup"><span data-stu-id="a3d21-212">String</span></span>|<span data-ttu-id="a3d21-213">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="a3d21-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a3d21-214">Требования</span><span class="sxs-lookup"><span data-stu-id="a3d21-214">Requirements</span></span>

|<span data-ttu-id="a3d21-215">Требование</span><span class="sxs-lookup"><span data-stu-id="a3d21-215">Requirement</span></span>| <span data-ttu-id="a3d21-216">Значение</span><span class="sxs-lookup"><span data-stu-id="a3d21-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="a3d21-217">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a3d21-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a3d21-218">1.0</span><span class="sxs-lookup"><span data-stu-id="a3d21-218">1.0</span></span>|
|[<span data-ttu-id="a3d21-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a3d21-219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a3d21-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a3d21-220">Compose or Read</span></span>|
