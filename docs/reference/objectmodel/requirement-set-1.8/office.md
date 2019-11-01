---
title: Пространство имен Office — набор обязательных элементов 1,8
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 91a0bef2a8280a068763c98b17644bd9268e2fb4
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902210"
---
# <a name="office"></a><span data-ttu-id="0f74b-102">Office</span><span class="sxs-lookup"><span data-stu-id="0f74b-102">Office</span></span>

<span data-ttu-id="0f74b-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0f74b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f74b-105">Требования</span><span class="sxs-lookup"><span data-stu-id="0f74b-105">Requirements</span></span>

|<span data-ttu-id="0f74b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="0f74b-106">Requirement</span></span>| <span data-ttu-id="0f74b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="0f74b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f74b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0f74b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f74b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0f74b-109">1.0</span></span>|
|[<span data-ttu-id="0f74b-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0f74b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f74b-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0f74b-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0f74b-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="0f74b-112">Members and methods</span></span>

| <span data-ttu-id="0f74b-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="0f74b-113">Member</span></span> | <span data-ttu-id="0f74b-114">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0f74b-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0f74b-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0f74b-116">Member</span><span class="sxs-lookup"><span data-stu-id="0f74b-116">Member</span></span> |
| [<span data-ttu-id="0f74b-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f74b-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0f74b-118">Member</span><span class="sxs-lookup"><span data-stu-id="0f74b-118">Member</span></span> |
| [<span data-ttu-id="0f74b-119">EventType</span><span class="sxs-lookup"><span data-stu-id="0f74b-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0f74b-120">Member</span><span class="sxs-lookup"><span data-stu-id="0f74b-120">Member</span></span> |
| [<span data-ttu-id="0f74b-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0f74b-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0f74b-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="0f74b-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0f74b-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="0f74b-123">Namespaces</span></span>

<span data-ttu-id="0f74b-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f74b-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0f74b-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="0f74b-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="0f74b-126">Members</span><span class="sxs-lookup"><span data-stu-id="0f74b-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0f74b-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="0f74b-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="0f74b-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0f74b-129">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-129">Type</span></span>

*   <span data-ttu-id="0f74b-130">String</span><span class="sxs-lookup"><span data-stu-id="0f74b-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f74b-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0f74b-131">Properties:</span></span>

|<span data-ttu-id="0f74b-132">Имя</span><span class="sxs-lookup"><span data-stu-id="0f74b-132">Name</span></span>| <span data-ttu-id="0f74b-133">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-133">Type</span></span>| <span data-ttu-id="0f74b-134">Описание</span><span class="sxs-lookup"><span data-stu-id="0f74b-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0f74b-135">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-135">String</span></span>|<span data-ttu-id="0f74b-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="0f74b-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0f74b-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="0f74b-137">String</span></span>|<span data-ttu-id="0f74b-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="0f74b-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f74b-139">Требования</span><span class="sxs-lookup"><span data-stu-id="0f74b-139">Requirements</span></span>

|<span data-ttu-id="0f74b-140">Требование</span><span class="sxs-lookup"><span data-stu-id="0f74b-140">Requirement</span></span>| <span data-ttu-id="0f74b-141">Значение</span><span class="sxs-lookup"><span data-stu-id="0f74b-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f74b-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0f74b-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f74b-143">1.0</span><span class="sxs-lookup"><span data-stu-id="0f74b-143">1.0</span></span>|
|[<span data-ttu-id="0f74b-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0f74b-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f74b-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0f74b-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="0f74b-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-146">CoercionType: String</span></span>

<span data-ttu-id="0f74b-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0f74b-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0f74b-148">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-148">Type</span></span>

*   <span data-ttu-id="0f74b-149">String</span><span class="sxs-lookup"><span data-stu-id="0f74b-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f74b-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0f74b-150">Properties:</span></span>

|<span data-ttu-id="0f74b-151">Имя</span><span class="sxs-lookup"><span data-stu-id="0f74b-151">Name</span></span>| <span data-ttu-id="0f74b-152">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-152">Type</span></span>| <span data-ttu-id="0f74b-153">Описание</span><span class="sxs-lookup"><span data-stu-id="0f74b-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0f74b-154">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-154">String</span></span>|<span data-ttu-id="0f74b-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="0f74b-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0f74b-156">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-156">String</span></span>|<span data-ttu-id="0f74b-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="0f74b-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f74b-158">Требования</span><span class="sxs-lookup"><span data-stu-id="0f74b-158">Requirements</span></span>

|<span data-ttu-id="0f74b-159">Требование</span><span class="sxs-lookup"><span data-stu-id="0f74b-159">Requirement</span></span>| <span data-ttu-id="0f74b-160">Значение</span><span class="sxs-lookup"><span data-stu-id="0f74b-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f74b-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0f74b-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f74b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="0f74b-162">1.0</span></span>|
|[<span data-ttu-id="0f74b-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0f74b-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f74b-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0f74b-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="0f74b-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-165">EventType: String</span></span>

<span data-ttu-id="0f74b-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="0f74b-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0f74b-167">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-167">Type</span></span>

*   <span data-ttu-id="0f74b-168">String</span><span class="sxs-lookup"><span data-stu-id="0f74b-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f74b-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0f74b-169">Properties:</span></span>

| <span data-ttu-id="0f74b-170">Имя</span><span class="sxs-lookup"><span data-stu-id="0f74b-170">Name</span></span> | <span data-ttu-id="0f74b-171">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-171">Type</span></span> | <span data-ttu-id="0f74b-172">Описание</span><span class="sxs-lookup"><span data-stu-id="0f74b-172">Description</span></span> | <span data-ttu-id="0f74b-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="0f74b-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="0f74b-174">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-174">String</span></span> | <span data-ttu-id="0f74b-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="0f74b-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="0f74b-176">1.7</span><span class="sxs-lookup"><span data-stu-id="0f74b-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="0f74b-177">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-177">String</span></span> | <span data-ttu-id="0f74b-178">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="0f74b-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="0f74b-179">1.8</span><span class="sxs-lookup"><span data-stu-id="0f74b-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="0f74b-180">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-180">String</span></span> | <span data-ttu-id="0f74b-181">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="0f74b-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="0f74b-182">1.8</span><span class="sxs-lookup"><span data-stu-id="0f74b-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="0f74b-183">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-183">String</span></span> | <span data-ttu-id="0f74b-184">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="0f74b-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="0f74b-185">1.5</span><span class="sxs-lookup"><span data-stu-id="0f74b-185">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="0f74b-186">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-186">String</span></span> | <span data-ttu-id="0f74b-187">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="0f74b-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="0f74b-188">1.7</span><span class="sxs-lookup"><span data-stu-id="0f74b-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="0f74b-189">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-189">String</span></span> | <span data-ttu-id="0f74b-190">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="0f74b-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="0f74b-191">1.7</span><span class="sxs-lookup"><span data-stu-id="0f74b-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f74b-192">Требования</span><span class="sxs-lookup"><span data-stu-id="0f74b-192">Requirements</span></span>

|<span data-ttu-id="0f74b-193">Требование</span><span class="sxs-lookup"><span data-stu-id="0f74b-193">Requirement</span></span>| <span data-ttu-id="0f74b-194">Значение</span><span class="sxs-lookup"><span data-stu-id="0f74b-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f74b-195">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0f74b-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f74b-196">1.5</span><span class="sxs-lookup"><span data-stu-id="0f74b-196">1.5</span></span> |
|[<span data-ttu-id="0f74b-197">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0f74b-197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f74b-198">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0f74b-198">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0f74b-199">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-199">SourceProperty: String</span></span>

<span data-ttu-id="0f74b-200">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0f74b-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0f74b-201">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-201">Type</span></span>

*   <span data-ttu-id="0f74b-202">String</span><span class="sxs-lookup"><span data-stu-id="0f74b-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f74b-203">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0f74b-203">Properties:</span></span>

|<span data-ttu-id="0f74b-204">Имя</span><span class="sxs-lookup"><span data-stu-id="0f74b-204">Name</span></span>| <span data-ttu-id="0f74b-205">Тип</span><span class="sxs-lookup"><span data-stu-id="0f74b-205">Type</span></span>| <span data-ttu-id="0f74b-206">Описание</span><span class="sxs-lookup"><span data-stu-id="0f74b-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0f74b-207">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-207">String</span></span>|<span data-ttu-id="0f74b-208">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0f74b-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0f74b-209">Строка</span><span class="sxs-lookup"><span data-stu-id="0f74b-209">String</span></span>|<span data-ttu-id="0f74b-210">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="0f74b-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f74b-211">Требования</span><span class="sxs-lookup"><span data-stu-id="0f74b-211">Requirements</span></span>

|<span data-ttu-id="0f74b-212">Требование</span><span class="sxs-lookup"><span data-stu-id="0f74b-212">Requirement</span></span>| <span data-ttu-id="0f74b-213">Значение</span><span class="sxs-lookup"><span data-stu-id="0f74b-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f74b-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0f74b-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f74b-215">1.0</span><span class="sxs-lookup"><span data-stu-id="0f74b-215">1.0</span></span>|
|[<span data-ttu-id="0f74b-216">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0f74b-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f74b-217">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0f74b-217">Compose or Read</span></span>|
