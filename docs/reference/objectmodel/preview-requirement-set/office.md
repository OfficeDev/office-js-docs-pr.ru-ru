---
title: Пространство имен Office — Предварительная версия набора требований
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: eb8ff0a755c1908d7b96438f96386056cc16b24f
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696437"
---
# <a name="office"></a><span data-ttu-id="6b0e2-102">Office</span><span class="sxs-lookup"><span data-stu-id="6b0e2-102">Office</span></span>

<span data-ttu-id="6b0e2-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6b0e2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b0e2-105">Требования</span><span class="sxs-lookup"><span data-stu-id="6b0e2-105">Requirements</span></span>

|<span data-ttu-id="6b0e2-106">Требование</span><span class="sxs-lookup"><span data-stu-id="6b0e2-106">Requirement</span></span>| <span data-ttu-id="6b0e2-107">Значение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b0e2-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b0e2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b0e2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6b0e2-109">1.0</span></span>|
|[<span data-ttu-id="6b0e2-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b0e2-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b0e2-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6b0e2-112">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="6b0e2-112">Members and methods</span></span>

| <span data-ttu-id="6b0e2-113">Элемент</span><span class="sxs-lookup"><span data-stu-id="6b0e2-113">Member</span></span> | <span data-ttu-id="6b0e2-114">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6b0e2-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6b0e2-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6b0e2-116">Member</span><span class="sxs-lookup"><span data-stu-id="6b0e2-116">Member</span></span> |
| [<span data-ttu-id="6b0e2-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6b0e2-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6b0e2-118">Member</span><span class="sxs-lookup"><span data-stu-id="6b0e2-118">Member</span></span> |
| [<span data-ttu-id="6b0e2-119">EventType</span><span class="sxs-lookup"><span data-stu-id="6b0e2-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6b0e2-120">Member</span><span class="sxs-lookup"><span data-stu-id="6b0e2-120">Member</span></span> |
| [<span data-ttu-id="6b0e2-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6b0e2-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6b0e2-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="6b0e2-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6b0e2-123">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="6b0e2-123">Namespaces</span></span>

<span data-ttu-id="6b0e2-124">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6b0e2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="6b0e2-126">Members</span><span class="sxs-lookup"><span data-stu-id="6b0e2-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6b0e2-127">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="6b0e2-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="6b0e2-128">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6b0e2-129">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-129">Type</span></span>

*   <span data-ttu-id="6b0e2-130">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b0e2-131">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6b0e2-131">Properties:</span></span>

|<span data-ttu-id="6b0e2-132">Имя</span><span class="sxs-lookup"><span data-stu-id="6b0e2-132">Name</span></span>| <span data-ttu-id="6b0e2-133">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-133">Type</span></span>| <span data-ttu-id="6b0e2-134">Описание</span><span class="sxs-lookup"><span data-stu-id="6b0e2-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6b0e2-135">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-135">String</span></span>|<span data-ttu-id="6b0e2-136">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6b0e2-137">Для указания</span><span class="sxs-lookup"><span data-stu-id="6b0e2-137">String</span></span>|<span data-ttu-id="6b0e2-138">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b0e2-139">Требования</span><span class="sxs-lookup"><span data-stu-id="6b0e2-139">Requirements</span></span>

|<span data-ttu-id="6b0e2-140">Требование</span><span class="sxs-lookup"><span data-stu-id="6b0e2-140">Requirement</span></span>| <span data-ttu-id="6b0e2-141">Значение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b0e2-142">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b0e2-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b0e2-143">1.0</span><span class="sxs-lookup"><span data-stu-id="6b0e2-143">1.0</span></span>|
|[<span data-ttu-id="6b0e2-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b0e2-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b0e2-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6b0e2-146">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="6b0e2-146">CoercionType: String</span></span>

<span data-ttu-id="6b0e2-147">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6b0e2-148">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-148">Type</span></span>

*   <span data-ttu-id="6b0e2-149">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b0e2-150">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6b0e2-150">Properties:</span></span>

|<span data-ttu-id="6b0e2-151">Имя</span><span class="sxs-lookup"><span data-stu-id="6b0e2-151">Name</span></span>| <span data-ttu-id="6b0e2-152">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-152">Type</span></span>| <span data-ttu-id="6b0e2-153">Описание</span><span class="sxs-lookup"><span data-stu-id="6b0e2-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6b0e2-154">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-154">String</span></span>|<span data-ttu-id="6b0e2-155">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6b0e2-156">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-156">String</span></span>|<span data-ttu-id="6b0e2-157">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b0e2-158">Требования</span><span class="sxs-lookup"><span data-stu-id="6b0e2-158">Requirements</span></span>

|<span data-ttu-id="6b0e2-159">Требование</span><span class="sxs-lookup"><span data-stu-id="6b0e2-159">Requirement</span></span>| <span data-ttu-id="6b0e2-160">Значение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b0e2-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b0e2-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b0e2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="6b0e2-162">1.0</span></span>|
|[<span data-ttu-id="6b0e2-163">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b0e2-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b0e2-164">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="6b0e2-165">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="6b0e2-165">EventType: String</span></span>

<span data-ttu-id="6b0e2-166">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6b0e2-167">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-167">Type</span></span>

*   <span data-ttu-id="6b0e2-168">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b0e2-169">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6b0e2-169">Properties:</span></span>

| <span data-ttu-id="6b0e2-170">Имя</span><span class="sxs-lookup"><span data-stu-id="6b0e2-170">Name</span></span> | <span data-ttu-id="6b0e2-171">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-171">Type</span></span> | <span data-ttu-id="6b0e2-172">Описание</span><span class="sxs-lookup"><span data-stu-id="6b0e2-172">Description</span></span> | <span data-ttu-id="6b0e2-173">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="6b0e2-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="6b0e2-174">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-174">String</span></span> | <span data-ttu-id="6b0e2-175">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6b0e2-176">1.7</span><span class="sxs-lookup"><span data-stu-id="6b0e2-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="6b0e2-177">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-177">String</span></span> | <span data-ttu-id="6b0e2-178">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="6b0e2-179">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6b0e2-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="6b0e2-180">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-180">String</span></span> | <span data-ttu-id="6b0e2-181">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="6b0e2-182">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6b0e2-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="6b0e2-183">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-183">String</span></span> | <span data-ttu-id="6b0e2-184">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="6b0e2-185">1.5</span><span class="sxs-lookup"><span data-stu-id="6b0e2-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="6b0e2-186">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-186">String</span></span> | <span data-ttu-id="6b0e2-187">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="6b0e2-188">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6b0e2-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6b0e2-189">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-189">String</span></span> | <span data-ttu-id="6b0e2-190">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6b0e2-191">1.7</span><span class="sxs-lookup"><span data-stu-id="6b0e2-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6b0e2-192">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-192">String</span></span> | <span data-ttu-id="6b0e2-193">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6b0e2-194">1.7</span><span class="sxs-lookup"><span data-stu-id="6b0e2-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6b0e2-195">Требования</span><span class="sxs-lookup"><span data-stu-id="6b0e2-195">Requirements</span></span>

|<span data-ttu-id="6b0e2-196">Требование</span><span class="sxs-lookup"><span data-stu-id="6b0e2-196">Requirement</span></span>| <span data-ttu-id="6b0e2-197">Значение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b0e2-198">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6b0e2-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b0e2-199">1.5</span><span class="sxs-lookup"><span data-stu-id="6b0e2-199">1.5</span></span> |
|[<span data-ttu-id="6b0e2-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b0e2-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b0e2-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6b0e2-202">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="6b0e2-202">SourceProperty: String</span></span>

<span data-ttu-id="6b0e2-203">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6b0e2-204">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-204">Type</span></span>

*   <span data-ttu-id="6b0e2-205">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6b0e2-206">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6b0e2-206">Properties:</span></span>

|<span data-ttu-id="6b0e2-207">Имя</span><span class="sxs-lookup"><span data-stu-id="6b0e2-207">Name</span></span>| <span data-ttu-id="6b0e2-208">Тип</span><span class="sxs-lookup"><span data-stu-id="6b0e2-208">Type</span></span>| <span data-ttu-id="6b0e2-209">Описание</span><span class="sxs-lookup"><span data-stu-id="6b0e2-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6b0e2-210">String</span><span class="sxs-lookup"><span data-stu-id="6b0e2-210">String</span></span>|<span data-ttu-id="6b0e2-211">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6b0e2-212">String.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-212">String</span></span>|<span data-ttu-id="6b0e2-213">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="6b0e2-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6b0e2-214">Требования</span><span class="sxs-lookup"><span data-stu-id="6b0e2-214">Requirements</span></span>

|<span data-ttu-id="6b0e2-215">Требование</span><span class="sxs-lookup"><span data-stu-id="6b0e2-215">Requirement</span></span>| <span data-ttu-id="6b0e2-216">Значение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b0e2-217">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b0e2-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b0e2-218">1.0</span><span class="sxs-lookup"><span data-stu-id="6b0e2-218">1.0</span></span>|
|[<span data-ttu-id="6b0e2-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b0e2-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b0e2-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b0e2-220">Compose or Read</span></span>|
