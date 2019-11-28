---
title: Пространство имен Office — Предварительная версия набора требований
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bd37b1be4d77d73cb56b0b2593ccc57dea6cab27
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629232"
---
# <a name="office"></a><span data-ttu-id="a5a93-102">Office</span><span class="sxs-lookup"><span data-stu-id="a5a93-102">Office</span></span>

<span data-ttu-id="a5a93-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a5a93-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5a93-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5a93-105">Requirements</span></span>

|<span data-ttu-id="a5a93-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a5a93-106">Requirement</span></span>| <span data-ttu-id="a5a93-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a5a93-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5a93-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5a93-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5a93-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-109">1.0</span></span>|
|[<span data-ttu-id="a5a93-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5a93-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a5a93-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a5a93-112">Properties</span><span class="sxs-lookup"><span data-stu-id="a5a93-112">Properties</span></span>

| <span data-ttu-id="a5a93-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="a5a93-113">Property</span></span> | <span data-ttu-id="a5a93-114">Способов</span><span class="sxs-lookup"><span data-stu-id="a5a93-114">Modes</span></span> | <span data-ttu-id="a5a93-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a5a93-115">Return type</span></span> | <span data-ttu-id="a5a93-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="a5a93-116">Minimum</span></span><br><span data-ttu-id="a5a93-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="a5a93-117">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="a5a93-118">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a5a93-118">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a5a93-119">Создание</span><span class="sxs-lookup"><span data-stu-id="a5a93-119">Compose</span></span><br><span data-ttu-id="a5a93-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-120">Read</span></span> | <span data-ttu-id="a5a93-121">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-121">String</span></span> | <span data-ttu-id="a5a93-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-122">1.0</span></span> |
| [<span data-ttu-id="a5a93-123">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a5a93-123">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a5a93-124">Создание</span><span class="sxs-lookup"><span data-stu-id="a5a93-124">Compose</span></span><br><span data-ttu-id="a5a93-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-125">Read</span></span> | <span data-ttu-id="a5a93-126">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-126">String</span></span> | <span data-ttu-id="a5a93-127">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-127">1.0</span></span> |
| [<span data-ttu-id="a5a93-128">EventType</span><span class="sxs-lookup"><span data-stu-id="a5a93-128">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a5a93-129">Создание</span><span class="sxs-lookup"><span data-stu-id="a5a93-129">Compose</span></span><br><span data-ttu-id="a5a93-130">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-130">Read</span></span> | <span data-ttu-id="a5a93-131">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-131">String</span></span> | <span data-ttu-id="a5a93-132">1.5</span><span class="sxs-lookup"><span data-stu-id="a5a93-132">1.5</span></span> |
| [<span data-ttu-id="a5a93-133">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a5a93-133">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a5a93-134">Создание</span><span class="sxs-lookup"><span data-stu-id="a5a93-134">Compose</span></span><br><span data-ttu-id="a5a93-135">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-135">Read</span></span> | <span data-ttu-id="a5a93-136">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-136">String</span></span> | <span data-ttu-id="a5a93-137">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-137">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a5a93-138">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="a5a93-138">Namespaces</span></span>

<span data-ttu-id="a5a93-139">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="a5a93-139">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a5a93-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): `ItemType`включает ряд перечислений, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a5a93-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="property-details"></a><span data-ttu-id="a5a93-141">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="a5a93-141">Property details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a5a93-142">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="a5a93-142">AsyncResultStatus: String</span></span>

<span data-ttu-id="a5a93-143">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="a5a93-143">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a5a93-144">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-144">Type</span></span>

*   <span data-ttu-id="a5a93-145">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-145">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5a93-146">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5a93-146">Properties:</span></span>

|<span data-ttu-id="a5a93-147">Имя</span><span class="sxs-lookup"><span data-stu-id="a5a93-147">Name</span></span>| <span data-ttu-id="a5a93-148">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-148">Type</span></span>| <span data-ttu-id="a5a93-149">Описание</span><span class="sxs-lookup"><span data-stu-id="a5a93-149">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a5a93-150">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-150">String</span></span>|<span data-ttu-id="a5a93-151">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="a5a93-151">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a5a93-152">Для указания</span><span class="sxs-lookup"><span data-stu-id="a5a93-152">String</span></span>|<span data-ttu-id="a5a93-153">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="a5a93-153">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5a93-154">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5a93-154">Requirements</span></span>

|<span data-ttu-id="a5a93-155">Требование</span><span class="sxs-lookup"><span data-stu-id="a5a93-155">Requirement</span></span>| <span data-ttu-id="a5a93-156">Значение</span><span class="sxs-lookup"><span data-stu-id="a5a93-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5a93-157">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5a93-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5a93-158">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-158">1.0</span></span>|
|[<span data-ttu-id="a5a93-159">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5a93-159">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a5a93-160">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-160">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a5a93-161">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="a5a93-161">CoercionType: String</span></span>

<span data-ttu-id="a5a93-162">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5a93-162">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5a93-163">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-163">Type</span></span>

*   <span data-ttu-id="a5a93-164">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-164">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5a93-165">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5a93-165">Properties:</span></span>

|<span data-ttu-id="a5a93-166">Имя</span><span class="sxs-lookup"><span data-stu-id="a5a93-166">Name</span></span>| <span data-ttu-id="a5a93-167">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-167">Type</span></span>| <span data-ttu-id="a5a93-168">Описание</span><span class="sxs-lookup"><span data-stu-id="a5a93-168">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a5a93-169">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-169">String</span></span>|<span data-ttu-id="a5a93-170">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="a5a93-170">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a5a93-171">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-171">String</span></span>|<span data-ttu-id="a5a93-172">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="a5a93-172">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5a93-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5a93-173">Requirements</span></span>

|<span data-ttu-id="a5a93-174">Требование</span><span class="sxs-lookup"><span data-stu-id="a5a93-174">Requirement</span></span>| <span data-ttu-id="a5a93-175">Значение</span><span class="sxs-lookup"><span data-stu-id="a5a93-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5a93-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5a93-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5a93-177">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-177">1.0</span></span>|
|[<span data-ttu-id="a5a93-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5a93-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a5a93-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-179">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="a5a93-180">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="a5a93-180">EventType: String</span></span>

<span data-ttu-id="a5a93-181">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="a5a93-181">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a5a93-182">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-182">Type</span></span>

*   <span data-ttu-id="a5a93-183">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-183">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5a93-184">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5a93-184">Properties:</span></span>

| <span data-ttu-id="a5a93-185">Имя</span><span class="sxs-lookup"><span data-stu-id="a5a93-185">Name</span></span> | <span data-ttu-id="a5a93-186">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-186">Type</span></span> | <span data-ttu-id="a5a93-187">Описание</span><span class="sxs-lookup"><span data-stu-id="a5a93-187">Description</span></span> | <span data-ttu-id="a5a93-188">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="a5a93-188">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="a5a93-189">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-189">String</span></span> | <span data-ttu-id="a5a93-190">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="a5a93-190">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a5a93-191">1.7</span><span class="sxs-lookup"><span data-stu-id="a5a93-191">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a5a93-192">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-192">String</span></span> | <span data-ttu-id="a5a93-193">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="a5a93-193">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a5a93-194">1.8</span><span class="sxs-lookup"><span data-stu-id="a5a93-194">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="a5a93-195">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-195">String</span></span> | <span data-ttu-id="a5a93-196">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="a5a93-196">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="a5a93-197">1.8</span><span class="sxs-lookup"><span data-stu-id="a5a93-197">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="a5a93-198">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-198">String</span></span> | <span data-ttu-id="a5a93-199">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="a5a93-199">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a5a93-200">1.5</span><span class="sxs-lookup"><span data-stu-id="a5a93-200">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a5a93-201">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-201">String</span></span> | <span data-ttu-id="a5a93-202">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="a5a93-202">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a5a93-203">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="a5a93-203">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a5a93-204">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-204">String</span></span> | <span data-ttu-id="a5a93-205">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="a5a93-205">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a5a93-206">1.7</span><span class="sxs-lookup"><span data-stu-id="a5a93-206">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a5a93-207">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-207">String</span></span> | <span data-ttu-id="a5a93-208">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="a5a93-208">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a5a93-209">1.7</span><span class="sxs-lookup"><span data-stu-id="a5a93-209">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a5a93-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5a93-210">Requirements</span></span>

|<span data-ttu-id="a5a93-211">Требование</span><span class="sxs-lookup"><span data-stu-id="a5a93-211">Requirement</span></span>| <span data-ttu-id="a5a93-212">Значение</span><span class="sxs-lookup"><span data-stu-id="a5a93-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5a93-213">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a5a93-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5a93-214">1.5</span><span class="sxs-lookup"><span data-stu-id="a5a93-214">1.5</span></span> |
|[<span data-ttu-id="a5a93-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5a93-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a5a93-216">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-216">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a5a93-217">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="a5a93-217">SourceProperty: String</span></span>

<span data-ttu-id="a5a93-218">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5a93-218">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5a93-219">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-219">Type</span></span>

*   <span data-ttu-id="a5a93-220">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-220">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5a93-221">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5a93-221">Properties:</span></span>

|<span data-ttu-id="a5a93-222">Имя</span><span class="sxs-lookup"><span data-stu-id="a5a93-222">Name</span></span>| <span data-ttu-id="a5a93-223">Тип</span><span class="sxs-lookup"><span data-stu-id="a5a93-223">Type</span></span>| <span data-ttu-id="a5a93-224">Описание</span><span class="sxs-lookup"><span data-stu-id="a5a93-224">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a5a93-225">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-225">String</span></span>|<span data-ttu-id="a5a93-226">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5a93-226">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a5a93-227">String</span><span class="sxs-lookup"><span data-stu-id="a5a93-227">String</span></span>|<span data-ttu-id="a5a93-228">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5a93-228">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5a93-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5a93-229">Requirements</span></span>

|<span data-ttu-id="a5a93-230">Требование</span><span class="sxs-lookup"><span data-stu-id="a5a93-230">Requirement</span></span>| <span data-ttu-id="a5a93-231">Значение</span><span class="sxs-lookup"><span data-stu-id="a5a93-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5a93-232">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5a93-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a5a93-233">1.0</span><span class="sxs-lookup"><span data-stu-id="a5a93-233">1.0</span></span>|
|[<span data-ttu-id="a5a93-234">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5a93-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a5a93-235">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5a93-235">Compose or Read</span></span>|
