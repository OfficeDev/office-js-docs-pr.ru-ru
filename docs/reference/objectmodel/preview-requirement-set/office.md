---
title: Пространство имен Office — Предварительная версия набора требований
description: Элементы пространства имен Office, доступные для надстроек Outlook с использованием набора обязательных элементов API почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d72e5c78a7fd8d3c00b8f84e7d9b05ee6defc0c5
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890860"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="b9638-103">Office (набор требований предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="b9638-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="b9638-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b9638-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9638-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9638-106">Requirements</span></span>

|<span data-ttu-id="b9638-107">Требование</span><span class="sxs-lookup"><span data-stu-id="b9638-107">Requirement</span></span>| <span data-ttu-id="b9638-108">Значение</span><span class="sxs-lookup"><span data-stu-id="b9638-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9638-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9638-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9638-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-110">1.1</span></span>|
|[<span data-ttu-id="b9638-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9638-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9638-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b9638-113">Properties</span><span class="sxs-lookup"><span data-stu-id="b9638-113">Properties</span></span>

| <span data-ttu-id="b9638-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="b9638-114">Property</span></span> | <span data-ttu-id="b9638-115">Способов</span><span class="sxs-lookup"><span data-stu-id="b9638-115">Modes</span></span> | <span data-ttu-id="b9638-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b9638-116">Return type</span></span> | <span data-ttu-id="b9638-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b9638-117">Minimum</span></span><br><span data-ttu-id="b9638-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="b9638-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b9638-119">контекст</span><span class="sxs-lookup"><span data-stu-id="b9638-119">context</span></span>](office.context.md) | <span data-ttu-id="b9638-120">Создание</span><span class="sxs-lookup"><span data-stu-id="b9638-120">Compose</span></span><br><span data-ttu-id="b9638-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-121">Read</span></span> | [<span data-ttu-id="b9638-122">Context</span><span class="sxs-lookup"><span data-stu-id="b9638-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="b9638-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="b9638-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="b9638-124">Enumerations</span></span>

| <span data-ttu-id="b9638-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="b9638-125">Enumeration</span></span> | <span data-ttu-id="b9638-126">Способов</span><span class="sxs-lookup"><span data-stu-id="b9638-126">Modes</span></span> | <span data-ttu-id="b9638-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b9638-127">Return type</span></span> | <span data-ttu-id="b9638-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="b9638-128">Minimum</span></span><br><span data-ttu-id="b9638-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="b9638-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b9638-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b9638-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b9638-131">Создание</span><span class="sxs-lookup"><span data-stu-id="b9638-131">Compose</span></span><br><span data-ttu-id="b9638-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-132">Read</span></span> | <span data-ttu-id="b9638-133">String</span><span class="sxs-lookup"><span data-stu-id="b9638-133">String</span></span> | [<span data-ttu-id="b9638-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9638-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b9638-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b9638-136">Создание</span><span class="sxs-lookup"><span data-stu-id="b9638-136">Compose</span></span><br><span data-ttu-id="b9638-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-137">Read</span></span> | <span data-ttu-id="b9638-138">String</span><span class="sxs-lookup"><span data-stu-id="b9638-138">String</span></span> | [<span data-ttu-id="b9638-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9638-140">EventType</span><span class="sxs-lookup"><span data-stu-id="b9638-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b9638-141">Создание</span><span class="sxs-lookup"><span data-stu-id="b9638-141">Compose</span></span><br><span data-ttu-id="b9638-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-142">Read</span></span> | <span data-ttu-id="b9638-143">String</span><span class="sxs-lookup"><span data-stu-id="b9638-143">String</span></span> | [<span data-ttu-id="b9638-144">1,5</span><span class="sxs-lookup"><span data-stu-id="b9638-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b9638-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b9638-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b9638-146">Создание</span><span class="sxs-lookup"><span data-stu-id="b9638-146">Compose</span></span><br><span data-ttu-id="b9638-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-147">Read</span></span> | <span data-ttu-id="b9638-148">String</span><span class="sxs-lookup"><span data-stu-id="b9638-148">String</span></span> | [<span data-ttu-id="b9638-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="b9638-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="b9638-150">Namespaces</span></span>

<span data-ttu-id="b9638-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="b9638-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b9638-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="b9638-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b9638-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="b9638-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="b9638-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="b9638-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b9638-155">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-155">Type</span></span>

*   <span data-ttu-id="b9638-156">String</span><span class="sxs-lookup"><span data-stu-id="b9638-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9638-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b9638-157">Properties:</span></span>

|<span data-ttu-id="b9638-158">Имя</span><span class="sxs-lookup"><span data-stu-id="b9638-158">Name</span></span>| <span data-ttu-id="b9638-159">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-159">Type</span></span>| <span data-ttu-id="b9638-160">Описание</span><span class="sxs-lookup"><span data-stu-id="b9638-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b9638-161">String</span><span class="sxs-lookup"><span data-stu-id="b9638-161">String</span></span>|<span data-ttu-id="b9638-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="b9638-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b9638-163">Для указания</span><span class="sxs-lookup"><span data-stu-id="b9638-163">String</span></span>|<span data-ttu-id="b9638-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="b9638-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9638-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9638-165">Requirements</span></span>

|<span data-ttu-id="b9638-166">Требование</span><span class="sxs-lookup"><span data-stu-id="b9638-166">Requirement</span></span>| <span data-ttu-id="b9638-167">Значение</span><span class="sxs-lookup"><span data-stu-id="b9638-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9638-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9638-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9638-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-169">1.1</span></span>|
|[<span data-ttu-id="b9638-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9638-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9638-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b9638-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="b9638-172">CoercionType: String</span></span>

<span data-ttu-id="b9638-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b9638-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9638-174">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-174">Type</span></span>

*   <span data-ttu-id="b9638-175">String</span><span class="sxs-lookup"><span data-stu-id="b9638-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9638-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b9638-176">Properties:</span></span>

|<span data-ttu-id="b9638-177">Имя</span><span class="sxs-lookup"><span data-stu-id="b9638-177">Name</span></span>| <span data-ttu-id="b9638-178">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-178">Type</span></span>| <span data-ttu-id="b9638-179">Описание</span><span class="sxs-lookup"><span data-stu-id="b9638-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b9638-180">String</span><span class="sxs-lookup"><span data-stu-id="b9638-180">String</span></span>|<span data-ttu-id="b9638-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="b9638-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b9638-182">String</span><span class="sxs-lookup"><span data-stu-id="b9638-182">String</span></span>|<span data-ttu-id="b9638-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="b9638-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9638-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9638-184">Requirements</span></span>

|<span data-ttu-id="b9638-185">Требование</span><span class="sxs-lookup"><span data-stu-id="b9638-185">Requirement</span></span>| <span data-ttu-id="b9638-186">Значение</span><span class="sxs-lookup"><span data-stu-id="b9638-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9638-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9638-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9638-188">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-188">1.1</span></span>|
|[<span data-ttu-id="b9638-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9638-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9638-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="b9638-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="b9638-191">EventType: String</span></span>

<span data-ttu-id="b9638-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="b9638-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b9638-193">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-193">Type</span></span>

*   <span data-ttu-id="b9638-194">String</span><span class="sxs-lookup"><span data-stu-id="b9638-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9638-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b9638-195">Properties:</span></span>

| <span data-ttu-id="b9638-196">Имя</span><span class="sxs-lookup"><span data-stu-id="b9638-196">Name</span></span> | <span data-ttu-id="b9638-197">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-197">Type</span></span> | <span data-ttu-id="b9638-198">Описание</span><span class="sxs-lookup"><span data-stu-id="b9638-198">Description</span></span> | <span data-ttu-id="b9638-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="b9638-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="b9638-200">String</span><span class="sxs-lookup"><span data-stu-id="b9638-200">String</span></span> | <span data-ttu-id="b9638-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="b9638-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="b9638-202">1.7</span><span class="sxs-lookup"><span data-stu-id="b9638-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="b9638-203">String</span><span class="sxs-lookup"><span data-stu-id="b9638-203">String</span></span> | <span data-ttu-id="b9638-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="b9638-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="b9638-205">1.8</span><span class="sxs-lookup"><span data-stu-id="b9638-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="b9638-206">String</span><span class="sxs-lookup"><span data-stu-id="b9638-206">String</span></span> | <span data-ttu-id="b9638-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="b9638-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="b9638-208">1.8</span><span class="sxs-lookup"><span data-stu-id="b9638-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="b9638-209">String</span><span class="sxs-lookup"><span data-stu-id="b9638-209">String</span></span> | <span data-ttu-id="b9638-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="b9638-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="b9638-211">1.5</span><span class="sxs-lookup"><span data-stu-id="b9638-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="b9638-212">String</span><span class="sxs-lookup"><span data-stu-id="b9638-212">String</span></span> | <span data-ttu-id="b9638-213">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="b9638-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="b9638-214">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="b9638-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="b9638-215">String</span><span class="sxs-lookup"><span data-stu-id="b9638-215">String</span></span> | <span data-ttu-id="b9638-216">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="b9638-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="b9638-217">1.7</span><span class="sxs-lookup"><span data-stu-id="b9638-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="b9638-218">String</span><span class="sxs-lookup"><span data-stu-id="b9638-218">String</span></span> | <span data-ttu-id="b9638-219">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="b9638-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="b9638-220">1.7</span><span class="sxs-lookup"><span data-stu-id="b9638-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b9638-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9638-221">Requirements</span></span>

|<span data-ttu-id="b9638-222">Требование</span><span class="sxs-lookup"><span data-stu-id="b9638-222">Requirement</span></span>| <span data-ttu-id="b9638-223">Значение</span><span class="sxs-lookup"><span data-stu-id="b9638-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9638-224">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b9638-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9638-225">1.5</span><span class="sxs-lookup"><span data-stu-id="b9638-225">1.5</span></span> |
|[<span data-ttu-id="b9638-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9638-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9638-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b9638-228">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="b9638-228">SourceProperty: String</span></span>

<span data-ttu-id="b9638-229">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="b9638-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b9638-230">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-230">Type</span></span>

*   <span data-ttu-id="b9638-231">String</span><span class="sxs-lookup"><span data-stu-id="b9638-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b9638-232">Свойства:</span><span class="sxs-lookup"><span data-stu-id="b9638-232">Properties:</span></span>

|<span data-ttu-id="b9638-233">Имя</span><span class="sxs-lookup"><span data-stu-id="b9638-233">Name</span></span>| <span data-ttu-id="b9638-234">Тип</span><span class="sxs-lookup"><span data-stu-id="b9638-234">Type</span></span>| <span data-ttu-id="b9638-235">Описание</span><span class="sxs-lookup"><span data-stu-id="b9638-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b9638-236">String</span><span class="sxs-lookup"><span data-stu-id="b9638-236">String</span></span>|<span data-ttu-id="b9638-237">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="b9638-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b9638-238">String</span><span class="sxs-lookup"><span data-stu-id="b9638-238">String</span></span>|<span data-ttu-id="b9638-239">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="b9638-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b9638-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="b9638-240">Requirements</span></span>

|<span data-ttu-id="b9638-241">Требование</span><span class="sxs-lookup"><span data-stu-id="b9638-241">Requirement</span></span>| <span data-ttu-id="b9638-242">Значение</span><span class="sxs-lookup"><span data-stu-id="b9638-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9638-243">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9638-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9638-244">1.1</span><span class="sxs-lookup"><span data-stu-id="b9638-244">1.1</span></span>|
|[<span data-ttu-id="b9638-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9638-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9638-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9638-246">Compose or Read</span></span>|
