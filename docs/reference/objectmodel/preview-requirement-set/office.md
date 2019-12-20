---
title: Пространство имен Office — Предварительная версия набора требований
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ef9634058fcdc633e9ad3a0adb74c4abebf8038b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815062"
---
# <a name="office"></a><span data-ttu-id="57ecf-102">Office</span><span class="sxs-lookup"><span data-stu-id="57ecf-102">Office</span></span>

<span data-ttu-id="57ecf-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="57ecf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="57ecf-105">Требования</span><span class="sxs-lookup"><span data-stu-id="57ecf-105">Requirements</span></span>

|<span data-ttu-id="57ecf-106">Требование</span><span class="sxs-lookup"><span data-stu-id="57ecf-106">Requirement</span></span>| <span data-ttu-id="57ecf-107">Значение</span><span class="sxs-lookup"><span data-stu-id="57ecf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ecf-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="57ecf-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="57ecf-109">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-109">1.1</span></span>|
|[<span data-ttu-id="57ecf-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="57ecf-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ecf-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="57ecf-112">Properties</span><span class="sxs-lookup"><span data-stu-id="57ecf-112">Properties</span></span>

| <span data-ttu-id="57ecf-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="57ecf-113">Property</span></span> | <span data-ttu-id="57ecf-114">Способов</span><span class="sxs-lookup"><span data-stu-id="57ecf-114">Modes</span></span> | <span data-ttu-id="57ecf-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="57ecf-115">Return type</span></span> | <span data-ttu-id="57ecf-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="57ecf-116">Minimum</span></span><br><span data-ttu-id="57ecf-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="57ecf-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="57ecf-118">контекст</span><span class="sxs-lookup"><span data-stu-id="57ecf-118">context</span></span>](office.context.md) | <span data-ttu-id="57ecf-119">Создание</span><span class="sxs-lookup"><span data-stu-id="57ecf-119">Compose</span></span><br><span data-ttu-id="57ecf-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-120">Read</span></span> | [<span data-ttu-id="57ecf-121">Context</span><span class="sxs-lookup"><span data-stu-id="57ecf-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="57ecf-122">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="57ecf-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="57ecf-123">Enumerations</span></span>

| <span data-ttu-id="57ecf-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="57ecf-124">Enumeration</span></span> | <span data-ttu-id="57ecf-125">Способов</span><span class="sxs-lookup"><span data-stu-id="57ecf-125">Modes</span></span> | <span data-ttu-id="57ecf-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="57ecf-126">Return type</span></span> | <span data-ttu-id="57ecf-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="57ecf-127">Minimum</span></span><br><span data-ttu-id="57ecf-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="57ecf-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="57ecf-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="57ecf-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="57ecf-130">Создание</span><span class="sxs-lookup"><span data-stu-id="57ecf-130">Compose</span></span><br><span data-ttu-id="57ecf-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-131">Read</span></span> | <span data-ttu-id="57ecf-132">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-132">String</span></span> | [<span data-ttu-id="57ecf-133">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="57ecf-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="57ecf-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="57ecf-135">Создание</span><span class="sxs-lookup"><span data-stu-id="57ecf-135">Compose</span></span><br><span data-ttu-id="57ecf-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-136">Read</span></span> | <span data-ttu-id="57ecf-137">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-137">String</span></span> | [<span data-ttu-id="57ecf-138">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="57ecf-139">EventType</span><span class="sxs-lookup"><span data-stu-id="57ecf-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="57ecf-140">Создание</span><span class="sxs-lookup"><span data-stu-id="57ecf-140">Compose</span></span><br><span data-ttu-id="57ecf-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-141">Read</span></span> | <span data-ttu-id="57ecf-142">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-142">String</span></span> | [<span data-ttu-id="57ecf-143">1,5</span><span class="sxs-lookup"><span data-stu-id="57ecf-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="57ecf-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="57ecf-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="57ecf-145">Создание</span><span class="sxs-lookup"><span data-stu-id="57ecf-145">Compose</span></span><br><span data-ttu-id="57ecf-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-146">Read</span></span> | <span data-ttu-id="57ecf-147">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-147">String</span></span> | [<span data-ttu-id="57ecf-148">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="57ecf-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="57ecf-149">Namespaces</span></span>

<span data-ttu-id="57ecf-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="57ecf-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="57ecf-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="57ecf-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="57ecf-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="57ecf-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="57ecf-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="57ecf-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="57ecf-154">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-154">Type</span></span>

*   <span data-ttu-id="57ecf-155">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57ecf-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="57ecf-156">Properties:</span></span>

|<span data-ttu-id="57ecf-157">Имя</span><span class="sxs-lookup"><span data-stu-id="57ecf-157">Name</span></span>| <span data-ttu-id="57ecf-158">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-158">Type</span></span>| <span data-ttu-id="57ecf-159">Описание</span><span class="sxs-lookup"><span data-stu-id="57ecf-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="57ecf-160">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-160">String</span></span>|<span data-ttu-id="57ecf-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="57ecf-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="57ecf-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="57ecf-162">String</span></span>|<span data-ttu-id="57ecf-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="57ecf-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57ecf-164">Требования</span><span class="sxs-lookup"><span data-stu-id="57ecf-164">Requirements</span></span>

|<span data-ttu-id="57ecf-165">Требование</span><span class="sxs-lookup"><span data-stu-id="57ecf-165">Requirement</span></span>| <span data-ttu-id="57ecf-166">Значение</span><span class="sxs-lookup"><span data-stu-id="57ecf-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ecf-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="57ecf-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="57ecf-168">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-168">1.1</span></span>|
|[<span data-ttu-id="57ecf-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="57ecf-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ecf-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="57ecf-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="57ecf-171">CoercionType: String</span></span>

<span data-ttu-id="57ecf-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="57ecf-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="57ecf-173">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-173">Type</span></span>

*   <span data-ttu-id="57ecf-174">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57ecf-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="57ecf-175">Properties:</span></span>

|<span data-ttu-id="57ecf-176">Имя</span><span class="sxs-lookup"><span data-stu-id="57ecf-176">Name</span></span>| <span data-ttu-id="57ecf-177">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-177">Type</span></span>| <span data-ttu-id="57ecf-178">Описание</span><span class="sxs-lookup"><span data-stu-id="57ecf-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="57ecf-179">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-179">String</span></span>|<span data-ttu-id="57ecf-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="57ecf-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="57ecf-181">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-181">String</span></span>|<span data-ttu-id="57ecf-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="57ecf-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57ecf-183">Требования</span><span class="sxs-lookup"><span data-stu-id="57ecf-183">Requirements</span></span>

|<span data-ttu-id="57ecf-184">Требование</span><span class="sxs-lookup"><span data-stu-id="57ecf-184">Requirement</span></span>| <span data-ttu-id="57ecf-185">Значение</span><span class="sxs-lookup"><span data-stu-id="57ecf-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ecf-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="57ecf-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="57ecf-187">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-187">1.1</span></span>|
|[<span data-ttu-id="57ecf-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="57ecf-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ecf-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="57ecf-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="57ecf-190">EventType: String</span></span>

<span data-ttu-id="57ecf-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="57ecf-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="57ecf-192">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-192">Type</span></span>

*   <span data-ttu-id="57ecf-193">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57ecf-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="57ecf-194">Properties:</span></span>

| <span data-ttu-id="57ecf-195">Имя</span><span class="sxs-lookup"><span data-stu-id="57ecf-195">Name</span></span> | <span data-ttu-id="57ecf-196">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-196">Type</span></span> | <span data-ttu-id="57ecf-197">Описание</span><span class="sxs-lookup"><span data-stu-id="57ecf-197">Description</span></span> | <span data-ttu-id="57ecf-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="57ecf-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="57ecf-199">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-199">String</span></span> | <span data-ttu-id="57ecf-200">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="57ecf-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="57ecf-201">1.7</span><span class="sxs-lookup"><span data-stu-id="57ecf-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="57ecf-202">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-202">String</span></span> | <span data-ttu-id="57ecf-203">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="57ecf-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="57ecf-204">1.8</span><span class="sxs-lookup"><span data-stu-id="57ecf-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="57ecf-205">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-205">String</span></span> | <span data-ttu-id="57ecf-206">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="57ecf-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="57ecf-207">1.8</span><span class="sxs-lookup"><span data-stu-id="57ecf-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="57ecf-208">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-208">String</span></span> | <span data-ttu-id="57ecf-209">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="57ecf-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="57ecf-210">1.5</span><span class="sxs-lookup"><span data-stu-id="57ecf-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="57ecf-211">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-211">String</span></span> | <span data-ttu-id="57ecf-212">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="57ecf-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="57ecf-213">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="57ecf-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="57ecf-214">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-214">String</span></span> | <span data-ttu-id="57ecf-215">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="57ecf-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="57ecf-216">1.7</span><span class="sxs-lookup"><span data-stu-id="57ecf-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="57ecf-217">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-217">String</span></span> | <span data-ttu-id="57ecf-218">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="57ecf-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="57ecf-219">1.7</span><span class="sxs-lookup"><span data-stu-id="57ecf-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="57ecf-220">Требования</span><span class="sxs-lookup"><span data-stu-id="57ecf-220">Requirements</span></span>

|<span data-ttu-id="57ecf-221">Требование</span><span class="sxs-lookup"><span data-stu-id="57ecf-221">Requirement</span></span>| <span data-ttu-id="57ecf-222">Значение</span><span class="sxs-lookup"><span data-stu-id="57ecf-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ecf-223">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="57ecf-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="57ecf-224">1.5</span><span class="sxs-lookup"><span data-stu-id="57ecf-224">1.5</span></span> |
|[<span data-ttu-id="57ecf-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="57ecf-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ecf-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-226">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="57ecf-227">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="57ecf-227">SourceProperty: String</span></span>

<span data-ttu-id="57ecf-228">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="57ecf-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="57ecf-229">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-229">Type</span></span>

*   <span data-ttu-id="57ecf-230">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="57ecf-231">Свойства:</span><span class="sxs-lookup"><span data-stu-id="57ecf-231">Properties:</span></span>

|<span data-ttu-id="57ecf-232">Имя</span><span class="sxs-lookup"><span data-stu-id="57ecf-232">Name</span></span>| <span data-ttu-id="57ecf-233">Тип</span><span class="sxs-lookup"><span data-stu-id="57ecf-233">Type</span></span>| <span data-ttu-id="57ecf-234">Описание</span><span class="sxs-lookup"><span data-stu-id="57ecf-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="57ecf-235">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-235">String</span></span>|<span data-ttu-id="57ecf-236">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="57ecf-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="57ecf-237">String</span><span class="sxs-lookup"><span data-stu-id="57ecf-237">String</span></span>|<span data-ttu-id="57ecf-238">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="57ecf-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="57ecf-239">Требования</span><span class="sxs-lookup"><span data-stu-id="57ecf-239">Requirements</span></span>

|<span data-ttu-id="57ecf-240">Требование</span><span class="sxs-lookup"><span data-stu-id="57ecf-240">Requirement</span></span>| <span data-ttu-id="57ecf-241">Значение</span><span class="sxs-lookup"><span data-stu-id="57ecf-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="57ecf-242">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="57ecf-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="57ecf-243">1.1</span><span class="sxs-lookup"><span data-stu-id="57ecf-243">1.1</span></span>|
|[<span data-ttu-id="57ecf-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="57ecf-244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="57ecf-245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="57ecf-245">Compose or Read</span></span>|
