---
title: Пространство имен Office — Предварительная версия набора требований
description: Элементы пространства имен Office, доступные для надстроек Outlook с использованием набора обязательных элементов API почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 1e0f932106df462c7cd172327082992f6e4d9a58
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431124"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="a5e50-103">Office (набор требований предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="a5e50-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="a5e50-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a5e50-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a5e50-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5e50-106">Requirements</span></span>

|<span data-ttu-id="a5e50-107">Требование</span><span class="sxs-lookup"><span data-stu-id="a5e50-107">Requirement</span></span>| <span data-ttu-id="a5e50-108">Значение</span><span class="sxs-lookup"><span data-stu-id="a5e50-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5e50-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5e50-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a5e50-110">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-110">1.1</span></span>|
|[<span data-ttu-id="a5e50-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e50-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a5e50-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a5e50-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="a5e50-113">Properties</span></span>

| <span data-ttu-id="a5e50-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="a5e50-114">Property</span></span> | <span data-ttu-id="a5e50-115">Способов</span><span class="sxs-lookup"><span data-stu-id="a5e50-115">Modes</span></span> | <span data-ttu-id="a5e50-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a5e50-116">Return type</span></span> | <span data-ttu-id="a5e50-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="a5e50-117">Minimum</span></span><br><span data-ttu-id="a5e50-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="a5e50-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a5e50-119">контекст</span><span class="sxs-lookup"><span data-stu-id="a5e50-119">context</span></span>](office.context.md) | <span data-ttu-id="a5e50-120">Создание</span><span class="sxs-lookup"><span data-stu-id="a5e50-120">Compose</span></span><br><span data-ttu-id="a5e50-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-121">Read</span></span> | [<span data-ttu-id="a5e50-122">Context</span><span class="sxs-lookup"><span data-stu-id="a5e50-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="a5e50-123">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="a5e50-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="a5e50-124">Enumerations</span></span>

| <span data-ttu-id="a5e50-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="a5e50-125">Enumeration</span></span> | <span data-ttu-id="a5e50-126">Способов</span><span class="sxs-lookup"><span data-stu-id="a5e50-126">Modes</span></span> | <span data-ttu-id="a5e50-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a5e50-127">Return type</span></span> | <span data-ttu-id="a5e50-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="a5e50-128">Minimum</span></span><br><span data-ttu-id="a5e50-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="a5e50-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a5e50-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a5e50-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a5e50-131">Создание</span><span class="sxs-lookup"><span data-stu-id="a5e50-131">Compose</span></span><br><span data-ttu-id="a5e50-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-132">Read</span></span> | <span data-ttu-id="a5e50-133">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-133">String</span></span> | [<span data-ttu-id="a5e50-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a5e50-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a5e50-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a5e50-136">Создание</span><span class="sxs-lookup"><span data-stu-id="a5e50-136">Compose</span></span><br><span data-ttu-id="a5e50-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-137">Read</span></span> | <span data-ttu-id="a5e50-138">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-138">String</span></span> | [<span data-ttu-id="a5e50-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a5e50-140">EventType</span><span class="sxs-lookup"><span data-stu-id="a5e50-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a5e50-141">Создание</span><span class="sxs-lookup"><span data-stu-id="a5e50-141">Compose</span></span><br><span data-ttu-id="a5e50-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-142">Read</span></span> | <span data-ttu-id="a5e50-143">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-143">String</span></span> | [<span data-ttu-id="a5e50-144">1,5</span><span class="sxs-lookup"><span data-stu-id="a5e50-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="a5e50-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a5e50-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a5e50-146">Создание</span><span class="sxs-lookup"><span data-stu-id="a5e50-146">Compose</span></span><br><span data-ttu-id="a5e50-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-147">Read</span></span> | <span data-ttu-id="a5e50-148">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-148">String</span></span> | [<span data-ttu-id="a5e50-149">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="a5e50-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="a5e50-150">Namespaces</span></span>

<span data-ttu-id="a5e50-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="a5e50-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="a5e50-152">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="a5e50-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a5e50-153">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="a5e50-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="a5e50-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="a5e50-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a5e50-155">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-155">Type</span></span>

*   <span data-ttu-id="a5e50-156">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5e50-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5e50-157">Properties:</span></span>

|<span data-ttu-id="a5e50-158">Имя</span><span class="sxs-lookup"><span data-stu-id="a5e50-158">Name</span></span>| <span data-ttu-id="a5e50-159">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-159">Type</span></span>| <span data-ttu-id="a5e50-160">Описание</span><span class="sxs-lookup"><span data-stu-id="a5e50-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a5e50-161">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-161">String</span></span>|<span data-ttu-id="a5e50-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="a5e50-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a5e50-163">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-163">String</span></span>|<span data-ttu-id="a5e50-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="a5e50-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5e50-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5e50-165">Requirements</span></span>

|<span data-ttu-id="a5e50-166">Требование</span><span class="sxs-lookup"><span data-stu-id="a5e50-166">Requirement</span></span>| <span data-ttu-id="a5e50-167">Значение</span><span class="sxs-lookup"><span data-stu-id="a5e50-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5e50-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5e50-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a5e50-169">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-169">1.1</span></span>|
|[<span data-ttu-id="a5e50-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e50-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a5e50-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a5e50-172">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="a5e50-172">CoercionType: String</span></span>

<span data-ttu-id="a5e50-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5e50-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5e50-174">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-174">Type</span></span>

*   <span data-ttu-id="a5e50-175">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5e50-176">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5e50-176">Properties:</span></span>

|<span data-ttu-id="a5e50-177">Имя</span><span class="sxs-lookup"><span data-stu-id="a5e50-177">Name</span></span>| <span data-ttu-id="a5e50-178">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-178">Type</span></span>| <span data-ttu-id="a5e50-179">Описание</span><span class="sxs-lookup"><span data-stu-id="a5e50-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a5e50-180">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-180">String</span></span>|<span data-ttu-id="a5e50-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="a5e50-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a5e50-182">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-182">String</span></span>|<span data-ttu-id="a5e50-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="a5e50-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5e50-184">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5e50-184">Requirements</span></span>

|<span data-ttu-id="a5e50-185">Требование</span><span class="sxs-lookup"><span data-stu-id="a5e50-185">Requirement</span></span>| <span data-ttu-id="a5e50-186">Значение</span><span class="sxs-lookup"><span data-stu-id="a5e50-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5e50-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5e50-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a5e50-188">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-188">1.1</span></span>|
|[<span data-ttu-id="a5e50-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e50-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a5e50-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="a5e50-191">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="a5e50-191">EventType: String</span></span>

<span data-ttu-id="a5e50-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="a5e50-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a5e50-193">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-193">Type</span></span>

*   <span data-ttu-id="a5e50-194">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5e50-195">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5e50-195">Properties:</span></span>

| <span data-ttu-id="a5e50-196">Имя</span><span class="sxs-lookup"><span data-stu-id="a5e50-196">Name</span></span> | <span data-ttu-id="a5e50-197">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-197">Type</span></span> | <span data-ttu-id="a5e50-198">Описание</span><span class="sxs-lookup"><span data-stu-id="a5e50-198">Description</span></span> | <span data-ttu-id="a5e50-199">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="a5e50-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="a5e50-200">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-200">String</span></span> | <span data-ttu-id="a5e50-201">Дата или время выбранной встречи или ряда изменились.</span><span class="sxs-lookup"><span data-stu-id="a5e50-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a5e50-202">1.7</span><span class="sxs-lookup"><span data-stu-id="a5e50-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a5e50-203">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-203">String</span></span> | <span data-ttu-id="a5e50-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="a5e50-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a5e50-205">1.8</span><span class="sxs-lookup"><span data-stu-id="a5e50-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="a5e50-206">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-206">String</span></span> | <span data-ttu-id="a5e50-207">Расположение выбранной встречи изменилось.</span><span class="sxs-lookup"><span data-stu-id="a5e50-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="a5e50-208">1.8</span><span class="sxs-lookup"><span data-stu-id="a5e50-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="a5e50-209">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-209">String</span></span> | <span data-ttu-id="a5e50-210">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="a5e50-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a5e50-211">1.5</span><span class="sxs-lookup"><span data-stu-id="a5e50-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a5e50-212">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-212">String</span></span> | <span data-ttu-id="a5e50-213">Тема Office в почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="a5e50-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a5e50-214">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="a5e50-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a5e50-215">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-215">String</span></span> | <span data-ttu-id="a5e50-216">Список получателей выбранного элемента или места встречи изменился.</span><span class="sxs-lookup"><span data-stu-id="a5e50-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a5e50-217">1.7</span><span class="sxs-lookup"><span data-stu-id="a5e50-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a5e50-218">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-218">String</span></span> | <span data-ttu-id="a5e50-219">Шаблон повторения выбранного ряда изменился.</span><span class="sxs-lookup"><span data-stu-id="a5e50-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a5e50-220">1.7</span><span class="sxs-lookup"><span data-stu-id="a5e50-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a5e50-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5e50-221">Requirements</span></span>

|<span data-ttu-id="a5e50-222">Требование</span><span class="sxs-lookup"><span data-stu-id="a5e50-222">Requirement</span></span>| <span data-ttu-id="a5e50-223">Значение</span><span class="sxs-lookup"><span data-stu-id="a5e50-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5e50-224">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a5e50-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a5e50-225">1.5</span><span class="sxs-lookup"><span data-stu-id="a5e50-225">1.5</span></span> |
|[<span data-ttu-id="a5e50-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e50-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a5e50-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a5e50-228">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="a5e50-228">SourceProperty: String</span></span>

<span data-ttu-id="a5e50-229">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="a5e50-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a5e50-230">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-230">Type</span></span>

*   <span data-ttu-id="a5e50-231">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a5e50-232">Свойства:</span><span class="sxs-lookup"><span data-stu-id="a5e50-232">Properties:</span></span>

|<span data-ttu-id="a5e50-233">Имя</span><span class="sxs-lookup"><span data-stu-id="a5e50-233">Name</span></span>| <span data-ttu-id="a5e50-234">Тип</span><span class="sxs-lookup"><span data-stu-id="a5e50-234">Type</span></span>| <span data-ttu-id="a5e50-235">Описание</span><span class="sxs-lookup"><span data-stu-id="a5e50-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a5e50-236">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-236">String</span></span>|<span data-ttu-id="a5e50-237">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5e50-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a5e50-238">String</span><span class="sxs-lookup"><span data-stu-id="a5e50-238">String</span></span>|<span data-ttu-id="a5e50-239">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="a5e50-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a5e50-240">Requirements</span><span class="sxs-lookup"><span data-stu-id="a5e50-240">Requirements</span></span>

|<span data-ttu-id="a5e50-241">Требование</span><span class="sxs-lookup"><span data-stu-id="a5e50-241">Requirement</span></span>| <span data-ttu-id="a5e50-242">Значение</span><span class="sxs-lookup"><span data-stu-id="a5e50-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="a5e50-243">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a5e50-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a5e50-244">1.1</span><span class="sxs-lookup"><span data-stu-id="a5e50-244">1.1</span></span>|
|[<span data-ttu-id="a5e50-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a5e50-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a5e50-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a5e50-246">Compose or Read</span></span>|
