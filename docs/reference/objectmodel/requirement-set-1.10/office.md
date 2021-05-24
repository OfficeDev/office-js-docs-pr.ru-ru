---
title: Office пространства имен — набор требований 1.10
description: Office членов пространства имен, доступных для Outlook надстройки с помощью API почтовых ящиков, установленного 1.10.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e7b7ab9127ebf8ce9b7394d348144fe63b47de6c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592061"
---
# <a name="office-mailbox-requirement-set-110"></a><span data-ttu-id="14e56-103">Office (набор требований к почтовым ящикам 1.10)</span><span class="sxs-lookup"><span data-stu-id="14e56-103">Office (Mailbox requirement set 1.10)</span></span>

<span data-ttu-id="14e56-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="14e56-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="14e56-106">Требования</span><span class="sxs-lookup"><span data-stu-id="14e56-106">Requirements</span></span>

|<span data-ttu-id="14e56-107">Требование</span><span class="sxs-lookup"><span data-stu-id="14e56-107">Requirement</span></span>| <span data-ttu-id="14e56-108">Значение</span><span class="sxs-lookup"><span data-stu-id="14e56-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="14e56-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14e56-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="14e56-110">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-110">1.1</span></span>|
|[<span data-ttu-id="14e56-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14e56-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="14e56-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="14e56-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="14e56-113">Properties</span></span>

| <span data-ttu-id="14e56-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="14e56-114">Property</span></span> | <span data-ttu-id="14e56-115">Режимы</span><span class="sxs-lookup"><span data-stu-id="14e56-115">Modes</span></span> | <span data-ttu-id="14e56-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="14e56-116">Return type</span></span> | <span data-ttu-id="14e56-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="14e56-117">Minimum</span></span><br><span data-ttu-id="14e56-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="14e56-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="14e56-119">контекст</span><span class="sxs-lookup"><span data-stu-id="14e56-119">context</span></span>](office.context.md) | <span data-ttu-id="14e56-120">Создание</span><span class="sxs-lookup"><span data-stu-id="14e56-120">Compose</span></span><br><span data-ttu-id="14e56-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-121">Read</span></span> | [<span data-ttu-id="14e56-122">Context</span><span class="sxs-lookup"><span data-stu-id="14e56-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="14e56-123">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="14e56-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="14e56-124">Enumerations</span></span>

| <span data-ttu-id="14e56-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="14e56-125">Enumeration</span></span> | <span data-ttu-id="14e56-126">Режимы</span><span class="sxs-lookup"><span data-stu-id="14e56-126">Modes</span></span> | <span data-ttu-id="14e56-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="14e56-127">Return type</span></span> | <span data-ttu-id="14e56-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="14e56-128">Minimum</span></span><br><span data-ttu-id="14e56-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="14e56-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="14e56-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="14e56-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="14e56-131">Создание</span><span class="sxs-lookup"><span data-stu-id="14e56-131">Compose</span></span><br><span data-ttu-id="14e56-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-132">Read</span></span> | <span data-ttu-id="14e56-133">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-133">String</span></span> | [<span data-ttu-id="14e56-134">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="14e56-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="14e56-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="14e56-136">Создание</span><span class="sxs-lookup"><span data-stu-id="14e56-136">Compose</span></span><br><span data-ttu-id="14e56-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-137">Read</span></span> | <span data-ttu-id="14e56-138">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-138">String</span></span> | [<span data-ttu-id="14e56-139">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="14e56-140">EventType</span><span class="sxs-lookup"><span data-stu-id="14e56-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="14e56-141">Создание</span><span class="sxs-lookup"><span data-stu-id="14e56-141">Compose</span></span><br><span data-ttu-id="14e56-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-142">Read</span></span> | <span data-ttu-id="14e56-143">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-143">String</span></span> | [<span data-ttu-id="14e56-144">1.5</span><span class="sxs-lookup"><span data-stu-id="14e56-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="14e56-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="14e56-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="14e56-146">Создание</span><span class="sxs-lookup"><span data-stu-id="14e56-146">Compose</span></span><br><span data-ttu-id="14e56-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-147">Read</span></span> | <span data-ttu-id="14e56-148">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-148">String</span></span> | [<span data-ttu-id="14e56-149">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="14e56-150">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="14e56-150">Namespaces</span></span>

<span data-ttu-id="14e56-151">[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="14e56-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="14e56-152">Сведения о переумериях</span><span class="sxs-lookup"><span data-stu-id="14e56-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="14e56-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="14e56-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="14e56-154">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="14e56-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="14e56-155">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-155">Type</span></span>

*   <span data-ttu-id="14e56-156">String</span><span class="sxs-lookup"><span data-stu-id="14e56-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="14e56-157">Свойства</span><span class="sxs-lookup"><span data-stu-id="14e56-157">Properties</span></span>

|<span data-ttu-id="14e56-158">Имя</span><span class="sxs-lookup"><span data-stu-id="14e56-158">Name</span></span>| <span data-ttu-id="14e56-159">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-159">Type</span></span>| <span data-ttu-id="14e56-160">Описание</span><span class="sxs-lookup"><span data-stu-id="14e56-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="14e56-161">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-161">String</span></span>|<span data-ttu-id="14e56-162">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="14e56-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="14e56-163">String</span><span class="sxs-lookup"><span data-stu-id="14e56-163">String</span></span>|<span data-ttu-id="14e56-164">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="14e56-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14e56-165">Требования</span><span class="sxs-lookup"><span data-stu-id="14e56-165">Requirements</span></span>

|<span data-ttu-id="14e56-166">Требование</span><span class="sxs-lookup"><span data-stu-id="14e56-166">Requirement</span></span>| <span data-ttu-id="14e56-167">Значение</span><span class="sxs-lookup"><span data-stu-id="14e56-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="14e56-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14e56-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="14e56-169">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-169">1.1</span></span>|
|[<span data-ttu-id="14e56-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14e56-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="14e56-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="14e56-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="14e56-172">CoercionType: String</span></span>

<span data-ttu-id="14e56-173">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="14e56-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="14e56-174">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-174">Type</span></span>

*   <span data-ttu-id="14e56-175">String</span><span class="sxs-lookup"><span data-stu-id="14e56-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="14e56-176">Свойства</span><span class="sxs-lookup"><span data-stu-id="14e56-176">Properties</span></span>

|<span data-ttu-id="14e56-177">Имя</span><span class="sxs-lookup"><span data-stu-id="14e56-177">Name</span></span>| <span data-ttu-id="14e56-178">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-178">Type</span></span>| <span data-ttu-id="14e56-179">Описание</span><span class="sxs-lookup"><span data-stu-id="14e56-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="14e56-180">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-180">String</span></span>|<span data-ttu-id="14e56-181">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="14e56-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="14e56-182">String</span><span class="sxs-lookup"><span data-stu-id="14e56-182">String</span></span>|<span data-ttu-id="14e56-183">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="14e56-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14e56-184">Требования</span><span class="sxs-lookup"><span data-stu-id="14e56-184">Requirements</span></span>

|<span data-ttu-id="14e56-185">Требование</span><span class="sxs-lookup"><span data-stu-id="14e56-185">Requirement</span></span>| <span data-ttu-id="14e56-186">Значение</span><span class="sxs-lookup"><span data-stu-id="14e56-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="14e56-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14e56-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="14e56-188">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-188">1.1</span></span>|
|[<span data-ttu-id="14e56-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14e56-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="14e56-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="14e56-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="14e56-191">EventType: String</span></span>

<span data-ttu-id="14e56-192">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="14e56-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="14e56-193">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-193">Type</span></span>

*   <span data-ttu-id="14e56-194">String</span><span class="sxs-lookup"><span data-stu-id="14e56-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="14e56-195">Свойства</span><span class="sxs-lookup"><span data-stu-id="14e56-195">Properties</span></span>

| <span data-ttu-id="14e56-196">Имя</span><span class="sxs-lookup"><span data-stu-id="14e56-196">Name</span></span> | <span data-ttu-id="14e56-197">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-197">Type</span></span> | <span data-ttu-id="14e56-198">Описание</span><span class="sxs-lookup"><span data-stu-id="14e56-198">Description</span></span> | <span data-ttu-id="14e56-199">Минимальный набор требований</span><span class="sxs-lookup"><span data-stu-id="14e56-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="14e56-200">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-200">String</span></span> | <span data-ttu-id="14e56-201">Изменилась дата или время выбранной встречи или серии.</span><span class="sxs-lookup"><span data-stu-id="14e56-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="14e56-202">1.7</span><span class="sxs-lookup"><span data-stu-id="14e56-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="14e56-203">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-203">String</span></span> | <span data-ttu-id="14e56-204">Вложение было добавлено или удалено из элемента.</span><span class="sxs-lookup"><span data-stu-id="14e56-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="14e56-205">1.8</span><span class="sxs-lookup"><span data-stu-id="14e56-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="14e56-206">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-206">String</span></span> | <span data-ttu-id="14e56-207">Расположение выбранного назначения изменилось.</span><span class="sxs-lookup"><span data-stu-id="14e56-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="14e56-208">1.8</span><span class="sxs-lookup"><span data-stu-id="14e56-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="14e56-209">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-209">String</span></span> | <span data-ttu-id="14e56-210">Другой элемент Outlook для просмотра при закреплении области задач.</span><span class="sxs-lookup"><span data-stu-id="14e56-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="14e56-211">1.5</span><span class="sxs-lookup"><span data-stu-id="14e56-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="14e56-212">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-212">String</span></span> | <span data-ttu-id="14e56-213">Тема Office на почтовом ящике изменилась.</span><span class="sxs-lookup"><span data-stu-id="14e56-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="14e56-214">1.10</span><span class="sxs-lookup"><span data-stu-id="14e56-214">1.10</span></span> |
|`RecipientsChanged`| <span data-ttu-id="14e56-215">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-215">String</span></span> | <span data-ttu-id="14e56-216">Список получателей выбранного элемента или расположения встречи изменен.</span><span class="sxs-lookup"><span data-stu-id="14e56-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="14e56-217">1.7</span><span class="sxs-lookup"><span data-stu-id="14e56-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="14e56-218">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-218">String</span></span> | <span data-ttu-id="14e56-219">Изменился шаблон повторяемости выбранной серии.</span><span class="sxs-lookup"><span data-stu-id="14e56-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="14e56-220">1.7</span><span class="sxs-lookup"><span data-stu-id="14e56-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14e56-221">Требования</span><span class="sxs-lookup"><span data-stu-id="14e56-221">Requirements</span></span>

|<span data-ttu-id="14e56-222">Требование</span><span class="sxs-lookup"><span data-stu-id="14e56-222">Requirement</span></span>| <span data-ttu-id="14e56-223">Значение</span><span class="sxs-lookup"><span data-stu-id="14e56-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="14e56-224">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14e56-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="14e56-225">1.5</span><span class="sxs-lookup"><span data-stu-id="14e56-225">1.5</span></span> |
|[<span data-ttu-id="14e56-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14e56-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="14e56-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="14e56-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="14e56-228">SourceProperty: String</span></span>

<span data-ttu-id="14e56-229">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="14e56-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="14e56-230">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-230">Type</span></span>

*   <span data-ttu-id="14e56-231">String</span><span class="sxs-lookup"><span data-stu-id="14e56-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="14e56-232">Свойства</span><span class="sxs-lookup"><span data-stu-id="14e56-232">Properties</span></span>

|<span data-ttu-id="14e56-233">Имя</span><span class="sxs-lookup"><span data-stu-id="14e56-233">Name</span></span>| <span data-ttu-id="14e56-234">Тип</span><span class="sxs-lookup"><span data-stu-id="14e56-234">Type</span></span>| <span data-ttu-id="14e56-235">Описание</span><span class="sxs-lookup"><span data-stu-id="14e56-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="14e56-236">Строка</span><span class="sxs-lookup"><span data-stu-id="14e56-236">String</span></span>|<span data-ttu-id="14e56-237">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="14e56-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="14e56-238">String</span><span class="sxs-lookup"><span data-stu-id="14e56-238">String</span></span>|<span data-ttu-id="14e56-239">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="14e56-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14e56-240">Требования</span><span class="sxs-lookup"><span data-stu-id="14e56-240">Requirements</span></span>

|<span data-ttu-id="14e56-241">Требование</span><span class="sxs-lookup"><span data-stu-id="14e56-241">Requirement</span></span>| <span data-ttu-id="14e56-242">Значение</span><span class="sxs-lookup"><span data-stu-id="14e56-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="14e56-243">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14e56-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="14e56-244">1.1</span><span class="sxs-lookup"><span data-stu-id="14e56-244">1.1</span></span>|
|[<span data-ttu-id="14e56-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14e56-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="14e56-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14e56-246">Compose or Read</span></span>|
