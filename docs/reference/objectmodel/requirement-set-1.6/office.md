---
title: Пространство имен Office — набор обязательных элементов 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e15f01db9423a9df38608f18098d2c808f5d944b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814670"
---
# <a name="office"></a><span data-ttu-id="2a61d-102">Office</span><span class="sxs-lookup"><span data-stu-id="2a61d-102">Office</span></span>

<span data-ttu-id="2a61d-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2a61d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2a61d-105">Требования</span><span class="sxs-lookup"><span data-stu-id="2a61d-105">Requirements</span></span>

|<span data-ttu-id="2a61d-106">Требование</span><span class="sxs-lookup"><span data-stu-id="2a61d-106">Requirement</span></span>| <span data-ttu-id="2a61d-107">Значение</span><span class="sxs-lookup"><span data-stu-id="2a61d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a61d-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a61d-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a61d-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-109">1.1</span></span>|
|[<span data-ttu-id="2a61d-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a61d-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a61d-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2a61d-112">Properties</span><span class="sxs-lookup"><span data-stu-id="2a61d-112">Properties</span></span>

| <span data-ttu-id="2a61d-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="2a61d-113">Property</span></span> | <span data-ttu-id="2a61d-114">Способов</span><span class="sxs-lookup"><span data-stu-id="2a61d-114">Modes</span></span> | <span data-ttu-id="2a61d-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="2a61d-115">Return type</span></span> | <span data-ttu-id="2a61d-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="2a61d-116">Minimum</span></span><br><span data-ttu-id="2a61d-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="2a61d-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a61d-118">контекст</span><span class="sxs-lookup"><span data-stu-id="2a61d-118">context</span></span>](office.context.md) | <span data-ttu-id="2a61d-119">Создание</span><span class="sxs-lookup"><span data-stu-id="2a61d-119">Compose</span></span><br><span data-ttu-id="2a61d-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-120">Read</span></span> | [<span data-ttu-id="2a61d-121">Context</span><span class="sxs-lookup"><span data-stu-id="2a61d-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="2a61d-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2a61d-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="2a61d-123">Enumerations</span></span>

| <span data-ttu-id="2a61d-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="2a61d-124">Enumeration</span></span> | <span data-ttu-id="2a61d-125">Способов</span><span class="sxs-lookup"><span data-stu-id="2a61d-125">Modes</span></span> | <span data-ttu-id="2a61d-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="2a61d-126">Return type</span></span> | <span data-ttu-id="2a61d-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="2a61d-127">Minimum</span></span><br><span data-ttu-id="2a61d-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="2a61d-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2a61d-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2a61d-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2a61d-130">Создание</span><span class="sxs-lookup"><span data-stu-id="2a61d-130">Compose</span></span><br><span data-ttu-id="2a61d-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-131">Read</span></span> | <span data-ttu-id="2a61d-132">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-132">String</span></span> | [<span data-ttu-id="2a61d-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a61d-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2a61d-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2a61d-135">Создание</span><span class="sxs-lookup"><span data-stu-id="2a61d-135">Compose</span></span><br><span data-ttu-id="2a61d-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-136">Read</span></span> | <span data-ttu-id="2a61d-137">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-137">String</span></span> | [<span data-ttu-id="2a61d-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2a61d-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2a61d-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2a61d-140">Создание</span><span class="sxs-lookup"><span data-stu-id="2a61d-140">Compose</span></span><br><span data-ttu-id="2a61d-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-141">Read</span></span> | <span data-ttu-id="2a61d-142">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-142">String</span></span> | [<span data-ttu-id="2a61d-143">1,5</span><span class="sxs-lookup"><span data-stu-id="2a61d-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2a61d-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2a61d-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2a61d-145">Создание</span><span class="sxs-lookup"><span data-stu-id="2a61d-145">Compose</span></span><br><span data-ttu-id="2a61d-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-146">Read</span></span> | <span data-ttu-id="2a61d-147">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-147">String</span></span> | [<span data-ttu-id="2a61d-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2a61d-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="2a61d-149">Namespaces</span></span>

<span data-ttu-id="2a61d-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="2a61d-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2a61d-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="2a61d-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2a61d-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="2a61d-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2a61d-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="2a61d-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2a61d-154">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-154">Type</span></span>

*   <span data-ttu-id="2a61d-155">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a61d-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a61d-156">Properties:</span></span>

|<span data-ttu-id="2a61d-157">Имя</span><span class="sxs-lookup"><span data-stu-id="2a61d-157">Name</span></span>| <span data-ttu-id="2a61d-158">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-158">Type</span></span>| <span data-ttu-id="2a61d-159">Описание</span><span class="sxs-lookup"><span data-stu-id="2a61d-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2a61d-160">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-160">String</span></span>|<span data-ttu-id="2a61d-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="2a61d-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2a61d-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="2a61d-162">String</span></span>|<span data-ttu-id="2a61d-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="2a61d-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a61d-164">Требования</span><span class="sxs-lookup"><span data-stu-id="2a61d-164">Requirements</span></span>

|<span data-ttu-id="2a61d-165">Требование</span><span class="sxs-lookup"><span data-stu-id="2a61d-165">Requirement</span></span>| <span data-ttu-id="2a61d-166">Значение</span><span class="sxs-lookup"><span data-stu-id="2a61d-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a61d-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a61d-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a61d-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-168">1.1</span></span>|
|[<span data-ttu-id="2a61d-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a61d-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a61d-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2a61d-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="2a61d-171">CoercionType: String</span></span>

<span data-ttu-id="2a61d-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="2a61d-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a61d-173">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-173">Type</span></span>

*   <span data-ttu-id="2a61d-174">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a61d-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a61d-175">Properties:</span></span>

|<span data-ttu-id="2a61d-176">Имя</span><span class="sxs-lookup"><span data-stu-id="2a61d-176">Name</span></span>| <span data-ttu-id="2a61d-177">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-177">Type</span></span>| <span data-ttu-id="2a61d-178">Описание</span><span class="sxs-lookup"><span data-stu-id="2a61d-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2a61d-179">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-179">String</span></span>|<span data-ttu-id="2a61d-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="2a61d-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2a61d-181">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-181">String</span></span>|<span data-ttu-id="2a61d-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="2a61d-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a61d-183">Требования</span><span class="sxs-lookup"><span data-stu-id="2a61d-183">Requirements</span></span>

|<span data-ttu-id="2a61d-184">Требование</span><span class="sxs-lookup"><span data-stu-id="2a61d-184">Requirement</span></span>| <span data-ttu-id="2a61d-185">Значение</span><span class="sxs-lookup"><span data-stu-id="2a61d-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a61d-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a61d-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a61d-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-187">1.1</span></span>|
|[<span data-ttu-id="2a61d-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a61d-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a61d-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2a61d-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="2a61d-190">EventType: String</span></span>

<span data-ttu-id="2a61d-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="2a61d-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2a61d-192">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-192">Type</span></span>

*   <span data-ttu-id="2a61d-193">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a61d-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a61d-194">Properties:</span></span>

| <span data-ttu-id="2a61d-195">Имя</span><span class="sxs-lookup"><span data-stu-id="2a61d-195">Name</span></span> | <span data-ttu-id="2a61d-196">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-196">Type</span></span> | <span data-ttu-id="2a61d-197">Описание</span><span class="sxs-lookup"><span data-stu-id="2a61d-197">Description</span></span> | <span data-ttu-id="2a61d-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="2a61d-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="2a61d-199">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-199">String</span></span> | <span data-ttu-id="2a61d-200">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="2a61d-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2a61d-201">1.5</span><span class="sxs-lookup"><span data-stu-id="2a61d-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2a61d-202">Требования</span><span class="sxs-lookup"><span data-stu-id="2a61d-202">Requirements</span></span>

|<span data-ttu-id="2a61d-203">Требование</span><span class="sxs-lookup"><span data-stu-id="2a61d-203">Requirement</span></span>| <span data-ttu-id="2a61d-204">Значение</span><span class="sxs-lookup"><span data-stu-id="2a61d-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a61d-205">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2a61d-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a61d-206">1.5</span><span class="sxs-lookup"><span data-stu-id="2a61d-206">1.5</span></span> |
|[<span data-ttu-id="2a61d-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a61d-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a61d-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2a61d-209">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="2a61d-209">SourceProperty: String</span></span>

<span data-ttu-id="2a61d-210">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="2a61d-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2a61d-211">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-211">Type</span></span>

*   <span data-ttu-id="2a61d-212">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2a61d-213">Свойства:</span><span class="sxs-lookup"><span data-stu-id="2a61d-213">Properties:</span></span>

|<span data-ttu-id="2a61d-214">Имя</span><span class="sxs-lookup"><span data-stu-id="2a61d-214">Name</span></span>| <span data-ttu-id="2a61d-215">Тип</span><span class="sxs-lookup"><span data-stu-id="2a61d-215">Type</span></span>| <span data-ttu-id="2a61d-216">Описание</span><span class="sxs-lookup"><span data-stu-id="2a61d-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2a61d-217">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-217">String</span></span>|<span data-ttu-id="2a61d-218">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="2a61d-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2a61d-219">String</span><span class="sxs-lookup"><span data-stu-id="2a61d-219">String</span></span>|<span data-ttu-id="2a61d-220">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="2a61d-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2a61d-221">Требования</span><span class="sxs-lookup"><span data-stu-id="2a61d-221">Requirements</span></span>

|<span data-ttu-id="2a61d-222">Требование</span><span class="sxs-lookup"><span data-stu-id="2a61d-222">Requirement</span></span>| <span data-ttu-id="2a61d-223">Значение</span><span class="sxs-lookup"><span data-stu-id="2a61d-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="2a61d-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2a61d-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2a61d-225">1.1</span><span class="sxs-lookup"><span data-stu-id="2a61d-225">1.1</span></span>|
|[<span data-ttu-id="2a61d-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2a61d-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2a61d-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2a61d-227">Compose or Read</span></span>|
