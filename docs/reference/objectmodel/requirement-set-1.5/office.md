---
title: Пространство имен Office — набор обязательных элементов 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165386"
---
# <a name="office"></a><span data-ttu-id="4bcae-102">Office</span><span class="sxs-lookup"><span data-stu-id="4bcae-102">Office</span></span>

<span data-ttu-id="4bcae-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4bcae-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4bcae-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bcae-105">Requirements</span></span>

|<span data-ttu-id="4bcae-106">Требование</span><span class="sxs-lookup"><span data-stu-id="4bcae-106">Requirement</span></span>| <span data-ttu-id="4bcae-107">Значение</span><span class="sxs-lookup"><span data-stu-id="4bcae-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bcae-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bcae-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4bcae-109">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-109">1.1</span></span>|
|[<span data-ttu-id="4bcae-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bcae-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4bcae-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4bcae-112">Properties</span><span class="sxs-lookup"><span data-stu-id="4bcae-112">Properties</span></span>

| <span data-ttu-id="4bcae-113">Свойство</span><span class="sxs-lookup"><span data-stu-id="4bcae-113">Property</span></span> | <span data-ttu-id="4bcae-114">Способов</span><span class="sxs-lookup"><span data-stu-id="4bcae-114">Modes</span></span> | <span data-ttu-id="4bcae-115">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="4bcae-115">Return type</span></span> | <span data-ttu-id="4bcae-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bcae-116">Minimum</span></span><br><span data-ttu-id="4bcae-117">набор требований</span><span class="sxs-lookup"><span data-stu-id="4bcae-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4bcae-118">контекст</span><span class="sxs-lookup"><span data-stu-id="4bcae-118">context</span></span>](office.context.md) | <span data-ttu-id="4bcae-119">Создание</span><span class="sxs-lookup"><span data-stu-id="4bcae-119">Compose</span></span><br><span data-ttu-id="4bcae-120">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-120">Read</span></span> | [<span data-ttu-id="4bcae-121">Context</span><span class="sxs-lookup"><span data-stu-id="4bcae-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="4bcae-122">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="4bcae-123">Перечисления</span><span class="sxs-lookup"><span data-stu-id="4bcae-123">Enumerations</span></span>

| <span data-ttu-id="4bcae-124">Перечисление</span><span class="sxs-lookup"><span data-stu-id="4bcae-124">Enumeration</span></span> | <span data-ttu-id="4bcae-125">Способов</span><span class="sxs-lookup"><span data-stu-id="4bcae-125">Modes</span></span> | <span data-ttu-id="4bcae-126">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="4bcae-126">Return type</span></span> | <span data-ttu-id="4bcae-127">Минимальные</span><span class="sxs-lookup"><span data-stu-id="4bcae-127">Minimum</span></span><br><span data-ttu-id="4bcae-128">набор требований</span><span class="sxs-lookup"><span data-stu-id="4bcae-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4bcae-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4bcae-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4bcae-130">Создание</span><span class="sxs-lookup"><span data-stu-id="4bcae-130">Compose</span></span><br><span data-ttu-id="4bcae-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-131">Read</span></span> | <span data-ttu-id="4bcae-132">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-132">String</span></span> | [<span data-ttu-id="4bcae-133">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4bcae-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4bcae-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4bcae-135">Создание</span><span class="sxs-lookup"><span data-stu-id="4bcae-135">Compose</span></span><br><span data-ttu-id="4bcae-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-136">Read</span></span> | <span data-ttu-id="4bcae-137">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-137">String</span></span> | [<span data-ttu-id="4bcae-138">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4bcae-139">EventType</span><span class="sxs-lookup"><span data-stu-id="4bcae-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4bcae-140">Создание</span><span class="sxs-lookup"><span data-stu-id="4bcae-140">Compose</span></span><br><span data-ttu-id="4bcae-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-141">Read</span></span> | <span data-ttu-id="4bcae-142">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-142">String</span></span> | [<span data-ttu-id="4bcae-143">1,5</span><span class="sxs-lookup"><span data-stu-id="4bcae-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="4bcae-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4bcae-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4bcae-145">Создание</span><span class="sxs-lookup"><span data-stu-id="4bcae-145">Compose</span></span><br><span data-ttu-id="4bcae-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-146">Read</span></span> | <span data-ttu-id="4bcae-147">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-147">String</span></span> | [<span data-ttu-id="4bcae-148">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="4bcae-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="4bcae-149">Namespaces</span></span>

<span data-ttu-id="4bcae-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="4bcae-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4bcae-151">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="4bcae-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4bcae-152">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="4bcae-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="4bcae-153">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="4bcae-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4bcae-154">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-154">Type</span></span>

*   <span data-ttu-id="4bcae-155">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bcae-156">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4bcae-156">Properties:</span></span>

|<span data-ttu-id="4bcae-157">Имя</span><span class="sxs-lookup"><span data-stu-id="4bcae-157">Name</span></span>| <span data-ttu-id="4bcae-158">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-158">Type</span></span>| <span data-ttu-id="4bcae-159">Описание</span><span class="sxs-lookup"><span data-stu-id="4bcae-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4bcae-160">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-160">String</span></span>|<span data-ttu-id="4bcae-161">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="4bcae-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4bcae-162">Для указания</span><span class="sxs-lookup"><span data-stu-id="4bcae-162">String</span></span>|<span data-ttu-id="4bcae-163">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="4bcae-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bcae-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bcae-164">Requirements</span></span>

|<span data-ttu-id="4bcae-165">Требование</span><span class="sxs-lookup"><span data-stu-id="4bcae-165">Requirement</span></span>| <span data-ttu-id="4bcae-166">Значение</span><span class="sxs-lookup"><span data-stu-id="4bcae-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bcae-167">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bcae-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4bcae-168">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-168">1.1</span></span>|
|[<span data-ttu-id="4bcae-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bcae-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4bcae-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4bcae-171">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="4bcae-171">CoercionType: String</span></span>

<span data-ttu-id="4bcae-172">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="4bcae-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4bcae-173">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-173">Type</span></span>

*   <span data-ttu-id="4bcae-174">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bcae-175">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4bcae-175">Properties:</span></span>

|<span data-ttu-id="4bcae-176">Имя</span><span class="sxs-lookup"><span data-stu-id="4bcae-176">Name</span></span>| <span data-ttu-id="4bcae-177">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-177">Type</span></span>| <span data-ttu-id="4bcae-178">Описание</span><span class="sxs-lookup"><span data-stu-id="4bcae-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4bcae-179">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-179">String</span></span>|<span data-ttu-id="4bcae-180">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="4bcae-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4bcae-181">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-181">String</span></span>|<span data-ttu-id="4bcae-182">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="4bcae-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bcae-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bcae-183">Requirements</span></span>

|<span data-ttu-id="4bcae-184">Требование</span><span class="sxs-lookup"><span data-stu-id="4bcae-184">Requirement</span></span>| <span data-ttu-id="4bcae-185">Значение</span><span class="sxs-lookup"><span data-stu-id="4bcae-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bcae-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bcae-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4bcae-187">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-187">1.1</span></span>|
|[<span data-ttu-id="4bcae-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bcae-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4bcae-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4bcae-190">EventType: строка</span><span class="sxs-lookup"><span data-stu-id="4bcae-190">EventType: String</span></span>

<span data-ttu-id="4bcae-191">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="4bcae-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4bcae-192">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-192">Type</span></span>

*   <span data-ttu-id="4bcae-193">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bcae-194">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4bcae-194">Properties:</span></span>

| <span data-ttu-id="4bcae-195">Имя</span><span class="sxs-lookup"><span data-stu-id="4bcae-195">Name</span></span> | <span data-ttu-id="4bcae-196">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-196">Type</span></span> | <span data-ttu-id="4bcae-197">Описание</span><span class="sxs-lookup"><span data-stu-id="4bcae-197">Description</span></span> | <span data-ttu-id="4bcae-198">Набор минимальных требований</span><span class="sxs-lookup"><span data-stu-id="4bcae-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="4bcae-199">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-199">String</span></span> | <span data-ttu-id="4bcae-200">Для просмотра выбран другой элемент Outlook, когда область задач закреплена.</span><span class="sxs-lookup"><span data-stu-id="4bcae-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4bcae-201">1.5</span><span class="sxs-lookup"><span data-stu-id="4bcae-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4bcae-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bcae-202">Requirements</span></span>

|<span data-ttu-id="4bcae-203">Требование</span><span class="sxs-lookup"><span data-stu-id="4bcae-203">Requirement</span></span>| <span data-ttu-id="4bcae-204">Значение</span><span class="sxs-lookup"><span data-stu-id="4bcae-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bcae-205">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="4bcae-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4bcae-206">1.5</span><span class="sxs-lookup"><span data-stu-id="4bcae-206">1.5</span></span> |
|[<span data-ttu-id="4bcae-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bcae-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4bcae-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4bcae-209">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="4bcae-209">SourceProperty: String</span></span>

<span data-ttu-id="4bcae-210">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="4bcae-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4bcae-211">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-211">Type</span></span>

*   <span data-ttu-id="4bcae-212">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4bcae-213">Свойства:</span><span class="sxs-lookup"><span data-stu-id="4bcae-213">Properties:</span></span>

|<span data-ttu-id="4bcae-214">Имя</span><span class="sxs-lookup"><span data-stu-id="4bcae-214">Name</span></span>| <span data-ttu-id="4bcae-215">Тип</span><span class="sxs-lookup"><span data-stu-id="4bcae-215">Type</span></span>| <span data-ttu-id="4bcae-216">Описание</span><span class="sxs-lookup"><span data-stu-id="4bcae-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4bcae-217">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-217">String</span></span>|<span data-ttu-id="4bcae-218">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bcae-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4bcae-219">String</span><span class="sxs-lookup"><span data-stu-id="4bcae-219">String</span></span>|<span data-ttu-id="4bcae-220">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="4bcae-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4bcae-221">Requirements</span><span class="sxs-lookup"><span data-stu-id="4bcae-221">Requirements</span></span>

|<span data-ttu-id="4bcae-222">Требование</span><span class="sxs-lookup"><span data-stu-id="4bcae-222">Requirement</span></span>| <span data-ttu-id="4bcae-223">Значение</span><span class="sxs-lookup"><span data-stu-id="4bcae-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="4bcae-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4bcae-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4bcae-225">1.1</span><span class="sxs-lookup"><span data-stu-id="4bcae-225">1.1</span></span>|
|[<span data-ttu-id="4bcae-226">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4bcae-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4bcae-227">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4bcae-227">Compose or Read</span></span>|
