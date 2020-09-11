---
title: Пространство имен Office — набор обязательных элементов 1,2
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0dfde315cb71642a995b4c07a1966d3dee3c0d50
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431306"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="c07bc-103">Office (набор требований для почтового ящика 1,2)</span><span class="sxs-lookup"><span data-stu-id="c07bc-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="c07bc-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c07bc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c07bc-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c07bc-106">Requirements</span></span>

|<span data-ttu-id="c07bc-107">Требование</span><span class="sxs-lookup"><span data-stu-id="c07bc-107">Requirement</span></span>| <span data-ttu-id="c07bc-108">Значение</span><span class="sxs-lookup"><span data-stu-id="c07bc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c07bc-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c07bc-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c07bc-110">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-110">1.1</span></span>|
|[<span data-ttu-id="c07bc-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c07bc-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c07bc-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c07bc-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="c07bc-113">Properties</span></span>

| <span data-ttu-id="c07bc-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="c07bc-114">Property</span></span> | <span data-ttu-id="c07bc-115">Способов</span><span class="sxs-lookup"><span data-stu-id="c07bc-115">Modes</span></span> | <span data-ttu-id="c07bc-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="c07bc-116">Return type</span></span> | <span data-ttu-id="c07bc-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="c07bc-117">Minimum</span></span><br><span data-ttu-id="c07bc-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="c07bc-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c07bc-119">контекст</span><span class="sxs-lookup"><span data-stu-id="c07bc-119">context</span></span>](office.context.md) | <span data-ttu-id="c07bc-120">Создание</span><span class="sxs-lookup"><span data-stu-id="c07bc-120">Compose</span></span><br><span data-ttu-id="c07bc-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-121">Read</span></span> | [<span data-ttu-id="c07bc-122">Context</span><span class="sxs-lookup"><span data-stu-id="c07bc-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="c07bc-123">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="c07bc-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="c07bc-124">Enumerations</span></span>

| <span data-ttu-id="c07bc-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="c07bc-125">Enumeration</span></span> | <span data-ttu-id="c07bc-126">Способов</span><span class="sxs-lookup"><span data-stu-id="c07bc-126">Modes</span></span> | <span data-ttu-id="c07bc-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="c07bc-127">Return type</span></span> | <span data-ttu-id="c07bc-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="c07bc-128">Minimum</span></span><br><span data-ttu-id="c07bc-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="c07bc-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c07bc-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c07bc-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c07bc-131">Создание</span><span class="sxs-lookup"><span data-stu-id="c07bc-131">Compose</span></span><br><span data-ttu-id="c07bc-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-132">Read</span></span> | <span data-ttu-id="c07bc-133">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-133">String</span></span> | [<span data-ttu-id="c07bc-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c07bc-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c07bc-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c07bc-136">Создание</span><span class="sxs-lookup"><span data-stu-id="c07bc-136">Compose</span></span><br><span data-ttu-id="c07bc-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-137">Read</span></span> | <span data-ttu-id="c07bc-138">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-138">String</span></span> | [<span data-ttu-id="c07bc-139">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c07bc-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c07bc-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c07bc-141">Создание</span><span class="sxs-lookup"><span data-stu-id="c07bc-141">Compose</span></span><br><span data-ttu-id="c07bc-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-142">Read</span></span> | <span data-ttu-id="c07bc-143">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-143">String</span></span> | [<span data-ttu-id="c07bc-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="c07bc-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c07bc-145">Namespaces</span></span>

<span data-ttu-id="c07bc-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="c07bc-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="c07bc-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="c07bc-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c07bc-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="c07bc-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="c07bc-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c07bc-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c07bc-150">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-150">Type</span></span>

*   <span data-ttu-id="c07bc-151">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c07bc-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c07bc-152">Properties:</span></span>

|<span data-ttu-id="c07bc-153">Имя</span><span class="sxs-lookup"><span data-stu-id="c07bc-153">Name</span></span>| <span data-ttu-id="c07bc-154">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-154">Type</span></span>| <span data-ttu-id="c07bc-155">Описание</span><span class="sxs-lookup"><span data-stu-id="c07bc-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c07bc-156">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-156">String</span></span>|<span data-ttu-id="c07bc-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c07bc-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c07bc-158">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-158">String</span></span>|<span data-ttu-id="c07bc-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="c07bc-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c07bc-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="c07bc-160">Requirements</span></span>

|<span data-ttu-id="c07bc-161">Требование</span><span class="sxs-lookup"><span data-stu-id="c07bc-161">Requirement</span></span>| <span data-ttu-id="c07bc-162">Значение</span><span class="sxs-lookup"><span data-stu-id="c07bc-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="c07bc-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c07bc-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c07bc-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-164">1.1</span></span>|
|[<span data-ttu-id="c07bc-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c07bc-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c07bc-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="c07bc-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="c07bc-167">CoercionType: String</span></span>

<span data-ttu-id="c07bc-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c07bc-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c07bc-169">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-169">Type</span></span>

*   <span data-ttu-id="c07bc-170">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c07bc-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c07bc-171">Properties:</span></span>

|<span data-ttu-id="c07bc-172">Имя</span><span class="sxs-lookup"><span data-stu-id="c07bc-172">Name</span></span>| <span data-ttu-id="c07bc-173">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-173">Type</span></span>| <span data-ttu-id="c07bc-174">Описание</span><span class="sxs-lookup"><span data-stu-id="c07bc-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c07bc-175">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-175">String</span></span>|<span data-ttu-id="c07bc-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c07bc-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c07bc-177">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-177">String</span></span>|<span data-ttu-id="c07bc-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c07bc-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c07bc-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="c07bc-179">Requirements</span></span>

|<span data-ttu-id="c07bc-180">Требование</span><span class="sxs-lookup"><span data-stu-id="c07bc-180">Requirement</span></span>| <span data-ttu-id="c07bc-181">Значение</span><span class="sxs-lookup"><span data-stu-id="c07bc-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c07bc-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c07bc-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c07bc-183">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-183">1.1</span></span>|
|[<span data-ttu-id="c07bc-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c07bc-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c07bc-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="c07bc-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="c07bc-186">SourceProperty: String</span></span>

<span data-ttu-id="c07bc-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c07bc-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c07bc-188">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-188">Type</span></span>

*   <span data-ttu-id="c07bc-189">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c07bc-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c07bc-190">Properties:</span></span>

|<span data-ttu-id="c07bc-191">Имя</span><span class="sxs-lookup"><span data-stu-id="c07bc-191">Name</span></span>| <span data-ttu-id="c07bc-192">Тип</span><span class="sxs-lookup"><span data-stu-id="c07bc-192">Type</span></span>| <span data-ttu-id="c07bc-193">Описание</span><span class="sxs-lookup"><span data-stu-id="c07bc-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c07bc-194">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-194">String</span></span>|<span data-ttu-id="c07bc-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c07bc-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c07bc-196">String</span><span class="sxs-lookup"><span data-stu-id="c07bc-196">String</span></span>|<span data-ttu-id="c07bc-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c07bc-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c07bc-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="c07bc-198">Requirements</span></span>

|<span data-ttu-id="c07bc-199">Требование</span><span class="sxs-lookup"><span data-stu-id="c07bc-199">Requirement</span></span>| <span data-ttu-id="c07bc-200">Значение</span><span class="sxs-lookup"><span data-stu-id="c07bc-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="c07bc-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c07bc-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c07bc-202">1.1</span><span class="sxs-lookup"><span data-stu-id="c07bc-202">1.1</span></span>|
|[<span data-ttu-id="c07bc-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c07bc-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c07bc-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c07bc-204">Compose or Read</span></span>|
