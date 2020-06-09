---
title: Пространство имен Office — набор обязательных элементов 1.3
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,3.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f9eb0c47afa36622fac33286b19b3a2d8f6340c7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612005"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="cfeac-103">Office (набор требований для почтового ящика 1,3)</span><span class="sxs-lookup"><span data-stu-id="cfeac-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="cfeac-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="cfeac-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cfeac-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="cfeac-106">Requirements</span></span>

|<span data-ttu-id="cfeac-107">Требование</span><span class="sxs-lookup"><span data-stu-id="cfeac-107">Requirement</span></span>| <span data-ttu-id="cfeac-108">Значение</span><span class="sxs-lookup"><span data-stu-id="cfeac-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfeac-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cfeac-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cfeac-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-110">1.1</span></span>|
|[<span data-ttu-id="cfeac-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cfeac-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cfeac-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cfeac-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cfeac-113">Properties</span><span class="sxs-lookup"><span data-stu-id="cfeac-113">Properties</span></span>

| <span data-ttu-id="cfeac-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="cfeac-114">Property</span></span> | <span data-ttu-id="cfeac-115">Способов</span><span class="sxs-lookup"><span data-stu-id="cfeac-115">Modes</span></span> | <span data-ttu-id="cfeac-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cfeac-116">Return type</span></span> | <span data-ttu-id="cfeac-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cfeac-117">Minimum</span></span><br><span data-ttu-id="cfeac-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="cfeac-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cfeac-119">контекст</span><span class="sxs-lookup"><span data-stu-id="cfeac-119">context</span></span>](office.context.md) | <span data-ttu-id="cfeac-120">Создание</span><span class="sxs-lookup"><span data-stu-id="cfeac-120">Compose</span></span><br><span data-ttu-id="cfeac-121">Read</span><span class="sxs-lookup"><span data-stu-id="cfeac-121">Read</span></span> | [<span data-ttu-id="cfeac-122">Context</span><span class="sxs-lookup"><span data-stu-id="cfeac-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="cfeac-123">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cfeac-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="cfeac-124">Enumerations</span></span>

| <span data-ttu-id="cfeac-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="cfeac-125">Enumeration</span></span> | <span data-ttu-id="cfeac-126">Способов</span><span class="sxs-lookup"><span data-stu-id="cfeac-126">Modes</span></span> | <span data-ttu-id="cfeac-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="cfeac-127">Return type</span></span> | <span data-ttu-id="cfeac-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="cfeac-128">Minimum</span></span><br><span data-ttu-id="cfeac-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="cfeac-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cfeac-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cfeac-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cfeac-131">Создание</span><span class="sxs-lookup"><span data-stu-id="cfeac-131">Compose</span></span><br><span data-ttu-id="cfeac-132">Read</span><span class="sxs-lookup"><span data-stu-id="cfeac-132">Read</span></span> | <span data-ttu-id="cfeac-133">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-133">String</span></span> | [<span data-ttu-id="cfeac-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cfeac-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cfeac-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cfeac-136">Создание</span><span class="sxs-lookup"><span data-stu-id="cfeac-136">Compose</span></span><br><span data-ttu-id="cfeac-137">Read</span><span class="sxs-lookup"><span data-stu-id="cfeac-137">Read</span></span> | <span data-ttu-id="cfeac-138">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-138">String</span></span> | [<span data-ttu-id="cfeac-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cfeac-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cfeac-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cfeac-141">Создание</span><span class="sxs-lookup"><span data-stu-id="cfeac-141">Compose</span></span><br><span data-ttu-id="cfeac-142">Read</span><span class="sxs-lookup"><span data-stu-id="cfeac-142">Read</span></span> | <span data-ttu-id="cfeac-143">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-143">String</span></span> | [<span data-ttu-id="cfeac-144">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cfeac-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="cfeac-145">Namespaces</span></span>

<span data-ttu-id="cfeac-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="cfeac-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cfeac-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="cfeac-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cfeac-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="cfeac-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="cfeac-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="cfeac-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cfeac-150">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-150">Type</span></span>

*   <span data-ttu-id="cfeac-151">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cfeac-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cfeac-152">Properties:</span></span>

|<span data-ttu-id="cfeac-153">Имя</span><span class="sxs-lookup"><span data-stu-id="cfeac-153">Name</span></span>| <span data-ttu-id="cfeac-154">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-154">Type</span></span>| <span data-ttu-id="cfeac-155">Описание</span><span class="sxs-lookup"><span data-stu-id="cfeac-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cfeac-156">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-156">String</span></span>|<span data-ttu-id="cfeac-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="cfeac-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cfeac-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="cfeac-158">String</span></span>|<span data-ttu-id="cfeac-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="cfeac-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cfeac-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="cfeac-160">Requirements</span></span>

|<span data-ttu-id="cfeac-161">Требование</span><span class="sxs-lookup"><span data-stu-id="cfeac-161">Requirement</span></span>| <span data-ttu-id="cfeac-162">Значение</span><span class="sxs-lookup"><span data-stu-id="cfeac-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfeac-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cfeac-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cfeac-164">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-164">1.1</span></span>|
|[<span data-ttu-id="cfeac-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cfeac-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cfeac-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cfeac-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cfeac-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="cfeac-167">CoercionType: String</span></span>

<span data-ttu-id="cfeac-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cfeac-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cfeac-169">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-169">Type</span></span>

*   <span data-ttu-id="cfeac-170">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cfeac-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cfeac-171">Properties:</span></span>

|<span data-ttu-id="cfeac-172">Имя</span><span class="sxs-lookup"><span data-stu-id="cfeac-172">Name</span></span>| <span data-ttu-id="cfeac-173">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-173">Type</span></span>| <span data-ttu-id="cfeac-174">Описание</span><span class="sxs-lookup"><span data-stu-id="cfeac-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cfeac-175">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-175">String</span></span>|<span data-ttu-id="cfeac-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="cfeac-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cfeac-177">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-177">String</span></span>|<span data-ttu-id="cfeac-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="cfeac-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cfeac-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="cfeac-179">Requirements</span></span>

|<span data-ttu-id="cfeac-180">Требование</span><span class="sxs-lookup"><span data-stu-id="cfeac-180">Requirement</span></span>| <span data-ttu-id="cfeac-181">Значение</span><span class="sxs-lookup"><span data-stu-id="cfeac-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfeac-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cfeac-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cfeac-183">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-183">1.1</span></span>|
|[<span data-ttu-id="cfeac-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cfeac-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cfeac-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cfeac-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cfeac-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="cfeac-186">SourceProperty: String</span></span>

<span data-ttu-id="cfeac-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="cfeac-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cfeac-188">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-188">Type</span></span>

*   <span data-ttu-id="cfeac-189">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cfeac-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="cfeac-190">Properties:</span></span>

|<span data-ttu-id="cfeac-191">Имя</span><span class="sxs-lookup"><span data-stu-id="cfeac-191">Name</span></span>| <span data-ttu-id="cfeac-192">Тип</span><span class="sxs-lookup"><span data-stu-id="cfeac-192">Type</span></span>| <span data-ttu-id="cfeac-193">Описание</span><span class="sxs-lookup"><span data-stu-id="cfeac-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cfeac-194">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-194">String</span></span>|<span data-ttu-id="cfeac-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="cfeac-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cfeac-196">String</span><span class="sxs-lookup"><span data-stu-id="cfeac-196">String</span></span>|<span data-ttu-id="cfeac-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="cfeac-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cfeac-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="cfeac-198">Requirements</span></span>

|<span data-ttu-id="cfeac-199">Требование</span><span class="sxs-lookup"><span data-stu-id="cfeac-199">Requirement</span></span>| <span data-ttu-id="cfeac-200">Значение</span><span class="sxs-lookup"><span data-stu-id="cfeac-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="cfeac-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cfeac-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cfeac-202">1.1</span><span class="sxs-lookup"><span data-stu-id="cfeac-202">1.1</span></span>|
|[<span data-ttu-id="cfeac-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cfeac-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cfeac-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cfeac-204">Compose or Read</span></span>|
