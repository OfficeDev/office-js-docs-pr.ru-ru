---
title: Пространство имен Office — набор обязательных элементов 1.3
description: Элементы пространства имен Office, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,3.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 93f92759e893a836a7bb35b743089acdafbebb5a
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430711"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="10acc-103">Office (набор требований для почтового ящика 1,3)</span><span class="sxs-lookup"><span data-stu-id="10acc-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="10acc-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="10acc-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="10acc-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="10acc-106">Requirements</span></span>

|<span data-ttu-id="10acc-107">Требование</span><span class="sxs-lookup"><span data-stu-id="10acc-107">Requirement</span></span>| <span data-ttu-id="10acc-108">Значение</span><span class="sxs-lookup"><span data-stu-id="10acc-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="10acc-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="10acc-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="10acc-110">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-110">1.1</span></span>|
|[<span data-ttu-id="10acc-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="10acc-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="10acc-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="10acc-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="10acc-113">Properties</span></span>

| <span data-ttu-id="10acc-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="10acc-114">Property</span></span> | <span data-ttu-id="10acc-115">Способов</span><span class="sxs-lookup"><span data-stu-id="10acc-115">Modes</span></span> | <span data-ttu-id="10acc-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="10acc-116">Return type</span></span> | <span data-ttu-id="10acc-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="10acc-117">Minimum</span></span><br><span data-ttu-id="10acc-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="10acc-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="10acc-119">контекст</span><span class="sxs-lookup"><span data-stu-id="10acc-119">context</span></span>](office.context.md) | <span data-ttu-id="10acc-120">Создание</span><span class="sxs-lookup"><span data-stu-id="10acc-120">Compose</span></span><br><span data-ttu-id="10acc-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-121">Read</span></span> | [<span data-ttu-id="10acc-122">Context</span><span class="sxs-lookup"><span data-stu-id="10acc-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="10acc-123">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="10acc-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="10acc-124">Enumerations</span></span>

| <span data-ttu-id="10acc-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="10acc-125">Enumeration</span></span> | <span data-ttu-id="10acc-126">Способов</span><span class="sxs-lookup"><span data-stu-id="10acc-126">Modes</span></span> | <span data-ttu-id="10acc-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="10acc-127">Return type</span></span> | <span data-ttu-id="10acc-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="10acc-128">Minimum</span></span><br><span data-ttu-id="10acc-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="10acc-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="10acc-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="10acc-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="10acc-131">Создание</span><span class="sxs-lookup"><span data-stu-id="10acc-131">Compose</span></span><br><span data-ttu-id="10acc-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-132">Read</span></span> | <span data-ttu-id="10acc-133">String</span><span class="sxs-lookup"><span data-stu-id="10acc-133">String</span></span> | [<span data-ttu-id="10acc-134">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="10acc-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="10acc-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="10acc-136">Создание</span><span class="sxs-lookup"><span data-stu-id="10acc-136">Compose</span></span><br><span data-ttu-id="10acc-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-137">Read</span></span> | <span data-ttu-id="10acc-138">String</span><span class="sxs-lookup"><span data-stu-id="10acc-138">String</span></span> | [<span data-ttu-id="10acc-139">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="10acc-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="10acc-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="10acc-141">Создание</span><span class="sxs-lookup"><span data-stu-id="10acc-141">Compose</span></span><br><span data-ttu-id="10acc-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-142">Read</span></span> | <span data-ttu-id="10acc-143">String</span><span class="sxs-lookup"><span data-stu-id="10acc-143">String</span></span> | [<span data-ttu-id="10acc-144">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="10acc-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="10acc-145">Namespaces</span></span>

<span data-ttu-id="10acc-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): включает ряд специфических перечислений Outlook, например,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` и `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="10acc-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="10acc-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="10acc-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="10acc-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="10acc-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="10acc-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="10acc-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="10acc-150">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-150">Type</span></span>

*   <span data-ttu-id="10acc-151">String</span><span class="sxs-lookup"><span data-stu-id="10acc-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="10acc-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="10acc-152">Properties:</span></span>

|<span data-ttu-id="10acc-153">Имя</span><span class="sxs-lookup"><span data-stu-id="10acc-153">Name</span></span>| <span data-ttu-id="10acc-154">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-154">Type</span></span>| <span data-ttu-id="10acc-155">Описание</span><span class="sxs-lookup"><span data-stu-id="10acc-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="10acc-156">String</span><span class="sxs-lookup"><span data-stu-id="10acc-156">String</span></span>|<span data-ttu-id="10acc-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="10acc-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="10acc-158">String</span><span class="sxs-lookup"><span data-stu-id="10acc-158">String</span></span>|<span data-ttu-id="10acc-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="10acc-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="10acc-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="10acc-160">Requirements</span></span>

|<span data-ttu-id="10acc-161">Требование</span><span class="sxs-lookup"><span data-stu-id="10acc-161">Requirement</span></span>| <span data-ttu-id="10acc-162">Значение</span><span class="sxs-lookup"><span data-stu-id="10acc-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="10acc-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="10acc-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="10acc-164">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-164">1.1</span></span>|
|[<span data-ttu-id="10acc-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="10acc-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="10acc-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="10acc-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="10acc-167">CoercionType: String</span></span>

<span data-ttu-id="10acc-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="10acc-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="10acc-169">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-169">Type</span></span>

*   <span data-ttu-id="10acc-170">String</span><span class="sxs-lookup"><span data-stu-id="10acc-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="10acc-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="10acc-171">Properties:</span></span>

|<span data-ttu-id="10acc-172">Имя</span><span class="sxs-lookup"><span data-stu-id="10acc-172">Name</span></span>| <span data-ttu-id="10acc-173">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-173">Type</span></span>| <span data-ttu-id="10acc-174">Описание</span><span class="sxs-lookup"><span data-stu-id="10acc-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="10acc-175">String</span><span class="sxs-lookup"><span data-stu-id="10acc-175">String</span></span>|<span data-ttu-id="10acc-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="10acc-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="10acc-177">String</span><span class="sxs-lookup"><span data-stu-id="10acc-177">String</span></span>|<span data-ttu-id="10acc-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="10acc-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="10acc-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="10acc-179">Requirements</span></span>

|<span data-ttu-id="10acc-180">Требование</span><span class="sxs-lookup"><span data-stu-id="10acc-180">Requirement</span></span>| <span data-ttu-id="10acc-181">Значение</span><span class="sxs-lookup"><span data-stu-id="10acc-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="10acc-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="10acc-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="10acc-183">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-183">1.1</span></span>|
|[<span data-ttu-id="10acc-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="10acc-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="10acc-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="10acc-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="10acc-186">SourceProperty: String</span></span>

<span data-ttu-id="10acc-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="10acc-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="10acc-188">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-188">Type</span></span>

*   <span data-ttu-id="10acc-189">String</span><span class="sxs-lookup"><span data-stu-id="10acc-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="10acc-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="10acc-190">Properties:</span></span>

|<span data-ttu-id="10acc-191">Имя</span><span class="sxs-lookup"><span data-stu-id="10acc-191">Name</span></span>| <span data-ttu-id="10acc-192">Тип</span><span class="sxs-lookup"><span data-stu-id="10acc-192">Type</span></span>| <span data-ttu-id="10acc-193">Описание</span><span class="sxs-lookup"><span data-stu-id="10acc-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="10acc-194">String</span><span class="sxs-lookup"><span data-stu-id="10acc-194">String</span></span>|<span data-ttu-id="10acc-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="10acc-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="10acc-196">String</span><span class="sxs-lookup"><span data-stu-id="10acc-196">String</span></span>|<span data-ttu-id="10acc-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="10acc-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="10acc-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="10acc-198">Requirements</span></span>

|<span data-ttu-id="10acc-199">Требование</span><span class="sxs-lookup"><span data-stu-id="10acc-199">Requirement</span></span>| <span data-ttu-id="10acc-200">Значение</span><span class="sxs-lookup"><span data-stu-id="10acc-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="10acc-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="10acc-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="10acc-202">1.1</span><span class="sxs-lookup"><span data-stu-id="10acc-202">1.1</span></span>|
|[<span data-ttu-id="10acc-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="10acc-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="10acc-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="10acc-204">Compose or Read</span></span>|
