---
title: Пространство имен Office — набор обязательных элементов 1,1
description: Объектная модель для пространства имен верхнего уровня API надстроек Outlook (версия API почтовых ящиков 1,1).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e881f0f9eac054f2b95436504da24cc7d4dec86d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720192"
---
# <a name="office"></a><span data-ttu-id="6ddf9-103">Office</span><span class="sxs-lookup"><span data-stu-id="6ddf9-103">Office</span></span>

<span data-ttu-id="6ddf9-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6ddf9-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6ddf9-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="6ddf9-106">Requirements</span></span>

|<span data-ttu-id="6ddf9-107">Требование</span><span class="sxs-lookup"><span data-stu-id="6ddf9-107">Requirement</span></span>| <span data-ttu-id="6ddf9-108">Значение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ddf9-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6ddf9-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6ddf9-110">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-110">1.1</span></span>|
|[<span data-ttu-id="6ddf9-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6ddf9-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6ddf9-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6ddf9-113">Properties</span><span class="sxs-lookup"><span data-stu-id="6ddf9-113">Properties</span></span>

| <span data-ttu-id="6ddf9-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="6ddf9-114">Property</span></span> | <span data-ttu-id="6ddf9-115">Способов</span><span class="sxs-lookup"><span data-stu-id="6ddf9-115">Modes</span></span> | <span data-ttu-id="6ddf9-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="6ddf9-116">Return type</span></span> | <span data-ttu-id="6ddf9-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="6ddf9-117">Minimum</span></span><br><span data-ttu-id="6ddf9-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="6ddf9-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6ddf9-119">контекст</span><span class="sxs-lookup"><span data-stu-id="6ddf9-119">context</span></span>](office.context.md) | <span data-ttu-id="6ddf9-120">Создание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-120">Compose</span></span><br><span data-ttu-id="6ddf9-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-121">Read</span></span> | [<span data-ttu-id="6ddf9-122">Context</span><span class="sxs-lookup"><span data-stu-id="6ddf9-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="6ddf9-123">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="6ddf9-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="6ddf9-124">Enumerations</span></span>

| <span data-ttu-id="6ddf9-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="6ddf9-125">Enumeration</span></span> | <span data-ttu-id="6ddf9-126">Способов</span><span class="sxs-lookup"><span data-stu-id="6ddf9-126">Modes</span></span> | <span data-ttu-id="6ddf9-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="6ddf9-127">Return type</span></span> | <span data-ttu-id="6ddf9-128">Минимальные</span><span class="sxs-lookup"><span data-stu-id="6ddf9-128">Minimum</span></span><br><span data-ttu-id="6ddf9-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="6ddf9-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6ddf9-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6ddf9-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6ddf9-131">Создание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-131">Compose</span></span><br><span data-ttu-id="6ddf9-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-132">Read</span></span> | <span data-ttu-id="6ddf9-133">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-133">String</span></span> | [<span data-ttu-id="6ddf9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6ddf9-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6ddf9-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6ddf9-136">Создание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-136">Compose</span></span><br><span data-ttu-id="6ddf9-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-137">Read</span></span> | <span data-ttu-id="6ddf9-138">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-138">String</span></span> | [<span data-ttu-id="6ddf9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6ddf9-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6ddf9-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6ddf9-141">Создание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-141">Compose</span></span><br><span data-ttu-id="6ddf9-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-142">Read</span></span> | <span data-ttu-id="6ddf9-143">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-143">String</span></span> | [<span data-ttu-id="6ddf9-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="6ddf9-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="6ddf9-145">Namespaces</span></span>

<span data-ttu-id="6ddf9-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): `ItemType`включает ряд специфических перечислений Outlook, например `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, и `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="6ddf9-147">Сведения о перечислении</span><span class="sxs-lookup"><span data-stu-id="6ddf9-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="6ddf9-148">AsyncResultStatus: строка</span><span class="sxs-lookup"><span data-stu-id="6ddf9-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="6ddf9-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6ddf9-150">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-150">Type</span></span>

*   <span data-ttu-id="6ddf9-151">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ddf9-152">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6ddf9-152">Properties:</span></span>

|<span data-ttu-id="6ddf9-153">Имя</span><span class="sxs-lookup"><span data-stu-id="6ddf9-153">Name</span></span>| <span data-ttu-id="6ddf9-154">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-154">Type</span></span>| <span data-ttu-id="6ddf9-155">Описание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6ddf9-156">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-156">String</span></span>|<span data-ttu-id="6ddf9-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6ddf9-158">Для указания</span><span class="sxs-lookup"><span data-stu-id="6ddf9-158">String</span></span>|<span data-ttu-id="6ddf9-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ddf9-160">Requirements</span><span class="sxs-lookup"><span data-stu-id="6ddf9-160">Requirements</span></span>

|<span data-ttu-id="6ddf9-161">Требование</span><span class="sxs-lookup"><span data-stu-id="6ddf9-161">Requirement</span></span>| <span data-ttu-id="6ddf9-162">Значение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ddf9-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6ddf9-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6ddf9-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-164">1.1</span></span>|
|[<span data-ttu-id="6ddf9-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6ddf9-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6ddf9-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="6ddf9-167">CoercionType: строка</span><span class="sxs-lookup"><span data-stu-id="6ddf9-167">CoercionType: String</span></span>

<span data-ttu-id="6ddf9-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6ddf9-169">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-169">Type</span></span>

*   <span data-ttu-id="6ddf9-170">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ddf9-171">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6ddf9-171">Properties:</span></span>

|<span data-ttu-id="6ddf9-172">Имя</span><span class="sxs-lookup"><span data-stu-id="6ddf9-172">Name</span></span>| <span data-ttu-id="6ddf9-173">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-173">Type</span></span>| <span data-ttu-id="6ddf9-174">Описание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6ddf9-175">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-175">String</span></span>|<span data-ttu-id="6ddf9-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6ddf9-177">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-177">String</span></span>|<span data-ttu-id="6ddf9-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ddf9-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="6ddf9-179">Requirements</span></span>

|<span data-ttu-id="6ddf9-180">Требование</span><span class="sxs-lookup"><span data-stu-id="6ddf9-180">Requirement</span></span>| <span data-ttu-id="6ddf9-181">Значение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ddf9-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6ddf9-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6ddf9-183">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-183">1.1</span></span>|
|[<span data-ttu-id="6ddf9-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6ddf9-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6ddf9-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="6ddf9-186">Перестрока: строка</span><span class="sxs-lookup"><span data-stu-id="6ddf9-186">SourceProperty: String</span></span>

<span data-ttu-id="6ddf9-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6ddf9-188">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-188">Type</span></span>

*   <span data-ttu-id="6ddf9-189">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6ddf9-190">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6ddf9-190">Properties:</span></span>

|<span data-ttu-id="6ddf9-191">Имя</span><span class="sxs-lookup"><span data-stu-id="6ddf9-191">Name</span></span>| <span data-ttu-id="6ddf9-192">Тип</span><span class="sxs-lookup"><span data-stu-id="6ddf9-192">Type</span></span>| <span data-ttu-id="6ddf9-193">Описание</span><span class="sxs-lookup"><span data-stu-id="6ddf9-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6ddf9-194">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-194">String</span></span>|<span data-ttu-id="6ddf9-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6ddf9-196">String</span><span class="sxs-lookup"><span data-stu-id="6ddf9-196">String</span></span>|<span data-ttu-id="6ddf9-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="6ddf9-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ddf9-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="6ddf9-198">Requirements</span></span>

|<span data-ttu-id="6ddf9-199">Требование</span><span class="sxs-lookup"><span data-stu-id="6ddf9-199">Requirement</span></span>| <span data-ttu-id="6ddf9-200">Значение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ddf9-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6ddf9-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6ddf9-202">1.1</span><span class="sxs-lookup"><span data-stu-id="6ddf9-202">1.1</span></span>|
|[<span data-ttu-id="6ddf9-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6ddf9-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6ddf9-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6ddf9-204">Compose or Read</span></span>|