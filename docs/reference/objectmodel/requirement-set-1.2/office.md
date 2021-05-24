---
title: Office пространства имен — набор требований 1.2
description: Office пространства имен, доступных для Outlook надстройки с помощью набора API почтовых ящиков 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 4cd15d77d1c5d9b95152f038f3421c5838bfb84f
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590409"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="aee73-103">Office (набор требований к почтовым ящикам 1.2)</span><span class="sxs-lookup"><span data-stu-id="aee73-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="aee73-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="aee73-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="aee73-106">Требования</span><span class="sxs-lookup"><span data-stu-id="aee73-106">Requirements</span></span>

|<span data-ttu-id="aee73-107">Требование</span><span class="sxs-lookup"><span data-stu-id="aee73-107">Requirement</span></span>| <span data-ttu-id="aee73-108">Значение</span><span class="sxs-lookup"><span data-stu-id="aee73-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="aee73-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aee73-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aee73-110">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-110">1.1</span></span>|
|[<span data-ttu-id="aee73-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aee73-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aee73-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="aee73-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="aee73-113">Properties</span></span>

| <span data-ttu-id="aee73-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="aee73-114">Property</span></span> | <span data-ttu-id="aee73-115">Режимы</span><span class="sxs-lookup"><span data-stu-id="aee73-115">Modes</span></span> | <span data-ttu-id="aee73-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="aee73-116">Return type</span></span> | <span data-ttu-id="aee73-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="aee73-117">Minimum</span></span><br><span data-ttu-id="aee73-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="aee73-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aee73-119">контекст</span><span class="sxs-lookup"><span data-stu-id="aee73-119">context</span></span>](office.context.md) | <span data-ttu-id="aee73-120">Создание</span><span class="sxs-lookup"><span data-stu-id="aee73-120">Compose</span></span><br><span data-ttu-id="aee73-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-121">Read</span></span> | [<span data-ttu-id="aee73-122">Context</span><span class="sxs-lookup"><span data-stu-id="aee73-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="aee73-123">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="aee73-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="aee73-124">Enumerations</span></span>

| <span data-ttu-id="aee73-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="aee73-125">Enumeration</span></span> | <span data-ttu-id="aee73-126">Режимы</span><span class="sxs-lookup"><span data-stu-id="aee73-126">Modes</span></span> | <span data-ttu-id="aee73-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="aee73-127">Return type</span></span> | <span data-ttu-id="aee73-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="aee73-128">Minimum</span></span><br><span data-ttu-id="aee73-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="aee73-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="aee73-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="aee73-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="aee73-131">Создание</span><span class="sxs-lookup"><span data-stu-id="aee73-131">Compose</span></span><br><span data-ttu-id="aee73-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-132">Read</span></span> | <span data-ttu-id="aee73-133">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-133">String</span></span> | [<span data-ttu-id="aee73-134">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aee73-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="aee73-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="aee73-136">Создание</span><span class="sxs-lookup"><span data-stu-id="aee73-136">Compose</span></span><br><span data-ttu-id="aee73-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-137">Read</span></span> | <span data-ttu-id="aee73-138">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-138">String</span></span> | [<span data-ttu-id="aee73-139">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="aee73-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="aee73-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="aee73-141">Создание</span><span class="sxs-lookup"><span data-stu-id="aee73-141">Compose</span></span><br><span data-ttu-id="aee73-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-142">Read</span></span> | <span data-ttu-id="aee73-143">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-143">String</span></span> | [<span data-ttu-id="aee73-144">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="aee73-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="aee73-145">Namespaces</span></span>

<span data-ttu-id="aee73-146">[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="aee73-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="aee73-147">Сведения о переумериях</span><span class="sxs-lookup"><span data-stu-id="aee73-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="aee73-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="aee73-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="aee73-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="aee73-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="aee73-150">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-150">Type</span></span>

*   <span data-ttu-id="aee73-151">String</span><span class="sxs-lookup"><span data-stu-id="aee73-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aee73-152">Свойства</span><span class="sxs-lookup"><span data-stu-id="aee73-152">Properties</span></span>

|<span data-ttu-id="aee73-153">Имя</span><span class="sxs-lookup"><span data-stu-id="aee73-153">Name</span></span>| <span data-ttu-id="aee73-154">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-154">Type</span></span>| <span data-ttu-id="aee73-155">Описание</span><span class="sxs-lookup"><span data-stu-id="aee73-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="aee73-156">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-156">String</span></span>|<span data-ttu-id="aee73-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="aee73-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="aee73-158">String</span><span class="sxs-lookup"><span data-stu-id="aee73-158">String</span></span>|<span data-ttu-id="aee73-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="aee73-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aee73-160">Требования</span><span class="sxs-lookup"><span data-stu-id="aee73-160">Requirements</span></span>

|<span data-ttu-id="aee73-161">Требование</span><span class="sxs-lookup"><span data-stu-id="aee73-161">Requirement</span></span>| <span data-ttu-id="aee73-162">Значение</span><span class="sxs-lookup"><span data-stu-id="aee73-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="aee73-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aee73-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aee73-164">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-164">1.1</span></span>|
|[<span data-ttu-id="aee73-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aee73-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aee73-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="aee73-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="aee73-167">CoercionType: String</span></span>

<span data-ttu-id="aee73-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="aee73-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aee73-169">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-169">Type</span></span>

*   <span data-ttu-id="aee73-170">String</span><span class="sxs-lookup"><span data-stu-id="aee73-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aee73-171">Свойства</span><span class="sxs-lookup"><span data-stu-id="aee73-171">Properties</span></span>

|<span data-ttu-id="aee73-172">Имя</span><span class="sxs-lookup"><span data-stu-id="aee73-172">Name</span></span>| <span data-ttu-id="aee73-173">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-173">Type</span></span>| <span data-ttu-id="aee73-174">Описание</span><span class="sxs-lookup"><span data-stu-id="aee73-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="aee73-175">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-175">String</span></span>|<span data-ttu-id="aee73-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="aee73-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="aee73-177">String</span><span class="sxs-lookup"><span data-stu-id="aee73-177">String</span></span>|<span data-ttu-id="aee73-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="aee73-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aee73-179">Требования</span><span class="sxs-lookup"><span data-stu-id="aee73-179">Requirements</span></span>

|<span data-ttu-id="aee73-180">Требование</span><span class="sxs-lookup"><span data-stu-id="aee73-180">Requirement</span></span>| <span data-ttu-id="aee73-181">Значение</span><span class="sxs-lookup"><span data-stu-id="aee73-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="aee73-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aee73-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aee73-183">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-183">1.1</span></span>|
|[<span data-ttu-id="aee73-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aee73-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aee73-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="aee73-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="aee73-186">SourceProperty: String</span></span>

<span data-ttu-id="aee73-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="aee73-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="aee73-188">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-188">Type</span></span>

*   <span data-ttu-id="aee73-189">String</span><span class="sxs-lookup"><span data-stu-id="aee73-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="aee73-190">Свойства</span><span class="sxs-lookup"><span data-stu-id="aee73-190">Properties</span></span>

|<span data-ttu-id="aee73-191">Имя</span><span class="sxs-lookup"><span data-stu-id="aee73-191">Name</span></span>| <span data-ttu-id="aee73-192">Тип</span><span class="sxs-lookup"><span data-stu-id="aee73-192">Type</span></span>| <span data-ttu-id="aee73-193">Описание</span><span class="sxs-lookup"><span data-stu-id="aee73-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="aee73-194">Строка</span><span class="sxs-lookup"><span data-stu-id="aee73-194">String</span></span>|<span data-ttu-id="aee73-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="aee73-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="aee73-196">String</span><span class="sxs-lookup"><span data-stu-id="aee73-196">String</span></span>|<span data-ttu-id="aee73-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="aee73-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aee73-198">Требования</span><span class="sxs-lookup"><span data-stu-id="aee73-198">Requirements</span></span>

|<span data-ttu-id="aee73-199">Требование</span><span class="sxs-lookup"><span data-stu-id="aee73-199">Requirement</span></span>| <span data-ttu-id="aee73-200">Значение</span><span class="sxs-lookup"><span data-stu-id="aee73-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="aee73-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aee73-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="aee73-202">1.1</span><span class="sxs-lookup"><span data-stu-id="aee73-202">1.1</span></span>|
|[<span data-ttu-id="aee73-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aee73-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="aee73-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aee73-204">Compose or Read</span></span>|
