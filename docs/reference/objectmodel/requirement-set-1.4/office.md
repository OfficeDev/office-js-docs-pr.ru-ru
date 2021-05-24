---
title: Office пространства имен — набор требований 1.4
description: Office, доступные для Outlook надстройки с использованием набора API API почтовых ящиков 1.4.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0221ab09048719317c131f0204e2fc60c4f8f7d4
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591025"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="00855-103">Office (набор требований к почтовым ящикам 1.4)</span><span class="sxs-lookup"><span data-stu-id="00855-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="00855-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="00855-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="00855-106">Требования</span><span class="sxs-lookup"><span data-stu-id="00855-106">Requirements</span></span>

|<span data-ttu-id="00855-107">Требование</span><span class="sxs-lookup"><span data-stu-id="00855-107">Requirement</span></span>| <span data-ttu-id="00855-108">Значение</span><span class="sxs-lookup"><span data-stu-id="00855-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="00855-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00855-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="00855-110">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-110">1.1</span></span>|
|[<span data-ttu-id="00855-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00855-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="00855-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00855-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="00855-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="00855-113">Properties</span></span>

| <span data-ttu-id="00855-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="00855-114">Property</span></span> | <span data-ttu-id="00855-115">Режимы</span><span class="sxs-lookup"><span data-stu-id="00855-115">Modes</span></span> | <span data-ttu-id="00855-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="00855-116">Return type</span></span> | <span data-ttu-id="00855-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="00855-117">Minimum</span></span><br><span data-ttu-id="00855-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="00855-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="00855-119">контекст</span><span class="sxs-lookup"><span data-stu-id="00855-119">context</span></span>](office.context.md) | <span data-ttu-id="00855-120">Создание</span><span class="sxs-lookup"><span data-stu-id="00855-120">Compose</span></span><br><span data-ttu-id="00855-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="00855-121">Read</span></span> | [<span data-ttu-id="00855-122">Context</span><span class="sxs-lookup"><span data-stu-id="00855-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="00855-123">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="00855-124">Перечисления</span><span class="sxs-lookup"><span data-stu-id="00855-124">Enumerations</span></span>

| <span data-ttu-id="00855-125">Перечисление</span><span class="sxs-lookup"><span data-stu-id="00855-125">Enumeration</span></span> | <span data-ttu-id="00855-126">Режимы</span><span class="sxs-lookup"><span data-stu-id="00855-126">Modes</span></span> | <span data-ttu-id="00855-127">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="00855-127">Return type</span></span> | <span data-ttu-id="00855-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="00855-128">Minimum</span></span><br><span data-ttu-id="00855-129">набор требований</span><span class="sxs-lookup"><span data-stu-id="00855-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="00855-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="00855-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="00855-131">Создание</span><span class="sxs-lookup"><span data-stu-id="00855-131">Compose</span></span><br><span data-ttu-id="00855-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="00855-132">Read</span></span> | <span data-ttu-id="00855-133">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-133">String</span></span> | [<span data-ttu-id="00855-134">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="00855-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="00855-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="00855-136">Создание</span><span class="sxs-lookup"><span data-stu-id="00855-136">Compose</span></span><br><span data-ttu-id="00855-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="00855-137">Read</span></span> | <span data-ttu-id="00855-138">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-138">String</span></span> | [<span data-ttu-id="00855-139">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="00855-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="00855-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="00855-141">Создание</span><span class="sxs-lookup"><span data-stu-id="00855-141">Compose</span></span><br><span data-ttu-id="00855-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="00855-142">Read</span></span> | <span data-ttu-id="00855-143">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-143">String</span></span> | [<span data-ttu-id="00855-144">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="00855-145">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="00855-145">Namespaces</span></span>

<span data-ttu-id="00855-146">[MailboxEnums:](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true)включает ряд Outlook определенных списков, например , , `ItemType` `EntityType` , `AttachmentType` , , , `RecipientType` и `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="00855-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="00855-147">Сведения о переумериях</span><span class="sxs-lookup"><span data-stu-id="00855-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="00855-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="00855-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="00855-149">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="00855-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="00855-150">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-150">Type</span></span>

*   <span data-ttu-id="00855-151">String</span><span class="sxs-lookup"><span data-stu-id="00855-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00855-152">Свойства</span><span class="sxs-lookup"><span data-stu-id="00855-152">Properties</span></span>

|<span data-ttu-id="00855-153">Имя</span><span class="sxs-lookup"><span data-stu-id="00855-153">Name</span></span>| <span data-ttu-id="00855-154">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-154">Type</span></span>| <span data-ttu-id="00855-155">Описание</span><span class="sxs-lookup"><span data-stu-id="00855-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="00855-156">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-156">String</span></span>|<span data-ttu-id="00855-157">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="00855-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="00855-158">String</span><span class="sxs-lookup"><span data-stu-id="00855-158">String</span></span>|<span data-ttu-id="00855-159">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="00855-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00855-160">Требования</span><span class="sxs-lookup"><span data-stu-id="00855-160">Requirements</span></span>

|<span data-ttu-id="00855-161">Требование</span><span class="sxs-lookup"><span data-stu-id="00855-161">Requirement</span></span>| <span data-ttu-id="00855-162">Значение</span><span class="sxs-lookup"><span data-stu-id="00855-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="00855-163">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00855-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="00855-164">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-164">1.1</span></span>|
|[<span data-ttu-id="00855-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00855-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="00855-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00855-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="00855-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="00855-167">CoercionType: String</span></span>

<span data-ttu-id="00855-168">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="00855-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="00855-169">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-169">Type</span></span>

*   <span data-ttu-id="00855-170">String</span><span class="sxs-lookup"><span data-stu-id="00855-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00855-171">Свойства</span><span class="sxs-lookup"><span data-stu-id="00855-171">Properties</span></span>

|<span data-ttu-id="00855-172">Имя</span><span class="sxs-lookup"><span data-stu-id="00855-172">Name</span></span>| <span data-ttu-id="00855-173">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-173">Type</span></span>| <span data-ttu-id="00855-174">Описание</span><span class="sxs-lookup"><span data-stu-id="00855-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="00855-175">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-175">String</span></span>|<span data-ttu-id="00855-176">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="00855-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="00855-177">String</span><span class="sxs-lookup"><span data-stu-id="00855-177">String</span></span>|<span data-ttu-id="00855-178">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="00855-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00855-179">Требования</span><span class="sxs-lookup"><span data-stu-id="00855-179">Requirements</span></span>

|<span data-ttu-id="00855-180">Требование</span><span class="sxs-lookup"><span data-stu-id="00855-180">Requirement</span></span>| <span data-ttu-id="00855-181">Значение</span><span class="sxs-lookup"><span data-stu-id="00855-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="00855-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00855-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="00855-183">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-183">1.1</span></span>|
|[<span data-ttu-id="00855-184">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00855-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="00855-185">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00855-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="00855-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="00855-186">SourceProperty: String</span></span>

<span data-ttu-id="00855-187">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="00855-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="00855-188">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-188">Type</span></span>

*   <span data-ttu-id="00855-189">String</span><span class="sxs-lookup"><span data-stu-id="00855-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="00855-190">Свойства</span><span class="sxs-lookup"><span data-stu-id="00855-190">Properties</span></span>

|<span data-ttu-id="00855-191">Имя</span><span class="sxs-lookup"><span data-stu-id="00855-191">Name</span></span>| <span data-ttu-id="00855-192">Тип</span><span class="sxs-lookup"><span data-stu-id="00855-192">Type</span></span>| <span data-ttu-id="00855-193">Описание</span><span class="sxs-lookup"><span data-stu-id="00855-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="00855-194">Строка</span><span class="sxs-lookup"><span data-stu-id="00855-194">String</span></span>|<span data-ttu-id="00855-195">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="00855-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="00855-196">String</span><span class="sxs-lookup"><span data-stu-id="00855-196">String</span></span>|<span data-ttu-id="00855-197">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="00855-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="00855-198">Требования</span><span class="sxs-lookup"><span data-stu-id="00855-198">Requirements</span></span>

|<span data-ttu-id="00855-199">Требование</span><span class="sxs-lookup"><span data-stu-id="00855-199">Requirement</span></span>| <span data-ttu-id="00855-200">Значение</span><span class="sxs-lookup"><span data-stu-id="00855-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="00855-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="00855-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="00855-202">1.1</span><span class="sxs-lookup"><span data-stu-id="00855-202">1.1</span></span>|
|[<span data-ttu-id="00855-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="00855-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="00855-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="00855-204">Compose or Read</span></span>|
